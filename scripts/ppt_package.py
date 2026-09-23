#!/usr/bin/env python3
"""Read selected slide XML from a PowerPoint package snapshot."""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SLIDE_REL_TYPE = R_NS + "/slide"
REL_TAG = "{" + PKG_REL_NS + "}Relationship"
SHAPE_TAGS = {
    "{" + P_NS + "}sp",
    "{" + P_NS + "}pic",
    "{" + P_NS + "}graphicFrame",
    "{" + P_NS + "}grpSp",
    "{" + P_NS + "}cxnSp",
}
CNVPR_TAG = "{" + P_NS + "}cNvPr"
REL_ATTRS = {"embed", "link", "id"}

MAX_SNAPSHOT_BYTES = 2 * 1024 * 1024 * 1024
MAX_ZIP_ENTRIES = 20_000
MAX_XML_PART_BYTES = 64 * 1024 * 1024
MAX_OUTPUT_CHARS = 1_000_000
MAX_OOXML_ID = (1 << 32) - 1


class InspectorError(Exception):
    """A problem that can be returned cleanly to a command-line caller."""


class UsageError(Exception):
    pass


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise UsageError(message)


def positive_id(value: str) -> int:
    try:
        number = int(value, 10)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a base-10 integer") from exc
    if number < 1 or number > MAX_OOXML_ID:
        raise argparse.ArgumentTypeError(f"must be between 1 and {MAX_OOXML_ID}")
    return number


def bounded_max_chars(value: str) -> int:
    try:
        number = int(value, 10)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a base-10 integer") from exc
    if number < 1 or number > MAX_OUTPUT_CHARS:
        raise argparse.ArgumentTypeError(f"must be between 1 and {MAX_OUTPUT_CHARS}")
    return number


def qname(namespace: str, local_name: str) -> str:
    return "{" + namespace + "}" + local_name


def rels_part_for(source_part: str) -> str:
    directory, filename = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", filename + ".rels")


def read_member(archive: zipfile.ZipFile, part_name: str) -> bytes:
    try:
        info = archive.getinfo(part_name)
    except KeyError as exc:
        raise InspectorError(f"package part is missing: {part_name}") from exc
    if info.file_size > MAX_XML_PART_BYTES:
        raise InspectorError(
            f"XML part exceeds the {MAX_XML_PART_BYTES}-byte limit: {part_name}"
        )
    try:
        with archive.open(info, "r") as stream:
            data = stream.read(MAX_XML_PART_BYTES + 1)
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        raise InspectorError(f"could not read package part {part_name}: {exc}") from exc
    if len(data) > MAX_XML_PART_BYTES:
        raise InspectorError(
            f"XML part exceeds the {MAX_XML_PART_BYTES}-byte limit: {part_name}"
        )
    return data


def parse_xml(data: bytes, part_name: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise InspectorError(f"invalid XML in {part_name}: {exc}") from exc


def relationship_map(root: ET.Element, rels_part: str) -> dict[str, ET.Element]:
    if root.tag != qname(PKG_REL_NS, "Relationships"):
        raise InspectorError(f"invalid relationships root in {rels_part}")
    result: dict[str, ET.Element] = {}
    for relation in root.findall(REL_TAG):
        rel_id = relation.get("Id")
        if not rel_id:
            raise InspectorError(f"relationship without an Id in {rels_part}")
        if rel_id in result:
            raise InspectorError(f"duplicate relationship Id {rel_id!r} in {rels_part}")
        result[rel_id] = relation
    return result


def resolve_internal_target(source_part: str, target: str) -> str:
    """Resolve a relationship URI reference to a normalized ZIP member path."""
    try:
        parsed = urlsplit(target)
    except ValueError as exc:
        raise InspectorError(f"invalid internal relationship target {target!r}") from exc
    if parsed.scheme or parsed.netloc:
        raise InspectorError(f"internal relationship target is not a package path: {target!r}")
    raw_path = parsed.path
    if "\\" in raw_path:
        raise InspectorError(f"backslash in package relationship target: {target!r}")
    if re.search(r"%(?![0-9A-Fa-f]{2})", raw_path):
        raise InspectorError(f"invalid percent escape in relationship target: {target!r}")

    absolute = raw_path.startswith("/")
    if absolute:
        raw_path = raw_path[1:]
    elif raw_path == "":
        raw_path = source_part
        absolute = True
    else:
        raw_path = posixpath.join(posixpath.dirname(source_part), raw_path)

    normalized_segments: list[str] = []
    for raw_segment in raw_path.split("/"):
        try:
            segment = unquote(raw_segment, encoding="utf-8", errors="strict")
        except UnicodeDecodeError as exc:
            raise InspectorError(f"invalid UTF-8 escape in relationship target: {target!r}") from exc
        if "/" in segment or "\\" in segment or "\x00" in segment:
            raise InspectorError(f"invalid escaped path separator in relationship target: {target!r}")
        if segment in ("", "."):
            if segment == "" and raw_segment != "" or segment == "" and raw_path not in ("",):
                raise InspectorError(f"empty path segment in relationship target: {target!r}")
            continue
        if segment == "..":
            if not normalized_segments:
                raise InspectorError(f"relationship target escapes the package root: {target!r}")
            normalized_segments.pop()
            continue
        normalized_segments.append(segment)

    if not normalized_segments:
        raise InspectorError(f"relationship target does not identify a package part: {target!r}")
    resolved = "/".join(normalized_segments)
    if not absolute and not resolved:
        raise InspectorError(f"invalid relationship target: {target!r}")
    return resolved


def find_slide_part(archive: zipfile.ZipFile, slide_id: int) -> str:
    presentation_part = "ppt/presentation.xml"
    presentation = parse_xml(read_member(archive, presentation_part), presentation_part)
    if presentation.tag != qname(P_NS, "presentation"):
        raise InspectorError(f"invalid presentation root in {presentation_part}")
    slide_list = presentation.find(qname(P_NS, "sldIdLst"))
    if slide_list is None:
        raise InspectorError("presentation.xml has no p:sldIdLst")

    matches = [
        item
        for item in slide_list.findall(qname(P_NS, "sldId"))
        if item.get("id", "").isdecimal() and int(item.get("id", "0")) == slide_id
    ]
    if not matches:
        raise InspectorError(f"stable slide id {slide_id} was not found in presentation.xml")
    if len(matches) != 1:
        raise InspectorError(f"stable slide id {slide_id} occurs more than once in presentation.xml")
    relation_id = matches[0].get(qname(R_NS, "id"))
    if not relation_id:
        raise InspectorError(f"p:sldId {slide_id} has no r:id")

    rels_part = rels_part_for(presentation_part)
    rels_root = parse_xml(read_member(archive, rels_part), rels_part)
    relations = relationship_map(rels_root, rels_part)
    relation = relations.get(relation_id)
    if relation is None:
        raise InspectorError(f"presentation relationship {relation_id!r} is missing")
    if relation.get("Type") != SLIDE_REL_TYPE:
        raise InspectorError(f"presentation relationship {relation_id!r} is not a slide relationship")
    if relation.get("TargetMode", "Internal") == "External":
        raise InspectorError(f"slide relationship {relation_id!r} has an external target")
    target = relation.get("Target")
    if not target:
        raise InspectorError(f"slide relationship {relation_id!r} has no target")
    return resolve_internal_target(presentation_part, target)


def own_shape_ids(container: ET.Element, parents: dict[ET.Element, ET.Element]) -> set[int]:
    ids: set[int] = set()
    for node in container.iter(CNVPR_TAG):
        ancestor = parents.get(node)
        nested_shape = False
        while ancestor is not None and ancestor is not container:
            if ancestor.tag in SHAPE_TAGS:
                nested_shape = True
                break
            ancestor = parents.get(ancestor)
        if nested_shape:
            continue
        raw_id = node.get("id", "")
        if raw_id.isdecimal():
            ids.add(int(raw_id))
    return ids


def find_shape_fragment(slide_root: ET.Element, shape_id: int) -> ET.Element:
    parents = {child: parent for parent in slide_root.iter() for child in parent}
    matches = [
        node
        for node in slide_root.iter()
        if node.tag in SHAPE_TAGS and shape_id in own_shape_ids(node, parents)
    ]
    if not matches:
        raise InspectorError(f"shape id {shape_id} was not found in the selected slide")
    if len(matches) != 1:
        raise InspectorError(f"shape id {shape_id} occurs more than once in the selected slide")
    return matches[0]


def referenced_relationship_ids(fragment: ET.Element) -> list[str]:
    refs: list[str] = []
    seen: set[str] = set()
    for element in fragment.iter():
        for attr_name, value in element.attrib.items():
            if not attr_name.startswith("{" + R_NS + "}"):
                continue
            local_name = attr_name.rsplit("}", 1)[1]
            if local_name in REL_ATTRS:
                if not value:
                    raise InspectorError(f"empty r:{local_name} relationship reference in fragment")
                if value not in seen:
                    seen.add(value)
                    refs.append(value)
    return refs


def relevant_relationships(
    archive: zipfile.ZipFile,
    slide_part: str,
    rel_ids: list[str],
) -> list[dict[str, object]]:
    if not rel_ids:
        return []
    rels_part = rels_part_for(slide_part)
    try:
        rels_data = read_member(archive, rels_part)
    except InspectorError as exc:
        if str(exc).startswith("package part is missing:"):
            return [{"id": rel_id, "missing": True} for rel_id in rel_ids]
        raise
    relations = relationship_map(parse_xml(rels_data, rels_part), rels_part)
    result: list[dict[str, object]] = []
    for rel_id in rel_ids:
        relation = relations.get(rel_id)
        if relation is None:
            result.append({"id": rel_id, "missing": True})
            continue
        target = relation.get("Target")
        record: dict[str, object] = {
            "id": rel_id,
            "type": relation.get("Type"),
            "target": target,
            "target_mode": relation.get("TargetMode", "Internal"),
        }
        mode = relation.get("TargetMode", "Internal")
        if mode == "External":
            record["external"] = True
        elif mode == "Internal":
            if target:
                try:
                    record["part"] = resolve_internal_target(slide_part, target)
                except InspectorError as exc:
                    record["resolution_error"] = str(exc)
            else:
                record["resolution_error"] = "relationship has no target"
        else:
            record["resolution_error"] = f"unsupported TargetMode {mode!r}"
        result.append(record)
    return result


def serialize_fragment(fragment: ET.Element) -> str:
    for prefix, namespace in (
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("p", P_NS),
        ("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main"),
        ("r", R_NS),
        ("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
    ):
        ET.register_namespace(prefix, namespace)
    return ET.tostring(fragment, encoding="unicode", short_empty_elements=True)


def build_parser() -> JsonArgumentParser:
    parser = JsonArgumentParser(
        description="Read a slide or shape XML fragment from an offline PPTX snapshot.",
    )
    parser.add_argument("snapshot", help="path to a PPTX/PPTM snapshot")
    parser.add_argument("--slide-id", required=True, type=positive_id, help="stable p:sldId/@id")
    parser.add_argument("--shape-id", type=positive_id, help="p:cNvPr/@id within the slide")
    parser.add_argument(
        "--max-chars",
        type=bounded_max_chars,
        default=6000,
        help=f"maximum XML characters returned in JSON (1..{MAX_OUTPUT_CHARS}; default: 6000)",
    )
    parser.add_argument("--output", help="write the full fragment to a new absolute XML path")
    return parser


def inspect(args: argparse.Namespace) -> dict[str, object]:
    snapshot = os.path.abspath(args.snapshot)
    snapshot_path = Path(snapshot)
    if not snapshot_path.is_file():
        raise InspectorError(f"snapshot is not a readable file: {snapshot}")
    try:
        size = snapshot_path.stat().st_size
    except OSError as exc:
        raise InspectorError(f"could not stat snapshot {snapshot}: {exc}") from exc
    if size > MAX_SNAPSHOT_BYTES:
        raise InspectorError(f"snapshot exceeds the {MAX_SNAPSHOT_BYTES}-byte limit")

    output_path: str | None = None
    if args.output is not None:
        if not Path(args.output).is_absolute():
            raise InspectorError("--output must be an absolute path")
        output_path = os.path.abspath(args.output)

    try:
        with zipfile.ZipFile(snapshot_path, "r") as archive:
            infos = archive.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                raise InspectorError(f"package exceeds the {MAX_ZIP_ENTRIES}-entry limit")
            names = [info.filename for info in infos]
            if len(set(names)) != len(names):
                raise InspectorError("package contains duplicate ZIP member names")

            slide_part = find_slide_part(archive, args.slide_id)
            slide_root = parse_xml(read_member(archive, slide_part), slide_part)
            if slide_root.tag != qname(P_NS, "sld"):
                raise InspectorError(f"invalid slide root in {slide_part}")
            fragment = (
                find_shape_fragment(slide_root, args.shape_id)
                if args.shape_id is not None
                else slide_root
            )
            full_xml = serialize_fragment(fragment)
            relationships = relevant_relationships(
                archive, slide_part, referenced_relationship_ids(fragment)
            )
    except (zipfile.BadZipFile, OSError, EOFError) as exc:
        raise InspectorError(f"could not read snapshot {snapshot}: {exc}") from exc

    result: dict[str, object] = {
        "path": snapshot,
        "slide_id": args.slide_id,
        "part": slide_part,
        "xml": full_xml[: args.max_chars],
        "truncated": len(full_xml) > args.max_chars,
        "relationships": relationships,
    }
    if args.shape_id is not None:
        result["shape_id"] = args.shape_id
    if output_path is not None:
        try:
            with open(output_path, "x", encoding="utf-8", newline="\n") as stream:
                stream.write('<?xml version="1.0" encoding="utf-8"?>\n')
                stream.write(full_xml)
                stream.write("\n")
        except FileExistsError as exc:
            raise InspectorError(f"refusing to overwrite existing output: {output_path}") from exc
        except OSError as exc:
            raise InspectorError(f"could not create output {output_path}: {exc}") from exc
        result["output"] = output_path
    return result


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
        result = inspect(args)
    except UsageError as exc:
        print(json.dumps({"error": str(exc)}, separators=(",", ":")), file=sys.stderr)
        return 2
    except InspectorError as exc:
        print(json.dumps({"error": str(exc)}, separators=(",", ":")), file=sys.stderr)
        return 1
    except (ValueError, OverflowError) as exc:
        print(json.dumps({"error": str(exc)}, separators=(",", ":")), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=True, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
