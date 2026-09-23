#!/usr/bin/env python3
"""Read-only comparison of two PowerPoint packages."""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

PRESENTATION_PART = "ppt/presentation.xml"
PRESENTATION_RELS_PART = "ppt/_rels/presentation.xml.rels"
IMAGE_EXTENSIONS = {
    ".bmp",
    ".emf",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".svg",
    ".tif",
    ".tiff",
    ".webp",
    ".wmf",
}
SHAPE_TAGS = {
    f"{{{NS['p']}}}{name}": name
    for name in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart")
}
IGNORED_NOTE_PLACEHOLDERS = {"sldNum", "sldImg", "dt", "hdr", "ftr"}
DRIFT_LIMIT = 100


class InvalidInput(ValueError):
    """An input package or command line value cannot be compared."""


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise InvalidInput(message)


def read_xml(package: zipfile.ZipFile, part: str) -> ET.Element:
    try:
        data = package.read(part)
    except KeyError as exc:
        raise InvalidInput(f"PPTX package is missing {part}") from exc
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        raise InvalidInput(f"Could not read {part}: {exc}") from exc
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise InvalidInput(f"Invalid XML in {part}: {exc}") from exc


def relationships(package: zipfile.ZipFile, part: str) -> dict[str, dict[str, str]]:
    try:
        root = read_xml(package, part)
    except InvalidInput as exc:
        if str(exc) == f"PPTX package is missing {part}":
            return {}
        raise
    result = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.get("Id")
        if not rel_id:
            raise InvalidInput(f"Relationship without an Id in {part}")
        if rel_id in result:
            raise InvalidInput(f"Duplicate relationship Id {rel_id!r} in {part}")
        if rel.get("TargetMode") == "External":
            continue
        result[rel_id] = {
            "type": rel.get("Type", ""),
            "target": rel.get("Target", ""),
        }
    return result


def resolve_part(source_part: str, target: str) -> str:
    if not target:
        raise InvalidInput(f"Relationship from {source_part} has no target")
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))
    if resolved in ("", ".") or resolved == ".." or resolved.startswith("../"):
        raise InvalidInput(f"Relationship target escapes the package: {target}")
    return resolved


def ordered_slides(package: zipfile.ZipFile) -> list[str]:
    root = read_xml(package, PRESENTATION_PART)
    rels = relationships(package, PRESENTATION_RELS_PART)
    slide_list = root.find("p:sldIdLst", NS)
    if slide_list is None:
        raise InvalidInput("Presentation has no p:sldIdLst")

    parts = []
    for slide_id in slide_list.findall("p:sldId", NS):
        rel_id = slide_id.get(f"{{{NS['r']}}}id")
        if not rel_id:
            raise InvalidInput("Presentation slide has no relationship id")
        rel = rels.get(rel_id)
        if rel is None or not rel["type"].endswith("/slide"):
            raise InvalidInput(f"Presentation has an unresolved slide relationship: {rel_id}")
        parts.append(resolve_part(PRESENTATION_PART, rel["target"]))
    return parts


def paragraph_text(root: ET.Element) -> str:
    """Join runs within paragraphs so font-driven run splits do not drift."""
    paragraphs = []
    for paragraph in root.iter(f"{{{NS['a']}}}p"):
        value = "".join(node.text or "" for node in paragraph.iter(f"{{{NS['a']}}}t"))
        if value:
            paragraphs.append(value)
    return "\n".join(paragraphs)


def slide_text(package: zipfile.ZipFile, part: str) -> str:
    return paragraph_text(read_xml(package, part))


def slide_notes(package: zipfile.ZipFile, slide_part: str) -> str:
    rels_part = posixpath.join(
        posixpath.dirname(slide_part), "_rels", posixpath.basename(slide_part) + ".rels"
    )
    note_rel = next(
        (
            rel
            for rel in relationships(package, rels_part).values()
            if rel["type"].endswith("/notesSlide")
        ),
        None,
    )
    if note_rel is None:
        return ""

    root = read_xml(package, resolve_part(slide_part, note_rel["target"]))
    blocks = []
    for shape in root.findall(".//p:sp", NS):
        placeholder = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if placeholder is not None and placeholder.get("type", "") in IGNORED_NOTE_PLACEHOLDERS:
            continue
        text = paragraph_text(shape)
        if text:
            blocks.append(text)
    return "\n".join(blocks)


def visual_shapes(container: ET.Element, path: tuple[int, ...] = ()):
    for index, child in enumerate(list(container)):
        child_path = path + (index,)
        kind = SHAPE_TAGS.get(child.tag)
        if kind is not None:
            yield child, kind, child_path
            if kind == "grpSp":
                yield from visual_shapes(child, child_path)
        else:
            yield from visual_shapes(child, child_path)


def shape_identity(shape: ET.Element, path: tuple[int, ...]) -> dict[str, str]:
    nonvisual = shape.find(".//p:cNvPr", NS)
    return {
        "id": nonvisual.get("id", "") if nonvisual is not None else "",
        "name": nonvisual.get("name", "") if nonvisual is not None else "",
        "path": ".".join(str(value) for value in path),
    }


def shape_geometry(shape: ET.Element) -> dict[str, str] | None:
    xfrm = shape.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = shape.find("p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = shape.find("p:xfrm", NS)
    if xfrm is None:
        return None

    geometry = {f"xfrm_{key}": value for key, value in xfrm.attrib.items()}
    for tag, prefix, attrs in (
        ("off", "off", ("x", "y")),
        ("ext", "ext", ("cx", "cy")),
        ("chOff", "chOff", ("x", "y")),
        ("chExt", "chExt", ("cx", "cy")),
    ):
        element = xfrm.find(f"a:{tag}", NS)
        if element is not None:
            for attr in attrs:
                if attr in element.attrib:
                    geometry[f"{prefix}_{attr}"] = element.get(attr)
    return geometry


def slide_structure(package: zipfile.ZipFile, part: str):
    root = read_xml(package, part)
    tree = root.find("p:cSld/p:spTree", NS)
    if tree is None:
        return {}, collections.Counter()

    counts = collections.Counter()
    geometries = collections.Counter()
    for shape, kind, path in visual_shapes(tree):
        counts[kind] += 1
        signature = json.dumps(
            {"type": kind, "identity": shape_identity(shape, path),
             "geometry": shape_geometry(shape)},
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        geometries[signature] += 1
    return dict(sorted(counts.items())), geometries


def geometry_examples(counter: collections.Counter, limit: int = 3):
    examples = []
    for signature, count in sorted(counter.items()):
        item = json.loads(signature)
        item["count"] = count
        examples.append(item)
        if len(examples) == limit:
            break
    return examples


def image_hashes(package: zipfile.ZipFile) -> dict[str, str]:
    paths = sorted(
        name
        for name in package.namelist()
        if name.startswith("ppt/media/")
        and posixpath.splitext(name)[1].lower() in IMAGE_EXTENSIONS
    )
    return {name: hashlib.sha256(package.read(name)).hexdigest() for name in paths}


def picture_bindings(package: zipfile.ZipFile, slide_part: str) -> list[dict[str, object]]:
    root = read_xml(package, slide_part)
    tree = root.find("p:cSld/p:spTree", NS)
    if tree is None:
        return []
    rels_part = posixpath.join(
        posixpath.dirname(slide_part), "_rels", posixpath.basename(slide_part) + ".rels"
    )
    rels = relationships(package, rels_part)
    bindings = []
    for shape, kind, path in visual_shapes(tree):
        if kind != "pic":
            continue
        blip = shape.find(".//a:blip", NS)
        rel_id = blip.get(f"{{{NS['r']}}}embed") if blip is not None else None
        target = None
        digest = None
        if rel_id and rel_id in rels:
            target = resolve_part(slide_part, rels[rel_id]["target"])
            try:
                digest = hashlib.sha256(package.read(target)).hexdigest()
            except KeyError:
                digest = "missing"
        bindings.append({"identity": shape_identity(shape, path),
                         "target": target, "sha256": digest})
    return bindings


def image_hash_drift(source: dict[str, str], candidate: dict[str, str]) -> list[dict[str, str | None]]:
    drift = []
    for path in sorted(set(source) | set(candidate)):
        if source.get(path) != candidate.get(path):
            drift.append({"path": path, "source": source.get(path), "candidate": candidate.get(path)})
    return drift


def validate_archive(package: zipfile.ZipFile, label: str) -> None:
    names = [info.filename for info in package.infolist()]
    if len(names) != len(set(names)):
        raise InvalidInput(f"{label} package contains duplicate ZIP member names")


def compare(source_path: str, candidate_path: str, include_image_hashes: bool = False) -> dict:
    source_label = str(source_path)
    candidate_label = str(candidate_path)
    try:
        source_archive = zipfile.ZipFile(source_path, "r")
        candidate_archive = zipfile.ZipFile(candidate_path, "r")
    except (OSError, zipfile.BadZipFile) as exc:
        raise InvalidInput(f"Could not open package: {exc}") from exc

    with source_archive as source, candidate_archive as candidate:
        validate_archive(source, "Source")
        validate_archive(candidate, "Candidate")
        source_slides = ordered_slides(source)
        candidate_slides = ordered_slides(candidate)
        slide_total_match = len(source_slides) == len(candidate_slides)
        text_drift = []
        notes_drift = []
        type_count_drift = []
        geometry_drift = []
        picture_binding_drift = []
        drift_counts = collections.Counter()

        for index, (source_slide, candidate_slide) in enumerate(
            zip(source_slides, candidate_slides), start=1
        ):
            source_text = slide_text(source, source_slide)
            candidate_text = slide_text(candidate, candidate_slide)
            if source_text != candidate_text:
                drift_counts["text"] += 1
                if len(text_drift) < DRIFT_LIMIT:
                    text_drift.append({
                        "slide": index,
                        "source_chars": len(source_text),
                        "candidate_chars": len(candidate_text),
                    })

            source_notes = slide_notes(source, source_slide)
            candidate_notes = slide_notes(candidate, candidate_slide)
            if source_notes != candidate_notes:
                drift_counts["speaker_notes"] += 1
                if len(notes_drift) < DRIFT_LIMIT:
                    notes_drift.append({
                        "slide": index,
                        "source_chars": len(source_notes),
                        "candidate_chars": len(candidate_notes),
                    })

            source_counts, source_geometry = slide_structure(source, source_slide)
            candidate_counts, candidate_geometry = slide_structure(candidate, candidate_slide)
            if source_counts != candidate_counts:
                drift_counts["shape_type_count"] += 1
                if len(type_count_drift) < DRIFT_LIMIT:
                    type_count_drift.append({
                        "slide": index,
                        "source": source_counts,
                        "candidate": candidate_counts,
                    })

            source_only = source_geometry - candidate_geometry
            candidate_only = candidate_geometry - source_geometry
            if source_only or candidate_only:
                drift_counts["geometry"] += 1
                if len(geometry_drift) < DRIFT_LIMIT:
                    geometry_drift.append({
                        "slide": index,
                        "changed_shapes_estimate": max(
                            sum(source_only.values()), sum(candidate_only.values())
                        ),
                        "source_only_examples": geometry_examples(source_only),
                        "candidate_only_examples": geometry_examples(candidate_only),
                    })

            if include_image_hashes:
                source_bindings = picture_bindings(source, source_slide)
                candidate_bindings = picture_bindings(candidate, candidate_slide)
                if source_bindings != candidate_bindings:
                    drift_counts["picture_binding"] += 1
                    if len(picture_binding_drift) < DRIFT_LIMIT:
                        picture_binding_drift.append({
                            "slide": index,
                            "source": source_bindings[:5],
                            "candidate": candidate_bindings[:5],
                            "bindings_truncated": max(0, len(source_bindings)-5)
                            + max(0, len(candidate_bindings)-5),
                        })

        result = {
            "status": "pass",
            "source": source_label,
            "candidate": candidate_label,
            "slide_count": {
                "source": len(source_slides),
                "candidate": len(candidate_slides),
                "match": slide_total_match,
            },
            "text_drift": text_drift,
            "speaker_notes_drift": notes_drift,
            "shape_type_count_drift": type_count_drift,
            "geometry_drift": geometry_drift,
            "drift_counts": dict(sorted(drift_counts.items())),
        }

        if include_image_hashes:
            drift = image_hash_drift(image_hashes(source), image_hashes(candidate))
            drift_counts["image_hash"] = len(drift)
            result["image_hash_drift"] = drift[:DRIFT_LIMIT]
            result["picture_binding_drift"] = picture_binding_drift
            result["drift_counts"] = dict(sorted(drift_counts.items()))
        result["truncated_drift_examples"] = {
            "text": max(0, drift_counts["text"]-len(text_drift)),
            "speaker_notes": max(0, drift_counts["speaker_notes"]-len(notes_drift)),
            "shape_type_count": max(0, drift_counts["shape_type_count"]-len(type_count_drift)),
            "geometry": max(0, drift_counts["geometry"]-len(geometry_drift)),
            "picture_binding": max(0, drift_counts["picture_binding"]-len(picture_binding_drift)),
            "image_hash": max(0, drift_counts["image_hash"]-len(result.get("image_hash_drift", []))),
        }

        if (
            not slide_total_match
            or any(drift_counts.values())
        ):
            result["status"] = "drift"
        return result


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(
        description="Read-only PPTX text, notes, shape-count, geometry, and image comparison."
    )
    parser.add_argument("source", help="Path to the original PPTX")
    parser.add_argument("candidate", help="Path to the candidate PPTX")
    parser.add_argument(
        "--image-hashes",
        action="store_true",
        help="Compare embedded image bytes by ppt/media package path",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    try:
        args = build_parser().parse_args(argv)
        result = compare(args.source, args.candidate, args.image_hashes)
    except (InvalidInput, OSError, zipfile.BadZipFile) as exc:
        print(
            json.dumps(
                {"status": "error", "error": f"{type(exc).__name__}: {exc}"},
                separators=(",", ":"),
            )
        )
        return 2

    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    sys.exit(main())
