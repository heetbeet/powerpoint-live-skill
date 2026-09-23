from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "ppt_compare.py"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def package_parts(
    text: str = "Hello",
    notes: str = "Presenter note",
    x: int = 0,
    extra_shape: bool = False,
) -> dict[str, bytes]:
    extra = ""
    if extra_shape:
        extra = f"""
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="3" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="400" y="400"/><a:ext cx="200" cy="200"/></a:xfrm></p:spPr>
      </p:cxnSp>
"""
    slide = f"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="{A_NS}" xmlns:p="{P_NS}" xmlns:r="{R_NS}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{x}" y="100"/><a:ext cx="1000" cy="500"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>{text}</a:t></a:r></a:p></p:txBody>
    </p:sp>
{extra}  </p:spTree></p:cSld>
</p:sld>
""".encode()
    note = f"""<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="{A_NS}" xmlns:p="{P_NS}">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="1" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>ignored</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{notes}</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>
""".encode()
    return {
        "[Content_Types].xml": f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>
""".encode(),
        "ppt/presentation.xml": f"""<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>
""".encode(),
        "ppt/_rels/presentation.xml.rels": f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>
""".encode(),
        "ppt/slides/slide1.xml": slide,
        "ppt/slides/_rels/slide1.xml.rels": f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>
""".encode(),
        "ppt/notesSlides/notesSlide1.xml": note,
        "ppt/media/image1.png": b"synthetic image bytes",
    }


def write_package(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)


class CompareTests(unittest.TestCase):
    def test_identical_packages_pass_with_image_hashes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            candidate = Path(directory) / "candidate.pptx"
            parts = package_parts()
            write_package(source, parts)
            write_package(candidate, parts)

            completed = subprocess.run(
                [sys.executable, str(SCRIPT), str(source), str(candidate), "--image-hashes"],
                check=False,
                capture_output=True,
                text=True,
            )

            self.assertEqual(completed.returncode, 0)
            result = json.loads(completed.stdout)
            self.assertEqual(result["status"], "pass")
            self.assertEqual(result["slide_count"]["source"], 1)
            self.assertEqual(result["image_hash_drift"], [])

    def test_drift_reports_text_notes_shapes_geometry_and_image(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            candidate = Path(directory) / "candidate.pptx"
            source_parts = package_parts()
            candidate_parts = package_parts(
                text="Changed",
                notes="Changed note",
                x=25,
                extra_shape=True,
            )
            candidate_parts["ppt/media/image1.png"] = b"changed image bytes"
            write_package(source, source_parts)
            write_package(candidate, candidate_parts)

            completed = subprocess.run(
                [sys.executable, str(SCRIPT), str(source), str(candidate), "--image-hashes"],
                check=False,
                capture_output=True,
                text=True,
            )

            result = json.loads(completed.stdout)
            self.assertEqual(completed.returncode, 1)
            self.assertEqual(result["status"], "drift")
            self.assertEqual(result["text_drift"][0]["slide"], 1)
            self.assertEqual(result["speaker_notes_drift"][0]["slide"], 1)
            self.assertEqual(result["shape_type_count_drift"][0]["candidate"]["cxnSp"], 1)
            self.assertEqual(result["geometry_drift"][0]["slide"], 1)
            self.assertEqual(result["image_hash_drift"][0]["path"], "ppt/media/image1.png")

    def test_invalid_input_returns_json_and_exit_two(self) -> None:
        completed = subprocess.run(
            [sys.executable, str(SCRIPT), "missing-source.pptx", "missing-candidate.pptx"],
            check=False,
            capture_output=True,
            text=True,
        )

        self.assertEqual(completed.returncode, 2)
        self.assertEqual(json.loads(completed.stdout)["status"], "error")


if __name__ == "__main__":
    unittest.main()
