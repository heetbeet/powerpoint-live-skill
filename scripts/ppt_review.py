"""Read-only slide rendering and compact layout review helpers."""
from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any


class ReviewError(RuntimeError):
    def __init__(self, code: str, message: str, details: dict[str, Any] | None = None):
        super().__init__(message)
        self.code = code
        self.details = details


def _items(collection):
    return [collection.Item(i) for i in range(1, collection.Count + 1)]


def _round(value: Any) -> float:
    return round(float(value), 2)


def _text_bounds(shape):
    try:
        if not bool(shape.HasTextFrame):
            return None
        text_range = shape.TextFrame2.TextRange
        if not text_range.Text:
            return None
        return {
            "left": float(text_range.BoundLeft),
            "top": float(text_range.BoundTop),
            "width": float(text_range.BoundWidth),
            "height": float(text_range.BoundHeight),
            "margin_left": float(shape.TextFrame2.MarginLeft),
            "margin_top": float(shape.TextFrame2.MarginTop),
            "margin_right": float(shape.TextFrame2.MarginRight),
            "margin_bottom": float(shape.TextFrame2.MarginBottom),
        }
    except AttributeError:
        return None


def _outside_sides(left: float, top: float, width: float, height: float,
                   slide_width: float, slide_height: float) -> list[str]:
    sides = []
    if left < 0:
        sides.append("left")
    if top < 0:
        sides.append("top")
    if left + width > slide_width:
        sides.append("right")
    if top + height > slide_height:
        sides.append("bottom")
    return sides


def _inspect_shape(shape, slide_id: int, slide_width: float, slide_height: float) -> list[dict[str, Any]]:
    left = float(shape.Left)
    top = float(shape.Top)
    width = float(shape.Width)
    height = float(shape.Height)
    issues = []
    sides = _outside_sides(left, top, width, height, slide_width, slide_height)
    if sides:
        issues.append({
            "slide_id": slide_id,
            "shape_id": shape.Id,
            "kind": "outside_canvas",
            "sides": sides,
        })

    bounds = _text_bounds(shape)
    if bounds is not None:
        usable_width = max(0.0, width - bounds["margin_left"] - bounds["margin_right"])
        usable_height = max(0.0, height - bounds["margin_top"] - bounds["margin_bottom"])
        excess = [max(0.0, bounds["width"] - usable_width),
                  max(0.0, bounds["height"] - usable_height)]
        if max(excess) > 0.5:
            issues.append({
                "slide_id": slide_id,
                "shape_id": shape.Id,
                "kind": "text_overflow",
                "excess_wh": [_round(excess[0]), _round(excess[1])],
            })
        text_sides = _outside_sides(bounds["left"], bounds["top"], bounds["width"],
                                    bounds["height"], slide_width, slide_height)
        if text_sides:
            issues.append({
                "slide_id": slide_id,
                "shape_id": shape.Id,
                "kind": "text_outside_canvas",
                "sides": text_sides,
            })

    if shape.Type == 6:
        for child in _items(shape.GroupItems):
            issues.extend(_inspect_shape(child, slide_id, slide_width, slide_height))
    return issues


def review_slides(presentation, slides, request: dict[str, Any]) -> dict[str, Any]:
    output_value = request.get("output_dir")
    if not isinstance(output_value, str) or not Path(output_value).is_absolute():
        raise ReviewError("invalid_output_dir", "output_dir must be a new absolute directory")
    width = request.get("width", 1280)
    if isinstance(width, bool) or not isinstance(width, int) or not 64 <= width <= 8192:
        raise ReviewError("invalid_width", "width must be 64..8192 pixels")
    slide_width = float(presentation.PageSetup.SlideWidth)
    slide_height = float(presentation.PageSetup.SlideHeight)
    if slide_width <= 0 or slide_height <= 0:
        raise ReviewError("invalid_slide_size", "presentation slide size must be positive")

    output_dir = Path(output_value)
    if output_dir.exists():
        raise ReviewError("output_exists", "review output_dir must not already exist",
                          {"path": str(output_dir)})
    try:
        output_dir.mkdir()
    except OSError as exc:
        raise ReviewError("output_directory_error", "could not create review output_dir",
                          {"path": str(output_dir), "error": str(exc)})

    files = []
    deck_hash = hashlib.sha256()
    issues = []
    issue_count = 0
    issue_limit = 200
    for slide in slides:
        slide_id = int(slide.SlideID)
        height = round(width * slide_height / slide_width)
        destination = output_dir / f"slide-{slide_id}.png"
        slide.Export(str(destination), "PNG", width, height)
        if not destination.is_file() or destination.stat().st_size == 0:
            raise ReviewError("export_empty", "PowerPoint did not produce a slide PNG",
                              {"slide_id": slide_id, "path": str(destination)})
        image_hash = hashlib.sha256(destination.read_bytes()).hexdigest()[:20]
        deck_hash.update(str(slide_id).encode("ascii"))
        deck_hash.update(image_hash.encode("ascii"))
        files.append({"slide_id": slide_id, "path": str(destination),
                      "width": width, "height": height, "image_hash": image_hash})
        for shape in _items(slide.Shapes):
            found = _inspect_shape(shape, slide_id, slide_width, slide_height)
            issue_count += len(found)
            remaining = issue_limit - len(issues)
            if remaining > 0:
                issues.extend(found[:remaining])
    return {"output_dir": str(output_dir), "files": files, "issues": issues,
            "issue_count": issue_count, "issues_truncated": max(0, issue_count-len(issues)),
            "image_hash": deck_hash.hexdigest()[:20]}
