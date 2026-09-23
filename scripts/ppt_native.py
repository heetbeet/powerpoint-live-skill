"""Focused native PowerPoint COM operations for the persistent STA bridge."""

from __future__ import annotations

import math
import hashlib
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Callable


_IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*\Z")
_MAX_PATH_LENGTH = 512
_MAX_PATH_SEGMENTS = 32
_MAX_READ_PATHS = 96
_MAX_LIST_ITEMS = 32
_MAX_TEXT = 4096
_MAX_EXPORT_WIDTH = 4096
_AUTOSIZE = {"none": 0, "grow": 1, "shrink": 2}


class NativeOperationError(RuntimeError):
    """Bridge-facing error with compact structured details."""

    def __init__(self, code: str, message: str, details: Any = None):
        super().__init__(message)
        self.code = code
        self.details = details


def _fail(code: str, message: str, details: Any = None) -> None:
    raise NativeOperationError(code, message, details)


def _short_error(exc: BaseException) -> dict[str, str]:
    message = str(exc).strip() or type(exc).__name__
    return {"type": type(exc).__name__, "message": message[:_MAX_TEXT]}


def _compact(value: Any, depth: int = 0) -> Any:
    """Keep JSON results small and avoid walking arbitrary COM object models."""
    if depth > 8:
        return {"com_object": type(value).__name__, "truncated": True}
    if value is None or isinstance(value, (bool, int, str)):
        return value[:_MAX_TEXT] if isinstance(value, str) else value
    if isinstance(value, float):
        return value if math.isfinite(value) else str(value)
    if isinstance(value, (list, tuple)):
        items = [_compact(item, depth + 1) for item in value[:_MAX_LIST_ITEMS]]
        if len(value) <= _MAX_LIST_ITEMS:
            return items
        return {"kind": "list", "count": len(value), "items": items, "truncated": True}
    if isinstance(value, dict):
        keys = list(value)[:_MAX_LIST_ITEMS]
        result = {str(key)[:128]: _compact(value[key], depth + 1) for key in keys}
        if len(value) > _MAX_LIST_ITEMS:
            result["_truncated"] = len(value) - _MAX_LIST_ITEMS
        return result
    if isinstance(value, (bytes, bytearray, memoryview)):
        raw = bytes(value)
        return {"kind": "bytes", "length": len(raw), "preview_hex": raw[:32].hex()}
    return {"com_object": type(value).__name__}


def _parse_path(path: Any) -> list[tuple[str, int | None]]:
    if not isinstance(path, str) or not path or len(path) > _MAX_PATH_LENGTH:
        raise ValueError(f"path must be a non-empty string up to {_MAX_PATH_LENGTH} characters")
    result: list[tuple[str, int | None]] = []
    for segment in path.split("."):
        match = re.fullmatch(r"([A-Za-z_][A-Za-z0-9_]*)(?:\(([0-9]+)\))?", segment)
        if not match or match.group(1).startswith("__"):
            raise ValueError(f"invalid property path segment: {segment!r}")
        index = int(match.group(2)) if match.group(2) is not None else None
        if index is not None and match.group(1).lower() not in {
            'item','runs','paragraphs','characters','words','lines','sentences',
            'shapes','slides','groupitems','rows','columns','cells','seriescollection',
            'points','nodes','customlayouts','designs','sections','tabstops'
        }:
            raise ValueError('Use native write mode for method calls; this path accepts indexed reads only')
        result.append((match.group(1), index))
    if len(result) > _MAX_PATH_SEGMENTS:
        raise ValueError(f"path exceeds {_MAX_PATH_SEGMENTS} segments")
    return result


def _traverse(root: Any, path: str) -> Any:
    value = root
    for name, index in _parse_path(path):
        member = getattr(value, name)
        if index is None:
            value = member
            continue
        if callable(member):
            try:
                value = member(index)
                continue
            except Exception as call_error:
                try:
                    value = member.Item(index)
                    continue
                except Exception:
                    raise call_error
        item = getattr(member, "Item", None)
        if callable(item):
            value = item(index)
            continue
        try:
            value = member(index)
        except Exception as exc:
            raise TypeError(f"{name} does not support integer item access") from exc
    return value


def _shape_id(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError("id must be an integer shape ID")
    return value


def _resolve_shape(shape_resolver: Callable[[int], Any], value: Any) -> Any:
    return shape_resolver(_shape_id(value))


def _require_revision(request: dict[str, Any]) -> Any:
    revision = request.get("expected_revision")
    if not ((isinstance(revision, str) and revision) or
            (isinstance(revision, int) and not isinstance(revision, bool))):
        _fail("expected_revision_required", "expected_revision is required for this write")
    return revision


def _native_read(
    request: dict[str, Any],
    app: Any,
    presentation: Any,
    slide: Any,
    shape_resolver: Callable[[int], Any],
) -> dict[str, Any]:
    target = request.get('target')
    if target == 'application':
        root, root_name = app, 'application'
    elif target == 'presentation':
        root, root_name = presentation, 'presentation'
    elif "id" in request:
        root = _resolve_shape(shape_resolver, request["id"])
        root_name = "shape"
    elif slide is not None:
        root, root_name = slide, "slide"
    elif presentation is not None:
        root, root_name = presentation, "presentation"
    elif app is not None:
        root, root_name = app, "application"
    else:
        _fail("native_context_missing", "native_read requires a resolved COM context")

    paths = request.get("paths")
    if not isinstance(paths, list) or len(paths) > _MAX_READ_PATHS:
        _fail("invalid_paths", f"paths must be an array of at most {_MAX_READ_PATHS} strings")
    values = []
    for path in paths:
        try:
            values.append({"path": path, "value": _compact(_traverse(root, path))})
        except Exception as exc:
            values.append({"path": path, "error": _short_error(exc)})
    return {"root": root_name, "id": request.get("id"), "values": values}


def _native(
    request: dict[str, Any],
    app: Any,
    presentation: Any,
    slide: Any,
    shape_resolver: Callable[[int], Any],
) -> dict[str, Any]:
    mode = request.get("mode")
    if mode not in ("read", "write"):
        _fail("invalid_mode", "native mode must be explicitly 'read' or 'write'")
    revision = _require_revision(request) if mode == "write" else None
    source = request.get("path")
    if not isinstance(source, str) or not Path(source).is_absolute():
        _fail("invalid_script_path", "native path must be an absolute local script path")
    script_path = Path(source)
    try:
        code = script_path.read_text(encoding="utf-8-sig")
    except Exception as exc:
        _fail("native_script_read_error", "could not read native script", _short_error(exc))

    def shape(shape_id: int) -> Any:
        return _resolve_shape(shape_resolver, shape_id)

    namespace = {
        "app": app,
        "presentation": presentation,
        "slide": slide,
        "shape": shape,
        "result": None,
        "__file__": str(script_path),
        "__name__": "__ppt_native_request__",
    }
    try:
        exec(compile(code, str(script_path), "exec"), namespace, namespace)
    except Exception as exc:
        _fail(
            "native_script_error",
            "native script raised an exception; COM changes may already have occurred",
            {
                "mode": mode,
                "expected_revision": revision,
                "exception": _short_error(exc),
                "partial_changes_possible": True,
            },
        )
    return {
        "mode": mode,
        "path": str(script_path),
        "result": _compact(namespace.get("result")),
        "note": "Direct Python/COM execution is unrestricted under the bridge process privileges.",
    }


def _native_deck(
    request: dict[str, Any],
    app: Any,
    presentation: Any,
    slides: dict[int, Any],
    deck_shape_resolver: Callable[[int, int], Any],
) -> dict[str, Any]:
    mode = request.get("mode")
    if mode not in ("read", "write"):
        _fail("invalid_mode", "native_deck mode must be explicitly 'read' or 'write'")
    source = request.get("path")
    if not isinstance(source, str) or not Path(source).is_absolute():
        _fail("invalid_script_path", "native_deck path must be an absolute local script path")
    script_path = Path(source)
    try:
        code = script_path.read_text(encoding="utf-8-sig")
    except Exception as exc:
        _fail("native_script_read_error", "could not read native_deck script", _short_error(exc))

    def shape(slide_id: int, shape_id: int) -> Any:
        if isinstance(slide_id, bool) or not isinstance(slide_id, int):
            _fail("invalid_slide_id", "shape slide_id must be an integer SlideID")
        if isinstance(shape_id, bool) or not isinstance(shape_id, int):
            _fail("invalid_shape_id", "shape shape_id must be an integer shape ID")
        return deck_shape_resolver(slide_id, shape_id)

    namespace = {
        "app": app,
        "presentation": presentation,
        "slides": slides,
        "shape": shape,
        "result": None,
        "__file__": str(script_path),
        "__name__": "__ppt_native_deck_request__",
    }
    try:
        exec(compile(code, str(script_path), "exec"), namespace, namespace)
    except Exception as exc:
        _fail(
            "native_script_error",
            "native_deck script raised an exception; COM changes may already have occurred",
            {
                "mode": mode,
                "exception": _short_error(exc),
                "partial_changes_possible": mode == "write",
            },
        )
    return {
        "mode": mode,
        "path": str(script_path),
        "result": _compact(namespace.get("result")),
    }


def _bounds_and_geometry(shape: Any) -> dict[str, Any]:
    frame = shape.TextFrame2
    text_range = frame.TextRange
    margins = {
        "left": float(frame.MarginLeft),
        "top": float(frame.MarginTop),
        "right": float(frame.MarginRight),
        "bottom": float(frame.MarginBottom),
    }
    width, height = float(shape.Width), float(shape.Height)
    return {
        "text_bounds": {
            "left": float(text_range.BoundLeft),
            "top": float(text_range.BoundTop),
            "width": float(text_range.BoundWidth),
            "height": float(text_range.BoundHeight),
        },
        "shape": {
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": width,
            "height": height,
            "rotation": float(shape.Rotation),
        },
        "usable_geometry": {
            "width": max(0.0, width - margins["left"] - margins["right"]),
            "height": max(0.0, height - margins["top"] - margins["bottom"]),
            "margins": margins,
        },
    }


def _fit_text(request: dict[str, Any], shape_resolver: Callable[[int], Any]) -> dict[str, Any]:
    _require_revision(request)
    shape_id = _shape_id(request.get("id"))
    mode = request.get("mode", "shrink")
    if mode not in _AUTOSIZE:
        _fail("invalid_fit_mode", "fit_text mode must be 'shrink', 'grow', or 'none'")
    wrap = request.get("wrap")
    if wrap is not None and not isinstance(wrap, bool):
        _fail("invalid_wrap", "wrap must be a JSON boolean when supplied")
    shape = _resolve_shape(shape_resolver, shape_id)
    try:
        before = _bounds_and_geometry(shape)
        frame = shape.TextFrame2
        if wrap is not None:
            frame.WordWrap = -1 if wrap else 0
        frame.AutoSize = _AUTOSIZE[mode]
        after = _bounds_and_geometry(shape)
    except Exception as exc:
        _fail(
            "fit_text_error",
            "fit_text failed; the shape may have been partially changed",
            {"id": shape_id, "mode": mode, "exception": _short_error(exc),
             "partial_changes_possible": True},
        )
    return {
        "id": shape_id,
        "mode": mode,
        "wrap": wrap,
        "before": before,
        "after": after,
        "caveat": "Bounds and usable geometry are axis-aligned hints; rotated or complex text may need visual inspection.",
    }


def _navigate(
    request: dict[str, Any],
    app: Any,
    presentation: Any,
    slide: Any,
    shape_resolver: Callable[[int], Any],
) -> dict[str, Any]:
    if presentation is None or slide is None:
        _fail("native_context_missing", "navigate requires a resolved presentation and slide")
    try:
        windows = presentation.Windows
        if int(windows.Count) < 1:
            _fail("presentation_window_missing", "target presentation has no open window")
        window = windows.Item(1)
        window.Activate()
        slide_index = int(slide.SlideIndex)
        window.View.GotoSlide(slide_index)
    except NativeOperationError:
        raise
    except Exception as exc:
        _fail("navigate_error", "could not activate or navigate the target presentation", _short_error(exc))

    ids = request.get("ids", [])
    if not isinstance(ids, list):
        _fail("invalid_ids", "navigate ids must be an array of integer shape IDs")
    selected: list[int] = []
    try:
        for value in ids:
            shape_id = _shape_id(value)
            target = _resolve_shape(shape_resolver, shape_id)
            target.Select() if not selected else target.Select(False)
            selected.append(shape_id)
    except Exception as exc:
        _fail(
            "navigate_selection_error",
            "slide was activated, but shape selection may be incomplete",
            {"selected_ids": selected, "exception": _short_error(exc)},
        )
    return {"slide_index": slide_index, "selected_ids": selected, "presentation_activated": True}


def _render_shape(
    request: dict[str, Any],
    presentation: Any,
    shape_resolver: Callable[[int], Any],
) -> dict[str, Any]:
    shape_id = _shape_id(request.get("id"))
    output = request.get("path")
    if not isinstance(output, str) or not Path(output).is_absolute():
        _fail("invalid_output_path", "render_shape path must be an absolute local PNG path")
    destination = Path(output)
    if destination.suffix.lower() != ".png":
        _fail("invalid_output_path", "render_shape path must end in .png")
    if not destination.parent.is_dir():
        _fail("output_directory_missing", "render_shape parent directory must already exist")

    requested_width = request.get("width")
    export_args: tuple[Any, ...] = (str(destination), 2)
    if requested_width is not None:
        if (isinstance(requested_width, bool) or not isinstance(requested_width, int)
                or not 1 <= requested_width <= _MAX_EXPORT_WIDTH):
            _fail("invalid_width", f"width must be an integer from 1 to {_MAX_EXPORT_WIDTH} pixels")
        shape = _resolve_shape(shape_resolver, shape_id)
        try:
            slide_width = float(presentation.PageSetup.SlideWidth)
            slide_height = float(presentation.PageSetup.SlideHeight)
            shape_width = float(shape.Width)
            if slide_width <= 0 or slide_height <= 0 or shape_width <= 0:
                raise ValueError("slide and shape dimensions must be positive")
            # Shape.Export scales relative to slide points. Its default 96 DPI
            # makes shape_width * 4/3 the unscaled output width in pixels.
            factor = requested_width / (shape_width * 96.0 / 72.0)
            scale_width = round(slide_width * factor)
            scale_height = round(slide_height * factor)
            if not (1 <= scale_width <= 2_147_483_647 and 1 <= scale_height <= 2_147_483_647):
                raise ValueError("requested width produces unsupported scale dimensions")
            export_args = (str(destination), 2, scale_width, scale_height, 1)
        except Exception as exc:
            _fail("render_dimensions_error", "could not compute bounded PNG export dimensions", _short_error(exc))
    else:
        shape = _resolve_shape(shape_resolver, shape_id)

    fd, temporary_name = tempfile.mkstemp(prefix=".ppt-shape-", suffix=".png", dir=destination.parent)
    os.close(fd)
    temporary = Path(temporary_name)
    created_destination = False
    try:
        args = (str(temporary), *export_args[1:])
        shape.Export(*args)
        if not temporary.is_file() or temporary.stat().st_size == 0:
            _fail("shape_export_empty", "PowerPoint did not produce a PNG file")
        with temporary.open("rb") as source:
            with destination.open("xb") as target:
                created_destination = True
                shutil.copyfileobj(source, target)
        return {
            "id": shape_id,
            "path": str(destination),
            "width_px": requested_width,
            "bytes": destination.stat().st_size,
            "image_hash": hashlib.sha256(destination.read_bytes()).hexdigest()[:20],
        }
    except FileExistsError:
        _fail("output_exists", "render_shape refuses to overwrite an existing file", {"path": str(destination)})
    except NativeOperationError:
        raise
    except Exception as exc:
        if created_destination:
            try:
                destination.unlink()
            except OSError:
                pass
        _fail("shape_export_error", "could not export the target shape to PNG", _short_error(exc))
    finally:
        try:
            temporary.unlink()
        except OSError:
            pass


def handle(
    request: dict[str, Any],
    app: Any,
    presentation: Any,
    slide: Any,
    shape_resolver: Callable[[int], Any],
    *,
    slides: dict[int, Any] | None = None,
    deck_shape_resolver: Callable[[int, int], Any] | None = None,
) -> dict[str, Any] | None:
    """Handle native operations using bridge-resolved COM context; None means unhandled."""
    operation = request.get("op")
    if operation == "native_read":
        return _native_read(request, app, presentation, slide, shape_resolver)
    if operation == "native":
        return _native(request, app, presentation, slide, shape_resolver)
    if operation == "native_deck":
        if slides is None or deck_shape_resolver is None:
            _fail("native_context_missing", "native_deck requires a resolved presentation deck")
        return _native_deck(request, app, presentation, slides, deck_shape_resolver)
    if operation == "navigate":
        return _navigate(request, app, presentation, slide, shape_resolver)
    if operation == "render_shape":
        return _render_shape(request, presentation, shape_resolver)
    if operation == "fit_text":
        return _fit_text(request, shape_resolver)
    return None
