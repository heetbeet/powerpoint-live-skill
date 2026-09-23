# Request reference

Send JSON through stdin or `request --file C:\absolute\request.json`:

```powershell
'{"op":"context"}' | python scripts/ppt_live.py request
python scripts/ppt_live.py request --file C:\absolute\request.json
```

The first request starts the bridge. `start`, `status`, and `stop` manage it. Responses contain `ok`, `result` or `error`, and `elapsed_ms`.

## Acquire state

```json
{"op":"context"}
{"op":"context","depth":1,"text_limit":240,"detail":true}
{"op":"deck"}
{"op":"deck","detail":true}
{"op":"selection"}
{"op":"outline","deck":"HANDLE_OR_ABSOLUTE_PATH"}
```

`context` combines the active presentation, canonical path, slide, selection, compact inspection, and revision. If shapes are selected, its revision is scoped to those shapes. `deck` is compact by default; `detail:true` includes each slide ID, index, and shape count. `outline` returns one title guess, shape count, and text character count per slide.

All presentation requests accept either the current bridge handle or the presentation's absolute path in `deck`. Handles are session-local. Paths are useful after a bridge restart.

## Inspect and check

```json
{"op":"inspect","deck":"HANDLE","slide_id":256}
{"op":"inspect","deck":"HANDLE","slide_id":256,"ids":[4,7],"depth":1,"text_limit":240,"detail":true}
{"op":"check","deck":"HANDLE","slide_id":256}
```

`inspect` returns the slide size, shapes, and revision. Shape records include ID, name, type, geometry box, rotation, z order, short text, and aggregate font data. `depth` accepts 0 through 12. `text_limit` accepts 0 through 20,000. `detail:true` adds text bounds, margins, autofit, and fill.

Supplying `ids` fingerprints only those shapes and their group ancestry. Its scoped revision can guard edits to those objects. Unrestricted native writes require a full slide revision.

## Batch edits

```json
{
  "op":"batch",
  "deck":"HANDLE",
  "slide_id":256,
  "expected_revision":"FROM_INSPECT",
  "operations":[
    {"op":"set","id":4,"x":72,"y":48,"w":700,"font_size":28},
    {"op":"replace_text","id":4,"old":"Old title","new":"New title"},
    {"op":"align","ids":[7,8,9],"mode":"top"},
    {"op":"distribute","ids":[7,8,9],"direction":"horizontal"}
  ]
}
```

Supported operations:

| Operation | Main fields |
| --- | --- |
| `set` | `id`; optional `x`, `y`, `w`, `h`, `text`, `font_size`, `font_name`, `bold`, `fill`, `color`, `name` |
| `replace_text` | `id`, exact case-sensitive `old`, `new` |
| `add_text`, `add_rect` | `x`, `y`, `w`, `h`; optional supported style fields and `name` |
| `add_image` | absolute existing `path`, `x`, `y`, `w`, `h`; optional `name` |
| `replace_image` | picture `id`, absolute existing `path` |
| `align` | distinct `ids`; `mode`: left, right, top, bottom, center, middle |
| `distribute` | distinct `ids`; `direction`: horizontal or vertical |
| `duplicate`, `delete` | `id` |

Colors use `#RRGGBB`. `replace_image` preserves a top-level picture's geometry, name, rotation, and approximate z position. Pictures nested in groups require a native script because replacing one can change group structure. A failed batch can have applied earlier operations or part of the failing operation. Inspect before submitting a new batch.

## Review and files

```json
{"op":"review","deck":"HANDLE","slide_ids":[256,257],"output_dir":"C:\\absolute\\new-review-dir","width":1280}
{"op":"render","deck":"HANDLE","slide_id":256,"path":"C:\\absolute\\slide.png","width":1280}
{"op":"render_shape","deck":"HANDLE","slide_id":256,"id":7,"path":"C:\\absolute\\shape.png","width":800}
{"op":"snapshot","deck":"HANDLE","path":"C:\\absolute\\snapshot.pptx"}
{"op":"save","deck":"HANDLE"}
```

`review` exports the requested slides into a new directory and returns compact recursive issues for shapes outside the canvas, measured text overflow, and text outside the canvas. Issue examples are capped and accompanied by total and truncated counts. Omit `slide_ids` to review the full deck. Image files stay on disk until opened.

Snapshot and render destinations must not exist. Use `.pptm` for a macro-enabled presentation snapshot.

Compare two saved packages without PowerPoint:

```powershell
python scripts/ppt_compare.py C:\absolute\source.pptx C:\absolute\candidate.pptx --image-hashes
```

Exit code 0 means the compared structure matches, 1 means drift, and 2 means invalid input.

## Recovery rules

A stale revision means the user or another operation changed the live state. Use the returned revision or current scoped state, inspect once if needed, and formulate a new edit. Do not replay the old mutation.

`partial_batch`, `native_partial`, and transport timeout responses can mean some changes already happened. Inspect first. The bridge automatically retries only read operations when PowerPoint reports that it is busy.
