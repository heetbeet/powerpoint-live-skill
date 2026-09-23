# Native COM operations

The persistent bridge supports `native_read`, `native`, `native_deck`, `navigate`, `render_shape`, and `fit_text`. Use request files or stdin for scripts and Windows paths.

## Read native properties

`native_read` takes property paths relative to a shape, slide, presentation, or application. Known integer accessors such as `Runs(1)` are supported without `eval`. Each path succeeds or fails independently.

```json
{
  "op":"native_read",
  "deck":"HANDLE",
  "slide_id":256,
  "id":7,
  "paths":[
    "TextFrame2.MarginLeft",
    "TextFrame2.TextRange.Font.Size",
    "TextFrame2.TextRange.Runs(1).Text"
  ]
}
```

Set `target` to `slide`, `presentation`, or `application` when no shape ID is supplied. COM objects are summarized instead of recursively dumped.

## Run a slide script

`native` executes a UTF-8 Python file on the bridge's PowerPoint STA thread. Its namespace contains `app`, `presentation`, `slide`, `shape(id)`, and `result`.

```json
{"op":"native","mode":"read","deck":"HANDLE","slide_id":256,"path":"C:\\absolute\\inspect.py"}
```

For `mode:"write"`, supply a current full-slide `expected_revision`. The script can access unrestricted Python and COM. Set `result` to a small JSON-compatible value. An exception can follow partial changes.

Example grouping code:

```python
for index, shape_id in enumerate((101, 102)):
    shape(shape_id).Select() if index == 0 else shape(shape_id).Select(False)
group = app.ActiveWindow.Selection.ShapeRange.Group()
result = {"group_id": int(group.Id)}
```

Grouping and ungrouping can change collection membership and IDs. Inspect again before targeting the result.

## Run one cross-slide script

Use `native_deck` when one rule applies across several slides. Its namespace contains `app`, `presentation`, `slides`, `shape(slide_id, shape_id)`, and `result`.

Read selected slides:

```json
{
  "op":"native_deck",
  "mode":"read",
  "deck":"HANDLE_OR_ABSOLUTE_PATH",
  "slide_ids":[256,257],
  "path":"C:\\absolute\\audit.py"
}
```

For a write, provide a full-slide revision for every slide the script may change:

```json
{
  "op":"native_deck",
  "mode":"write",
  "deck":"HANDLE",
  "expected_revisions":{"256":"REVISION_A","257":"REVISION_B"},
  "path":"C:\\absolute\\edit-pass.py"
}
```

The `slides` mapping and `shape(slide_id, shape_id)` resolver expose only declared slides. The declared mode and revision map express intent; they are not a COM sandbox because `presentation` and `app` remain native objects. List and guard every slide the script may mutate. The bridge starts one undo entry and returns fresh revisions for declared slides. Any arbitrary script failure is reported as `native_partial`; inspect all declared slides before continuing.

A good script loops through the selected slides, applies one coherent rule, and returns counts plus a short exception list. Keep presentation-specific copy, coordinates, and design exceptions in the task script.

## Navigate, render, and fit text

`navigate` activates the resolved presentation window, moves to the slide's current index, and optionally selects `ids`.

`render_shape` exports one object or group to a fresh absolute `.png` path. Optional `width` accepts 1 through 4096 pixels.

`fit_text` takes `id`, `mode` (`shrink`, `grow`, or `none`), optional `wrap`, and `expected_revision`. It maps to `TextFrame2.AutoSize` and returns measured text bounds plus usable geometry before and after. PowerPoint can shrink text too far for comfortable reading, so inspect the font size and render when readability matters.

## Revision limits

Bridge revisions cover supported geometry, text, aggregate font settings, fill, and text frame properties. They do not cover every rich text run, chart cell, master, animation, or Office feature. For an uncovered feature, read and assert the relevant native value inside the script immediately before changing it.

## Microsoft references

- [TextFrame2.AutoSize](https://learn.microsoft.com/en-us/office/vba/api/office.textframe2.autosize) and [TextFrame2.WordWrap](https://learn.microsoft.com/en-us/office/vba/api/office.textframe2.wordwrap)
- [TextRange2 members](https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference/textrange2-members-office)
- [Shape.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export)
- [View.GotoSlide](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.view.gotoslide)
- [ShapeRange.Group](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.group) and [ShapeRange.Ungroup](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.ungroup)
