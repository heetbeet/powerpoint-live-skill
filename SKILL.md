---
name: powerpoint-live
description: Inspect and edit presentations live in Windows desktop PowerPoint through a persistent COM bridge. Use for native object access, slide and group navigation, text fitting, alignment, formatting, object images, and XML snapshots when fast feedback matters.
---

# PowerPoint live

Treat the open PowerPoint window as a live document. A persistent Python process owns one COM connection and accepts compact JSON requests through a local named pipe. PowerPoint must already be running.

## Fast loop

Run from this skill directory or use absolute script paths:

```powershell
'{"op":"context"}' | python scripts/ppt_live.py request
```

The first request starts the bridge. Keep the returned deck reference, slide ID, shape IDs, and revision in task state. Prefer the deck's canonical path when a handle may outlive one bridge session. Reacquire state after PowerPoint closes or the bridge restarts.

Use this loop while the user co-edits:

1. `context` for the active deck, slide, selection, and compact slide state.
2. `outline` only when the task spans several slides.
3. One targeted `inspect` for the objects involved. Expand groups or request detail only when needed.
4. One `batch` for related edits, guarded by the returned revision.
5. `check` for cheap geometry and text-bound hints.
6. `review` or `render_shape` only when appearance needs judgment.

Do not request a full deck dump after each change. A manual user edit can make a revision stale. On a stale error, use the returned current state or inspect once, adapt the change, and submit one new request.

## Edit safely

Work in PowerPoint points. Existing `x`, `y`, `w`, and `h` are shape geometry, not CSS rules. Use object IDs rather than names or screenshot positions. Shape IDs are local to a slide; slide IDs are stable across reordering.

Bundle related changes in `batch`. Prefer `replace_text` for an exact replacement within formatted text. Assigning all `text` can flatten formatting. Use `add_image` and `replace_image` for pictures without rebuilding the surrounding layout.

A batch or native write is not transactional. An error can follow partial changes. Never replay a timed-out or `native_partial` mutation. Inspect the affected scope first. Use `snapshot` before broad changes when a rollback copy is useful.

For cross-slide passes or unsupported object operations, run one short local Python file through `native_deck`. It executes inside the existing PowerPoint COM session and can use the actual presentation, slides, shapes, groups, and text objects. Keep deck-specific wording, visual exceptions, and coordinates in the task script. Return a compact summary instead of native object dumps.

Read [references/commands.md](references/commands.md) for request fields and [references/native.md](references/native.md) for direct COM access.

## Judge layout cheaply

Start with native geometry and measured text bounds. `check` flags possible overflow, off-slide shapes, and overlaps. These are hints: backgrounds, masks, grouped diagrams, rotations, effects, and intentional overlaps need visual judgment.

Use `review` for changed slides and `render_shape` for a local object. Open only the image needed for the next decision. Avoid injecting several full-resolution slide images when one changed-slide review or targeted crop will do.

Use native PowerPoint objects for alignment, distribution, margins, paragraph settings, text fitting, and group navigation. Autofit can make text technically fit while harming readability. Inspect the resulting font size and render when that tradeoff matters. Read [references/design.md](references/design.md) for known limits.

## Save, XML, and comparison

Edits appear immediately in the open document. Use `save` when saving is requested or needed to hand off the current document. Use `snapshot` for a separate copy. Do not set PowerPoint's `Saved` flag to suppress prompts.

Use `xml` when the package representation answers a question more cheaply or precisely than several COM reads. It makes a fresh `SaveCopyAs` snapshot and returns a bounded slide or shape fragment. Never edit the package beneath an open presentation. Read [references/xml.md](references/xml.md).

Use `scripts/ppt_compare.py` to check a generated copy against a source deck for text, notes, slide count, shape count, geometry, and optional image hashes. It is useful before replacing or delivering a deck.

## Keep token use low

Inspection text, tool output, and loaded images all consume context. Spend that context on slide decisions:

- Keep IDs, revisions, and selected facts in working state.
- Use `context`, scoped `inspect`, and compact native result dictionaries.
- Use one `native_deck` pass instead of many one-slide scripts for a consistent rule.
- Return file paths for renders and snapshots. Open only the files needed.
- Ask for `metrics` after a substantial phase, not after every edit.

The bridge records timing, payload sizes, approximate text tokens, repeated reads, large replies, failures, snapshots, and mutations. These are local workflow estimates, not billed model tokens. Read [references/efficiency.md](references/efficiency.md) for the measured evidence and interpretation.

If PowerPoint is busy or a modal dialog blocks COM, resolve that condition before retrying. Do not start another PowerPoint instance, close the user's presentation, or overwrite it as a connection workaround.
