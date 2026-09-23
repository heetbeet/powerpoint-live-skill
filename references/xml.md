# Offline XML inspection

## Open deck boundary

Use PowerPoint COM for edits to the open presentation. To save a fresh copy and read a selected XML fragment in one request:

```json
{"op":"xml","deck":"HANDLE","slide_id":256,"id":4,"path":"C:\\absolute\\fresh-snapshot.pptx","max_chars":3000}
```

Omit `id` for the whole slide. Each call takes a fresh snapshot, including current unsaved edits. Use a new path on each call, and `.pptm` if the presentation contains VBA. `snapshot` creates a copy without returning XML. Frequent snapshots are appropriate when package inspection costs less than many COM property calls.

`SaveCopyAs` writes a copy without changing the original presentation, as documented by [Microsoft](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas). Pass the resulting absolute snapshot path to the inspector. The inspector opens the package read-only and reads individual ZIP members without extracting the archive.

## Inspecting XML

Run:

```powershell
python scripts/ppt_package.py SNAPSHOT --slide-id N [--shape-id N] [--max-chars 6000] [--output C:\absolute\fragment.xml]
```

`--slide-id` is the stable numeric `p:sldId/@id` from `ppt/presentation.xml`. The script follows its `r:id` through `ppt/_rels/presentation.xml.rels` and resolves the relationship target. Never infer a slide's package part from its position or assume that a slide index matches a `slideN.xml` filename. `--shape-id` selects the supported `p:sp`, `p:pic`, `p:graphicFrame`, `p:grpSp`, or `p:cxnSp` container whose own `p:cNvPr/@id` matches. Without it, the complete slide XML is returned.

The JSON result contains the snapshot path, stable slide ID, resolved package part, XML text, a truncation flag, and only relationships referenced by `r:embed`, `r:link`, or `r:id` attributes in the returned fragment. `--max-chars` bounds XML included in JSON; when `--output` is given, the complete XML fragment is written to a new absolute path. Existing output files are never overwritten.

Bounds: snapshot size at most 2 GiB, at most 20,000 ZIP entries, selected XML parts at most 64 MiB each, and `--max-chars` from 1 through 1,000,000. Errors are emitted as a compact JSON object on stderr with a nonzero exit code.

## XML edits

The inspector is read-only. For an XML edit, modify a separate offline copy, preserve relationships and macro-enabled content, and reopen the edited copy in PowerPoint before relying on it. Plan that handoff around the user's current unsaved work. Do not edit or replace the open source deck's package parts directly.
