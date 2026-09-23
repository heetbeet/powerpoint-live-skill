# Design rationale

## Runtime and connection

Use the external 64-bit Python 3.13 process with pywin32 because it is available and live COM attachment has been verified. This is an availability choice, not a performance claim against C#. Keep one persistent PowerPoint COM connection on a single STA thread. Restrict the named pipe to the current user and session, serialize requests, and pump Windows messages while serving them.

The goal is an HTML-like live view with low latency and small responses. Return compact inspections of supported properties instead of full object dumps. A bounded property set keeps each response small, but COM queries still grow with the number of shapes.

## State and edits

Treat a fresh snapshot or fingerprint of supported properties as authoritative. Office selection events do not report every edit, so use them as hints rather than a complete change log. Before each edit, query a fresh revision and check it against the revision the caller inspected.

Batches are not atomic and can leave partial changes on failure. Never automatically replay a mutation after a timeout because its outcome is unknown. Render when visual judgment is needed or the user requests an image. For package-level XML work, use PowerPoint SaveCopyAs, then process the copy offline.

## Measurements and limits

In the prototype environment, direct COM attachment took 11.5 ms and an inventory of shape types and counts took 219.4 ms for a 9-slide deck sized 960 x 540. These are component observations, not final bridge latency.

The prototype does not provide comprehensive Office events, an add-in, an MCP plugin, permanent tag GUIDs, a full DOM or cascading layout model, or the entire Office API. Consider C# or an in-process add-in only if end-to-end measurements show the bridge is a bottleneck; current evidence does not compare their performance.

## Measured editing session

Live investigation used an open 9-slide presentation. Temporary shapes were created and removed, with the original shapes' inspected properties compared before and after. These are local Python-to-bridge measurements, excluding model thinking, tool orchestration, and starting a new CLI process.

| Request | Observed time | Observation |
| --- | ---: | --- |
| Small slide inspection | 85 to 128 ms | Three original shapes |
| Simple edit batch | 210 to 270 ms | Create, move, resize, fill, remove |
| Targeted native properties | 27 to 46 ms | Margins, text runs, autofit |
| Native Python read | 30 ms | Direct COM references |
| Fresh snapshot and shape XML | 111 ms | Bounded XML response |
| Entire 54-object slide | 1389 ms | 10,073 JSON characters in result |
| One object on that slide | 53 ms | 236 JSON characters in result |

The first implementation read every object even for a targeted request. Measurement exposed that cost. Targeted inspections now fingerprint only selected objects and group ancestry; a scoped revision cannot authorize edits to other existing objects. Full slide reads remain available when the task needs them.

Native text fitting was checked with an overflowing 140 by 30 point text box. PowerPoint reduced its measured text height from 604.8 to 23.2 points without changing the box height. The usable height was 22.8 points; allow for subpoint measurement differences and visually check small text rather than treating autofit as a readability guarantee.

## Microsoft primary sources

- [Office threading support](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio) and [COM apartments](https://learn.microsoft.com/en-us/windows/win32/com/processes--threads--and-apartments): Office calls are serialized through STA; the STA must pump messages.
- [PowerPoint Application event catalog](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application), [AfterShapeSizeChange](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.aftershapesizechange), [SlideSelectionChanged](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideselectionchanged), and [WindowSelectionChange](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.windowselectionchange): documented notifications and their limits.
- [Shape.Id](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.id) and [Shape.Tags](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.tags): ID and optional custom-property tags. The prototype uses the unmodified `Shape.Id`, qualified by bridge deck handle plus `SlideID`; this composite key is bridge-session scoped. Tags are optional, and read-only inspection does not stamp a persistent GUID or dirty the deck.
- [Presentation.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas), [Slide.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export), and [Shape.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export): snapshot copies and slide or shape exports.
- [PresentationML structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document) and [SDK slide relationship example](https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-get-all-the-text-in-all-slides-in-a-presentation): map each ordered `<p:sldId>` through its `r:id` relationship to the slide part; do not infer the XML filename from slide order.

Scope is PowerPoint running in the user's interactive Windows desktop session.
