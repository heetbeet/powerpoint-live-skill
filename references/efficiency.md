# Efficient PowerPoint work

## Default sequence

Use the cheapest request that answers the next decision:

1. `context` to acquire the active target and current revision.
2. `outline` once for a deck-wide task.
3. Scoped `inspect` or `native_read` for the affected objects.
4. One guarded `batch`, or one `native_deck` script for a consistent cross-slide pass.
5. `check` for geometry and text-bound hints.
6. `review` for changed slides, or `render_shape` for one local visual issue.
7. `snapshot` or `save` at a meaningful checkpoint.

Skip any step that does not inform the next action. Reuse IDs and revisions while they remain current. After a manual edit or stale-revision response, refresh the affected scope once.

## Evidence from a long co-editing session

The session that motivated these rules made 321 bridge requests. The request bodies contained about 9,300 estimated text tokens, while responses contained about 127,600. Inspection dominated the cost:

| Operation | Calls | Estimated response tokens |
| --- | ---: | ---: |
| `inspect` | 117 | 84,603 |
| `native` | 111 | 24,690 |
| `deck` | 25 | 13,716 |
| `batch` | 13 | 1,269 |

There were 17 failures, 10 repeated inspections, 106 mutations, and 11 snapshots. The productive patterns were compact native result dictionaries, grouped changes, stable slide and shape IDs, targeted renders, and disk paths for images. The expensive patterns were repeated full slide inspection, repeated deck enumeration, many one-off one-slide scripts, and loading several full slide images.

The response estimate is `ceil(serialized JSON characters / 4)`. It is a directional measure of text passed through the bridge. It excludes model reasoning, tool wrappers, and image tokens, so it is not a bill or an exact context count.

## Practical rules

- Prefer `context` over separate `deck`, `selection`, and `inspect` calls.
- Ask `deck` for detail only when choosing among presentations or auditing slide IDs.
- Use `outline` for deck structure instead of inspecting every slide.
- Restrict `inspect` with `ids`, a shallow `depth`, a low `text_limit`, and `detail:false` until detail is necessary.
- Use a full slide revision before an unrestricted native write. Use scoped revisions for scoped batch edits.
- Return counts, IDs, short labels, and exception lists from native scripts. Do not return COM object representations.
- Run consistent cross-slide changes in one `native_deck` call with expected revisions.
- Bundle alignment, spacing, font, color, and text corrections that share one visual decision.
- Render changed slides once after a pass. Open individual images only when they reveal an issue.
- Use `SaveCopyAs` snapshots for stable XML or package inspection instead of many property probes.

Never save tokens by skipping a visual check that materially affects quality. Use the targeted visual route: one object export for a local issue, or one changed-slide review for composition.

## Metrics

Use `{"op":"metrics"}` after a substantial phase or when the workflow feels slow. Counters reset when the bridge restarts. Numeric records remain in `%LOCALAPPDATA%/Codex/powerpoint-live/metrics.jsonl`, with one rotated prior file.

The log omits slide text, file paths, object names, scripts, and raw responses. It reports timing, request and response sizes, estimated text tokens, repeated same-state reads, repeated renders, large replies, failures, snapshots, and mutations. Treat the counters as prompts for a better next request rather than targets that override slide quality.
