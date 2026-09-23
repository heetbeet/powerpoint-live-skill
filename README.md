# PowerPoint live skill

`powerpoint-live` gives Codex fast access to the open Windows desktop PowerPoint session. A persistent local Python process owns the COM connection and accepts compact JSON requests through a current-user named pipe.

The skill supports compact active context, deck outlines, revision-guarded edits, native slide and deck scripts, image replacement, changed-slide review, text-fit checks, Open XML snapshots, package comparison, and local request-cost metrics. It keeps screenshots optional so ordinary edits use small text responses.

## Install

Requirements:

- Windows with desktop PowerPoint
- Python 3.10 or later
- `pywin32`

Clone the repository as the skill directory, then install its Python dependency:

```powershell
git clone https://github.com/heetbeet/powerpoint-live-skill.git "$env:USERPROFILE\.codex\skills\powerpoint-live"
python -m pip install -r "$env:USERPROFILE\.codex\skills\powerpoint-live\requirements.txt"
```

PowerPoint must already be open. The first request starts the bridge:

```powershell
Set-Location "$env:USERPROFILE\.codex\skills\powerpoint-live"
'{"op":"context"}' | python scripts/ppt_live.py request
```

See [SKILL.md](SKILL.md) for agent guidance and [references/commands.md](references/commands.md) for the request schema.

The bridge keeps timing and payload estimates locally. It does not log slide text, file paths, object names, scripts, or raw responses.

## Design boundary

COM edits the open presentation. Open XML reads a current `SaveCopyAs` snapshot. The bridge does not alter a `.pptx` package beneath an open PowerPoint document.
