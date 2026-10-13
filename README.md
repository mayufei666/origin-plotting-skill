# origin-plotting (Codex skill)

A Codex skill that drives an **already-open OriginLab Origin / OriginPro** over
Windows COM Automation to build, style, and export figures - no clipboard, no
manual clicking.

Verified end-to-end on Origin 64-bit (Origin 2021+) on Windows with
Python 3.12 + `pywin32`.

## What it does

- Connects to the running Origin instance via `Origin.ApplicationSI`
- Writes data into a worksheet with `PutWorksheet` and reads it back to verify
- Creates stacked XRD patterns, bar/column, line, scatter, and line+symbol plots
- Styles axes, legends, curve colors, and text annotations through LabTalk
- Exports a PNG with `expGraph` and checks the file was actually written
- Never saves `.opju` (you press Ctrl+S) and treats source data as read-only

## Repository layout

```
skills/origin-plotting/
├── SKILL.md
├── agents/openai.yaml
├── references/
│   ├── plot-types.md
│   └── xrd-recipe.md
└── scripts/
    └── origin_plot.py
```

## Install

### Option A - copy into your Codex skills folder

```powershell
git clone https://github.com/mayufei666/origin-plotting-skill.git
Copy-Item -Recurse -Force `
  .\origin-plotting-skill\skills\origin-plotting `
  "$env:USERPROFILE\.codex\skills\origin-plotting"
```

Reload or restart Codex, then invoke it with `$origin-plotting`.

### Option B - install straight from GitHub

If your Codex build supports the skill-installer:

```
python scripts/install-skill-from-github.py `
  --repo mayufei666/origin-plotting-skill `
  --path skills/origin-plotting
```

## Use the generic plotting script directly

```powershell
python skills\origin-plotting\scripts\origin_plot.py `
  --data "E:\data\sample.csv" `
  --x-col 1 --y-cols 2,3 `
  --plot-type bar `
  --out "C:\OriginData\bar.png" `
  --xlabel "Category" --ylabel "Value"
```

`--plot-type` accepts `line`, `scatter`, `line+symbol`, `bar`.

## Requirements

- Windows
- OriginLab Origin / OriginPro installed and **already running**
- Python 3.9+ with `pywin32` (`pip install pywin32`)

## Verified gotchas baked into the skill

| Problem | What works here |
| --- | --- |
| `originpro` package hangs | Use raw `win32com` COM instead |
| `Origin.Application` unreliable | Use `Origin.ApplicationSI` |
| Bar chart `plot:=300` / `310` gives an empty graph | Use `plot:=203` |
| Clipboard export fails | Use LabTalk `expGraph` |
| `.opju` save via COM crashes the connection | Let the user press Ctrl+S |
| Raw color numbers fail silently | Resolve via `color(r,g,b)` |
| LabTalk eats parentheses in labels | Quote the annotation text |

## License

MIT - see [LICENSE](LICENSE).
