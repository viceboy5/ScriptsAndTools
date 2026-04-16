# BambuScripts — Project Context for Claude

## Session Start Checklist

> **Do this at the start of every session, before anything else:**
> 1. Read `BambuScripts/CLAUDE.md` (this file) — confirm architecture and conventions loaded
> 2. Read `BambuScripts/PROGRESS.md` — confirm current branch status and next steps loaded
> 3. If working inside a worktree, also read that worktree's own `PROGRESS.md` for detailed task state
> 4. Confirm to the user: "Context loaded — [branch/worktree], next up: [top item from next steps]"

> **At the end of every session:**
> - Update the worktree's `PROGRESS.md` with what was done, decisions made, and revised next steps
> - Update the summary entry in the main `PROGRESS.md` worktree section
> - Remind the user to let you do this if you haven't yet

---

## Project Overview

A Windows toolset for managing Bambu 3D printer batch workflows. Handles the full pipeline:
color picking → 3MF merging → slicing → data extraction → preview image generation → card naming/renaming.

The **CardQueueEditor** (`CardQueueEditorWPF.ps1`) is the primary interactive UI on `main` — a
PowerShell/WPF application. A full Python/PySide6 rewrite is in progress on a separate worktree
(see [Worktree: Python Rewrite](#worktree-python-rewrite-claudeupbeat-grothendieck) below).

---

## Tech Stack — Main Branch

| Layer | Technology |
|---|---|
| Primary GUI | PowerShell WinForms / WPF (`CardQueueEditorWPF.ps1`) |
| Entry points | VBScript wrappers (`launchers/`), `.bat` callers (`callers/`) |
| Image generation | Python + Pillow (`generate_image_worker.py`) |
| Data | CSV, TSV, ZIP/3MF, JSON |
| Slicing | Bambu Studio CLI (`C:\Program Files\Bambu Studio\bambu-studio.exe`) |

**Python dependency (install manually):** `pip install Pillow`

---

## Directory Structure — Main Branch

```
BambuScripts/
├── CLAUDE.md              ← this file
├── PROGRESS.md            ← session progress log
├── callers/               ← user-facing drag-and-drop .bat scripts
├── launchers/             ← invisible VBS wrappers (hide PS1 console window)
├── workers/
│   ├── CardQueueEditorWPF.ps1     ← PRIMARY UI: card queue editor (WPF)
│   ├── Master-Controller.ps1      ← orchestrates the full build pipeline
│   ├── generate_image_worker.py   ← standalone 512×512 card image generator
│   ├── DataExtract_worker.ps1     ← parses gcode → TSV data row
│   ├── merge_3mf_worker.ps1       ← N-way 3MF merge with spatial grouping
│   ├── isolate_final_worker.ps1   ← isolates center object from merged 3MF
│   ├── Slice_worker.ps1           ← invokes Bambu Studio CLI
│   ├── Renamer.ps1                ← file renaming UI
│   └── ...other workers...
└── libraries/
    ├── FilamentLibrary.csv        ← filament brand color + gradient data
    ├── colorNamesCSV.csv          ← hex → filament display name mapping
    └── NamesLibrary.ps1           ← shared theme names, printer prefixes, adjectives
```

---

## How to Run (Main Branch)

- **CardQueueEditor UI:** Double-click `launchers/CardQueueEditorWPF.vbs`  
  (or: `powershell -STA -ExecutionPolicy Bypass -File workers/CardQueueEditorWPF.ps1`)
- **Image generator:** `python workers/generate_image_worker.py --name X --time 120 --out outdir/ ...`
- **Full pipeline:** Launch `Master-Controller.ps1` via its caller/launcher

---

## Key Data Files

- `libraries/colorNamesCSV.csv` — hex → filament name; used by both PS1 and Python UIs
- `libraries/FilamentLibrary.csv` — brand/color data for gradient rendering in image generator
- `libraries/NamesLibrary.ps1` — theme names, printer prefixes (`X1C`, `P2S`, `H2S`), adjective presets

---

## Code Conventions — PowerShell

- PascalCase functions and variables: `$chkColors`, `Invoke-SliceFile`
- Section markers for readability: `# --- 1. BUILD THE MAIN WINDOW ---`
- Color palette defined as hex strings near top of each file
- Tab-separated values (TSV) for inter-process communication
- WPF XAML defined inline as here-strings

---

## Important Constraints

- **Windows-only**: COM interop, WPF, WinForms, hardcoded Bambu Studio path
- **No package manager / no tests**: Manual testing through the GUI
- **Platform**: PowerShell 3.0+, .NET Framework (PresentationFramework, System.Windows.Forms)

---

## Progress File Convention

| Location | Purpose |
|---|---|
| `BambuScripts/PROGRESS.md` | Main branch log + brief summary of each active worktree |
| `<worktree>/BambuScripts/PROGRESS.md` | Detailed log for that specific worktree only |

**Every new worktree gets its own `PROGRESS.md` created at the start of the first session in it.**  
The main `PROGRESS.md` worktree section is updated with a brief summary at the end of each session.

---

## Worktree: Python Rewrite (`claude/upbeat-grothendieck`)

A full Python/PySide6 rewrite of the CardQueueEditor is in progress on a separate git worktree.
It is **not yet merged to main**.

**Worktree path:** `BambuScripts/.claude/worktrees/upbeat-grothendieck/`  
**Branch:** `claude/upbeat-grothendieck`  
**Run:** `cd BambuScripts/.claude/worktrees/upbeat-grothendieck/BambuScripts/workers/py && python app.py`

**Additional dependencies:** `pip install PySide6 Pillow`

### Python rewrite structure (`workers/py/`)

| File | Purpose |
|---|---|
| `app.py` | Entry point — QApplication + MainWindow |
| `main_window.py` | Top-level window, drag-drop, process queue |
| `parent_widget.py` | `PJobWidget` (one card row) + `_SquareCard` |
| `gp_widget.py` | `GpWidget` — one theme group of cards |
| `color_slot_widget.py` | Single filament slot row widget |
| `models.py` | Dataclasses + color palette constants |
| `color_library.py` | Loads `colorNamesCSV.csv`, hex↔name lookup |
| `file_utils.py` | 3MF parsing, SmartFill, file sort keys |
| `image_utils.py` | Pick-color randomize, merge-map visualization |
| `theme.py` | Global dark QSS stylesheet |
| `py_workers/` | Subprocess workers: extract, merge, isolate, slice |

### Key architecture decision: square card layout

`_SquareCard` is **passive** — only `set_side(n)` → `setFixedSize(n, n)` + fires callback.  
`PJobWidget.resizeEvent` drives all sizing: `_update_card_sizes()` → `set_side()` → `_on_card_resize(scale)` → `slot.set_scale(scale, swatch_px, combo_max)`.  
No `hasHeightForWidth` / no `resizeEvent` on the card itself — avoids oscillating resize loops.

### Python code conventions

- `snake_case` functions/variables, `PascalCase` classes, `SCREAMING_SNAKE` constants
- `from __future__ import annotations` at top of every file
- All color constants in `models.py` — never hardcode hex in widget files
- `QTimer.singleShot(0, fn)` for all deferred layout work
- Store widget/layout refs as `self._xxx` at build time so `set_scale` can update without rebuilding
