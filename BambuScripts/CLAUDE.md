# BambuScripts — Project Context for Claude

## Project Overview

A Windows toolset for managing Bambu 3D printer batch workflows. Handles the full pipeline:
color picking → 3MF merging → slicing → data extraction → preview image generation → card naming/renaming.

The **CardQueueEditor** is the primary interactive UI — originally built in PowerShell/WPF, currently being rewritten in Python/PySide6. It presents a queue of print jobs (cards), lets users review/edit filament color assignments, character names, and adjectives, then launches the pipeline workers.

---

## Tech Stack

| Layer | Technology |
|---|---|
| GUI (new) | Python 3.12+, PySide6 |
| GUI (legacy) | PowerShell WinForms / WPF |
| Entry points | VBScript wrappers, `.bat` callers |
| Image generation | Python + Pillow |
| 3MF manipulation | Python (zipfile + xml.etree) |
| Slicing | Bambu Studio CLI |
| Data | CSV, TSV, ZIP/3MF, JSON |

**Dependencies (not in requirements.txt — install manually):**
- `PySide6` — GUI framework
- `Pillow` — image processing in workers and `image_utils.py`

---

## Directory Structure

```
BambuScripts/
├── CLAUDE.md              ← this file
├── PROGRESS.md            ← session progress log
├── callers/               ← user-facing drag-and-drop .bat scripts
├── launchers/             ← invisible VBS wrappers for PS1 scripts
├── workers/
│   ├── *.ps1              ← legacy PowerShell workers (still in use)
│   ├── generate_image_worker.py  ← standalone 512×512 card image generator
│   └── py/                ← Python/PySide6 CardQueueEditor rewrite
│       ├── app.py             ← entry point — QApplication + MainWindow
│       ├── main_window.py     ← top-level window, drag-drop, process queue
│       ├── parent_widget.py   ← PJobWidget (one card row) + _SquareCard
│       ├── gp_widget.py       ← GpWidget (one theme group of cards)
│       ├── color_slot_widget.py  ← single filament slot row widget
│       ├── models.py          ← dataclasses + color palette constants
│       ├── color_library.py   ← loads colorNamesCSV.csv, hex↔name lookup
│       ├── file_utils.py      ← 3MF parsing, SmartFill, sort keys
│       ├── image_utils.py     ← pick-color randomize, merge-map vis, PIL
│       ├── theme.py           ← global dark QSS stylesheet
│       └── py_workers/        ← subprocess workers called by PJobWidget
│           ├── extract_worker.py   ← gcode parsing → TSV + image
│           ├── merge_worker.py     ← N-way 3MF merge
│           ├── isolate_worker.py   ← isolate center object from merged 3MF
│           └── slice_worker.py    ← invoke Bambu Studio CLI
└── libraries/
    ├── FilamentLibrary.csv    ← filament brand color data
    └── NamesLibrary.ps1       ← shared theme names, printer prefixes
```

---

## Run Command

```bash
cd BambuScripts/workers/py
python app.py
# or with initial files:
python app.py "path\to\folder1" "path\to\folder2"
```

No build step. No tests.

---

## Architecture: Card Panel Layout

The UI card (`_SquareCard`) must be a **perfect square** at all times.

**Key design decisions:**
- `_SquareCard` is **passive** — it only has `set_side(n)` which calls `setFixedSize(n, n)` and fires a callback.
- `PJobWidget.resizeEvent` drives all sizing via `_update_card_sizes()` → `set_side()` → `_on_card_resize(scale)`.
- No `hasHeightForWidth`, no `resizeEvent` on `_SquareCard` itself — avoids oscillating resize loops.
- `_SquareCard._SLOT_RATIO = 0.416` controls color-slots column width as fraction of card side.

**Proportional scaling chain:**
1. `PJobWidget.resizeEvent` fires
2. `QTimer.singleShot(0, _update_card_sizes)` — deferred to avoid mid-layout calls
3. `_update_card_sizes` computes `side`, calls `card.set_side(side)` and `pick.set_side(side)`
4. `set_side` calls `_on_card_resize(side)` callback
5. `_on_card_resize` computes `scale = side / DEFAULT`, resizes all child elements proportionally
6. Calls `slot_widget.set_scale(scale, swatch_px, combo_max)` on each `ColorSlotWidget`

---

## Code Conventions (Python/PySide6 files)

- **snake_case** for all functions, variables, private members (`_build_ui`, `_on_card_resize`)
- **PascalCase** for classes (`PJobWidget`, `ColorSlotWidget`, `_SquareCard`)
- **SCREAMING_SNAKE** for module-level constants (`COLOR_BG_DARK`, `STATUS_MATCHED`)
- **Type hints** used on public APIs; internal helpers may omit them
- `from __future__ import annotations` at top of every file
- `QTimer.singleShot(0, fn)` is the standard pattern for deferred layout work
- All color constants live in `models.py` — never hardcode hex strings in widget files
- Widget stylesheets use f-strings with `models.py` constants: `f'color:{COLOR_AMBER}; font-size:{fs}px;'`
- Private widget refs are stored as `self._xxx` at build time so `set_scale` can update them without rebuilding

---

## Key Data Files

- `libraries/colorNamesCSV.csv` — loaded by `ColorLibrary`; maps hex → filament name
- `libraries/FilamentLibrary.csv` — used by `generate_image_worker.py` for gradient rendering
- `libraries/NamesLibrary.ps1` — theme/printer/adjective lists (mirrored in `models.py`)

---

## Important Constraints

- **Windows-only**: COM interop, hardcoded Bambu Studio path (`C:\Program Files\Bambu Studio\`)
- **No package manager**: Dependencies installed ad hoc; no `requirements.txt`
- **No tests**: Manual testing through the GUI
- Active branch for Python work: `claude/upbeat-grothendieck`
- Worktree path: `BambuScripts/.claude/worktrees/upbeat-grothendieck/`
