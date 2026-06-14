# BambuScripts — Project Context for Claude

## Session Start Checklist

> **Do this at the start of every session, before anything else:**
> 1. Read `BambuScripts/CLAUDE.md` (this file) — confirm architecture and conventions loaded
> 2. Read `BambuScripts/PROGRESS.md` — confirm current status and next steps loaded
> 3. Confirm to the user: "Context loaded — next up: [top item from next steps]"

> **At the end of every session:**
> - Update `PROGRESS.md` with what was done, decisions made, and revised next steps
> - Remind the user to let you do this if you haven't yet

---

## Project Overview

A Windows toolset for managing Bambu 3D printer batch workflows. Handles the full pipeline:
color picking → 3MF merging → slicing → data extraction → preview image generation → card naming/renaming.

The **CardQueueEditor** (`CardQueueEditorWPF.ps1`) is the primary interactive UI — a
PowerShell/WPF application.

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

### CRITICAL — PS1 File Encoding

**Never use typographic/Unicode punctuation in `.ps1` files.** PowerShell parses them as
bare tokens and breaks the entire script silently (editor won't open, no error shown).

| Forbidden | Use instead |
|---|---|
| `—` em dash (U+2014) | `-` hyphen-minus |
| `–` en dash (U+2013) | `-` hyphen-minus |
| `'` `'` curly quotes | `'` straight single quote |
| `"` `"` curly quotes | `"` straight double quote |
| `…` ellipsis (U+2026) | `...` three periods |

These characters are injected silently by autocorrect/smart-quotes in editors and LLM outputs.
**Always run a parse check after editing PS1 files:**
```powershell
$e=$null; [System.Management.Automation.Language.Parser]::ParseFile('path\to\file.ps1',[ref]$null,[ref]$e); if($e){"ERRORS: $($e.Count)"; $e|%{"  Line $($_.Extent.StartLineNumber): $($_.Message)"}}else{"Parse OK"}
```

---

## Important Constraints

- **Windows-only**: COM interop, WPF, WinForms, hardcoded Bambu Studio path
- **No package manager / no tests**: Manual testing through the GUI
- **Platform**: PowerShell 3.0+, .NET Framework (PresentationFramework, System.Windows.Forms)

