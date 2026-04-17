# Progress Log — BambuScripts (Main Branch)

Newest entries at top.  
**Main branch work** is logged in full here.  
**Worktree work** is summarised here — see each worktree's own `PROGRESS.md` for details.

---

## Session: 2026-04-15

### MAIN BRANCH

**Context system set up:**
- `CLAUDE.md` and `PROGRESS.md` created; committed to `main` and to worktree
- Session-start checklist and end-of-session reminder added to `CLAUDE.md`
- Per-worktree `PROGRESS.md` convention established

**Filament dropdown is now searchable** (`CardQueueEditorWPF.ps1`)
- Added searchable ComboBox to filament color selection UI
- Partially working — search works once but not a second time (known bug, next steps below)

**FilamentLibrary.csv** — added Baby Pink entry

**Image name fixes** (merged from `claude/practical-hellman`):
- Fixed tag duplication in generated image filenames
- Fixed image name duplication issue

**Merge worker improvements** (merged from `claude/practical-hellman`):
- `merge_3mf_worker.ps1` now groups merge pairs based on spatial proximity
- Revert merge gate adjustments

### MAIN BRANCH — Next steps

- Fix filament dropdown search so it works on a second search
- Port the working searchable dropdown to Python rewrite once PS1 version is solid

---

### WORKTREE `claude/FeaturesAndEdits` — Summary

**Status:** Active — features and general edits to the PS1/WPF toolset
**Detailed log:** `BambuScripts/.claude/worktrees/FeaturesAndEdits/BambuScripts/PROGRESS.md`
**Branched from:** `main` at `039693a`

**This session:** Worktree created.

**Next up:** _(to be filled in)_

---

### WORKTREE `claude/upbeat-grothendieck` — Summary

**Status:** Python/PySide6 CardQueueEditor rewrite — active development  
**Detailed log:** `BambuScripts/.claude/worktrees/upbeat-grothendieck/BambuScripts/PROGRESS.md`

**This session:** Completed `ColorSlotWidget.set_scale()` full proportional scaling (margins, spacing, fonts, swatch, combo all scale with card size). Context files created.

**Next up:**
1. Verify pick panel scales correctly
2. `git merge main` into worktree (PS1 file on main is 45K larger than worktree copy)
3. Fix `generate_image_worker.py` path resolution from `extract_worker.py`
4. Add `.bat` / `.vbs` launcher for Python app
5. Add `requirements.txt`

---

## Earlier sessions — pre-2026-04-15

**Python rewrite foundation** (multiple sessions on `claude/upbeat-grothendieck`):
- Full PS1/WPF → Python/PySide6 CardQueueEditor rewrite completed
- Square card layout architecture settled (`_SquareCard` passive, parent-driven sizing)
- Multiple layout bugs fixed (portrait card, slots outside border, oscillating resize, right-panel collapse)
- `py_workers/` suite built: extract, merge, isolate, slice

**Merge worker spatial grouping** (via `claude/practical-hellman`):
- PS1 merge worker updated to pair objects by proximity rather than order

See `git log` on each branch for full commit history.
