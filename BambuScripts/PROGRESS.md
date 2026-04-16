# Progress Log — BambuScripts

Newest entries at top. Covers both the `main` branch (PS1/WPF tools) and the
`claude/upbeat-grothendieck` worktree (Python/PySide6 rewrite). Each section is
labeled so context is clear.

---

## Session: 2026-04-15

### MAIN BRANCH — Recent work

**Filament dropdown is now searchable** (`CardQueueEditorWPF.ps1`)
- Added a searchable ComboBox to the filament color selection UI
- Currently partially working — can search once but not a second time
- File changed: `workers/CardQueueEditorWPF.ps1` (+99 lines)

**FilamentLibrary.csv** — added Baby Pink entry

**Image name fixes** (merged from `claude/practical-hellman`):
- Fixed tag duplication in generated image filenames
- Fixed image name duplication issue

**Merge worker improvements** (merged from `claude/practical-hellman`):
- `merge_3mf_worker.ps1` now groups merge pairs based on spatial proximity
- Revert merge gate adjustments

### MAIN BRANCH — Immediate next steps

- Fix filament dropdown search so it works on a **second search** (not just first)
- Consider whether fixes to `CardQueueEditorWPF.ps1` need to be ported to the Python rewrite

---

### WORKTREE (`claude/upbeat-grothendieck`) — Work this session

**Context files created:**
- `CLAUDE.md` and `PROGRESS.md` added and committed to both branches

**`ColorSlotWidget` full proportional scaling — completed:**

`set_scale(scale, swatch_px, combo_max_px)` now updates ALL sub-elements proportionally:
- `_row` bottom margin + spacing
- `_text_col` right margin + spacing
- `_status_fs` → status label font size (re-renders current status colour)
- `_num_fs` → slot-number label font size
- `_num_color` kept in sync when `_update_swatch` changes swatch colour
- Combo min/max width + fit-to-text

`_set_status` and `_update_swatch` updated to use `self._status_fs` / `self._num_fs` instead of hardcoded pixel values.

Call-site in `_on_card_resize` updated: `sw.set_scale(scale, swatch_px, combo_max)`.

### WORKTREE — Immediate next steps

1. **Verify pick panel** — confirm `_build_pick_panel` uses the same `_SquareCard` architecture and scales correctly
2. **Merge main into worktree** — `CardQueueEditorWPF.ps1` on main is newer (171K vs 126K in worktree); user-requested merge pending:
   ```bash
   cd BambuScripts/.claude/worktrees/upbeat-grothendieck
   git merge main
   ```
3. **Fix `generate_image_worker.py` not-found error** — worker path may be hardcoded incorrectly when `extract_worker.py` calls it as a subprocess from various working directories
4. **Port searchable dropdown fix** — the PS1 filament dropdown search work should be mirrored in the Python `color_slot_widget.py` / `ColorLibrary` once the PS1 version is fully working
5. **Add launcher** — Python app has no `.bat` or `.vbs` launcher yet; can only be run via `python app.py`
6. **Add `requirements.txt`** — `PySide6` and `Pillow` are undocumented dependencies

### WORKTREE — Open questions / blockers

- Searchable dropdown in PS1 only works once — root cause unknown; needs investigation before porting to Python
- `generate_image_worker.py` lives outside `workers/py/` — subprocess path resolution from `extract_worker.py` may break depending on working directory at launch time
- No test suite; all validation is manual

---

## Earlier sessions — Python rewrite foundation (pre-context-summary)

### WORKTREE — Completed in prior sessions

**Full PS1/WPF → Python/PySide6 rewrite of CardQueueEditor:**
- `app.py`, `main_window.py`, `gp_widget.py` — fully functional
- `parent_widget.py` — card panel, pick panel, right panel, file list rows
- `color_slot_widget.py` — single filament slot (status badge + name combo + swatch)
- `models.py`, `color_library.py`, `file_utils.py`, `image_utils.py`, `theme.py`
- `py_workers/`: `extract_worker.py`, `merge_worker.py`, `isolate_worker.py`, `slice_worker.py`

**Card panel squareness and proportional scaling — architecture settled:**
- `_SquareCard` passive (only `setFixedSize`); `PJobWidget.resizeEvent` drives all sizing
- Deferred resize via `QTimer.singleShot(0, _update_card_sizes)` prevents oscillation
- `_SLOT_RATIO = 0.416` — color-slots column as fraction of card side
- Title bar spans full card width via `QVBoxLayout` on the card

**Key bugs fixed:**
- Portrait rectangle instead of square → `setFixedSize(both)`
- Slots outside card border → `QVBoxLayout` with title bar at top
- Oscillating resize loop → passive card, parent-owned sizing
- Right panel collapsing → `setMinimumWidth(~965px)` on `PJobWidget`
- Combo overflow → `combo_max = slot_w - swatch_px - margins` computed explicitly

**Git maintenance:**
- Restored accidentally-switched branch back to `claude/upbeat-grothendieck`
- Removed 5 orphaned nested worktrees
