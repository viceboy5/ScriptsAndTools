# Progress Log — BambuScripts CardQueueEditor Python Conversion

Newest entries at top. Each entry documents a session's work, decisions, and next steps.

---

## Session: 2026-04-15 (Context Setup + ColorSlotWidget Scaling)

### What was completed (prior sessions, reconstructed from summary)

**Python/PySide6 CardQueueEditor rewrite — initial conversion complete:**
- `app.py`, `main_window.py`, `gp_widget.py` — fully functional
- `parent_widget.py` — main card panel, pick panel, right panel, file list rows
- `color_slot_widget.py` — single filament slot (status badge + name combo + swatch)
- `models.py`, `color_library.py`, `file_utils.py`, `image_utils.py`, `theme.py`
- `py_workers/`: `extract_worker.py`, `merge_worker.py`, `isolate_worker.py`, `slice_worker.py`

**Card panel squareness and proportional scaling — architecture settled:**

The final working architecture for the square card:
- `_SquareCard` is **passive** — only `set_side(n)` → `setFixedSize(n, n)` + fires callback. No `resizeEvent` or `hasHeightForWidth` on the card itself.
- `PJobWidget.resizeEvent` → `QTimer.singleShot(0, _update_card_sizes)` → computes `side` → calls `card.set_side(side)` and `pick.set_side(side)`.
- `_on_card_resize(side)` receives the new side length, computes `scale = side / DEFAULT (380)`, resizes all children proportionally.
- `_SquareCard._SLOT_RATIO = 0.416` — color slots column width as fraction of card side.
- `set_scale(scale, swatch_px, combo_max)` on `ColorSlotWidget` — updates all slot sub-elements proportionally.

**Key bugs fixed along the way:**
- Card rendered as portrait rectangle → fixed with `setFixedSize` (both dimensions).
- Filament slots appearing outside card border → fixed by restructuring to `QVBoxLayout` with title bar spanning full width.
- Oscillating resize loop → fixed by making `_SquareCard` passive with external sizing control.
- Right panel collapsing → fixed with `setMinimumWidth(~965px)` on `PJobWidget`.
- Combo width overflow → fixed by computing `combo_max = slot_w - swatch_px - margins` explicitly.

**Git maintenance performed:**
- Restored accidentally switched branch back to `claude/upbeat-grothendieck`.
- Removed 5 orphaned nested worktrees from inside the worktree.

### What was most recently worked on (last thing before context ran out)

`color_slot_widget.py` — `_build_ui` was updated to store layout and font-size references:
- `self._row` — `QHBoxLayout` reference (for margin/spacing updates)
- `self._text_col` — `QVBoxLayout` reference (for margin/spacing updates)
- `self._status_fs = 10` — status label font size
- `self._num_fs = 14` — slot-number label font size
- `self._num_color` — contrast color for slot number

`_build_ui` edit was applied. **`set_scale` was NOT yet updated** — it still only resizes swatch + combo, does not touch margins, spacing, or fonts.

### Immediate next steps

**1. Complete `ColorSlotWidget` full proportional scaling** ← DO THIS FIRST

Update `set_scale` signature to `(self, scale: float, swatch_px: int, combo_max_px: int)` and add:
```python
self._row.setContentsMargins(0, 0, 0, max(4, int(15 * scale)))
self._row.setSpacing(max(2, int(6 * scale)))
self._text_col.setContentsMargins(0, 0, max(4, int(10 * scale)), 0)
self._text_col.setSpacing(max(1, int(2 * scale)))
self._status_fs = max(7, int(10 * scale))
self._num_fs = max(8, int(14 * scale))
self._combo.setMinimumWidth(max(30, int(35 * scale)))
# then update lbl_status and lbl_num stylesheets
```

Update `_set_status` to use `self._status_fs` instead of hardcoded `10px`.
Update `_update_swatch` to use `self._num_fs` instead of hardcoded `14px`.

**2. Update `_on_card_resize` in `parent_widget.py`**

Change `sw.set_scale(swatch_px, combo_max)` → `sw.set_scale(scale, swatch_px, combo_max)`.

**3. Verify pick panel**

Confirm `_build_pick_panel` uses the same `_SquareCard` architecture and scales correctly.

**4. Merge main branch changes**

User asked earlier to merge `main` into this worktree:
```bash
cd BambuScripts/.claude/worktrees/upbeat-grothendieck
git merge main
```
Check for conflicts before doing this.

**5. Fix `generate_image_worker.py` not-found error**

From earlier sessions — the worker path may be hardcoded incorrectly when called from `extract_worker.py`. Investigate.

### Open questions / blockers

- No `requirements.txt` exists — if someone clones fresh, they won't know to `pip install PySide6 Pillow`.
- The Python app has no launcher `.bat` or `.vbs` yet — it can only be run with `python app.py` from the command line.
- `generate_image_worker.py` is outside the `py/` package — unclear if `extract_worker.py` finds it correctly when launched as a subprocess from various working directories.

---

## Session: (earlier sessions — pre-context-summary)

CardQueueEditor Python conversion initiated. Full PS1/WPF → Python/PySide6 rewrite performed across multiple sessions. See git log on branch `claude/upbeat-grothendieck` for commit history.
