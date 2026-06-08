# Progress Log — main

Branch: `main`

Newest entries at top.

---

### Session 2026-06-08

**Status:** Active

**What was done:**
- Built a full Purge Dictionary editor panel inside CardQueueEditorWPF's Libraries section: sortable/filterable grid (From/Tuned/Tuned Volume filters), Tuned checkbox replacing old Iterations/Test_Volume columns, color swatches on From/To, centered numeric columns, dirty-cell highlighting (orange) with Save Changes gating, legible column headers.
- Fixed a CSV regression where PurgeDictionary.csv had been converted to tab-delimited without a BOM (broke Excel display); restored comma-delimited + UTF-8 BOM in `Save-PurgeDictionary`/`Load-PurgeDictionary`, with auto-migration of any tab-delimited copies on load. Reverted the matching `-Delimiter` hack in `UpdatePurgeMatrix_worker.ps1`.
- Made the three Purge Dictionary filter dropdowns typable/searchable (`Enable-PurgeComboTypeAhead` helper), mirroring the filament color-picker combo pattern: IsEditable + CollectionView text filtering + Enter/Tab/Escape handling.
- Added a "% Savings" column (Base vs Tuned volume) and a rolling "Avg Savings (tuned): X.X% across N combos" label at the top of the panel, computed across tuned rows only.

**Gotcha worth remembering:** `.GetNewClosure()` scriptblocks created *inside a function* silently resolve `$script:`-scoped variables to `$null` (and `$null.Count` quietly returns `0` rather than erroring) — always capture a direct local reference to script-scoped collections (e.g. `$capturedPurgeDict = $script:PurgeDict`) before building closures that need them, same pattern as `$capturedPurgeView`.

**Next up:** _(to be filled in — no open threads from this session; Purge Dictionary editor feature set is complete and smoke-tested)_

---

### Session 2026-06-03

**Status:** Active

**Recent commits on main:**
- `86eaa13` re-introduced combine TSV Button
- `10bd85c` added stop Queue button
- `feddc70` fixed P2S gcode parsing for model filament and purge filament
- `8bba3f7` ignore results
- `4b453e4` SKU Edits, Data Format, Prehistoric changed to Dinos, SKU Syncing

**Next up:** _(to be filled in)_
