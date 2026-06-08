# Progress Log — main

Branch: `main`

Newest entries at top.

---

### Session 2026-06-08

**Status:** Active

**Bug fix — BOD creation for Final-only design folders (`CardQueueEditorWPF.ps1` ~line 1435/1471):**
- Symptom: BOD.3mf appeared briefly in the design folder, a dated folder was created in the Printing Queue, but the sliced gcode never got moved there.
- Root cause: `$basePrefix` was computed as `$baseName + "_"` for any anchor not ending in "Full". For a Final-only folder the anchor is `..._Final.3mf`, so `$baseName = "..._Final"` and `$basePrefix` became `"..._Final_"`, which made `$bodFullPath` (`"$baseName.3mf"` = `"..._Final.3mf"`) coincidentally equal the *actual* Final.3mf path — tricking the mode-detection `if (Test-Path $bodFullPath) { 'Full' }` into always picking **Full mode** instead of **Grid mode** for Final-only designs. Full-mode `create_bod_worker.ps1` then ran its pair-pruning logic against a single isolated object, producing a malformed/empty BOD whose slice never emitted gcode (matching the existing `(?i)_Final$` stem-stripping convention used elsewhere, e.g. lines 3363/3561/4356).
- Fix: (1) added an `elseif ($baseName -match '(?i)_Final$')` branch to `$basePrefix` so it strips `_Final` the same way other code paths do; (2) replaced the coincidence-prone `Test-Path $bodFullPath` mode check with an explicit `$bodIsFullDesign = $baseName.ToLower().EndsWith("full")` flag baked into the generated worker script, so mode is now derived from the anchor's actual role rather than a path collision.
- Found and fixed two more downstream issues surfaced once Grid mode actually ran on a single-object Final design:
  - `create_bod_worker.ps1:116` `$origTxList = $origBuildItems | ForEach-Object {...}` flattened to a bare `[double[]]` (instead of an array-of-arrays) when there was exactly one build item, causing `$origTxList[$i].Clone()` to throw `"[System.Double] does not contain a method named 'Clone'"`. Fixed by wrapping each `Parse-Tx` result with the unary comma operator and `@()`: `@($origBuildItems | ForEach-Object { , (Parse-Tx (...)) })`.
  - The 15-copy grid layout selection in `create_bod_worker.ps1` (Grid mode) picked the column/row split that *maximised* the minimum gap between copies, which spread them almost edge-to-edge on the 256x256 plate and put outer copies inside the X1C's bed exclusion zone (Bambu Studio refused to slice: "too close to exclusion area"). First pass scored layouts by closeness to a `$TARGET_GAP` (8mm) instead of maximising — but the gap formula `(avail - n*size)/(n-1)` always *stretches* to consume the full available span regardless of score, so the result still spanned almost the whole plate (Hank correctly called this out from the preview screenshot). Replaced it entirely: layouts are now built at a **fixed** `$TARGET_GAP` and the cols x rows split is chosen to waste the fewest empty grid cells (ties broken toward the most square-ish shape); leftover plate space becomes extra centring margin instead of inflating the gaps. For the 22x33mm BlackKnight design this now picks a tight grid instead of a stretched 4x4 at ~46x31mm gaps. Per Hank's request, dropped `$MIN_GAP` entirely (superseded by `$TARGET_GAP`), lowered `$TARGET_GAP` to 3mm, and added a third tie-break (prefer more columns / fewer rows) so a 15-copy layout resolves to 5 cols x 3 rows rather than 3 cols x 5 rows when both are equally tight and square.
  - Centering alone wasn't enough — the corner cell still overlapped the bed's reserved exclusion rectangle at the plate origin (read from `Bambu Lab X1 Carbon 0.4 nozzle.json`: `bed_exclude_area` = (0,0)-(18,28) on a 256x256 `printable_area`, i.e. the X1C wiper/handle cutout). Added an `$EXCLUDE_W`/`$EXCLUDE_H` (18 x 28mm) check after centering that nudges the whole grid block along whichever axis requires the smaller shift to clear the corner rectangle, instead of uniformly inflating `$MARGIN` (which would waste space on the other three sides where there's no exclusion zone).
  - Also changed the BOD pipeline in `CardQueueEditorWPF.ps1` (~line 1493-1498) to only delete the temp `*_BOD.3mf` when the slice succeeds and the gcode is moved — on failure it's left in the design folder so Hank can open it in Bambu Studio and inspect the layout.
  - **Y-axis collisions despite uniform `$TARGET_GAP`:** Hank reported copies looked right in X but were touching in Y. Root cause: the design footprint (`$bboxW`/`$bboxD`) was being read from `plate_1.json`'s `bbox_objects` entry, which was *stale* — left over from before `isolate_final_worker.ps1` repositioned/regrouped this 81-part "colorcut" assembly (BlackKnight). Verified by computing the real footprint from mesh geometry: stale bbox said 22.07 x 33.10mm centred at (128.00, 126.16), but the true assembled footprint (walking all build items -> objects -> components -> meshes through their transform chains) is 23.48 x 36.78mm centred at (128.00, 128.00) — the Y error (3.68mm) is much larger than the X error (1.41mm), matching exactly what Hank saw. Fixed by adding geometry helpers (`Combine-Tx`, `Apply-Tx`, `Get-MeshVertices`, `Walk-FootprintObject`) that recursively compose 3MF transform matrices and union the transformed vertex bounds across every build item/component/mesh, replacing the `plate_1.json` bbox lookup entirely with a geometry-derived footprint that's always current regardless of how the 3mf was last saved.

**What was done:**
- Built a full Purge Dictionary editor panel inside CardQueueEditorWPF's Libraries section: sortable/filterable grid (From/Tuned/Tuned Volume filters), Tuned checkbox replacing old Iterations/Test_Volume columns, color swatches on From/To, centered numeric columns, dirty-cell highlighting (orange) with Save Changes gating, legible column headers.
- Fixed a CSV regression where PurgeDictionary.csv had been converted to tab-delimited without a BOM (broke Excel display); restored comma-delimited + UTF-8 BOM in `Save-PurgeDictionary`/`Load-PurgeDictionary`, with auto-migration of any tab-delimited copies on load. Reverted the matching `-Delimiter` hack in `UpdatePurgeMatrix_worker.ps1`.
- Made the three Purge Dictionary filter dropdowns typable/searchable (`Enable-PurgeComboTypeAhead` helper), mirroring the filament color-picker combo pattern: IsEditable + CollectionView text filtering + Enter/Tab/Escape handling.
- Added a "% Savings" column (Base vs Tuned volume) and a rolling "Avg Savings (tuned): X.X% across N combos" label at the top of the panel, computed across tuned rows only.
- **Integrated purge-matrix tuning into the "Save Colors" task** so it no longer needs a separate run of `UpdatePurgeMatrix_worker.ps1`: added `Get-PurgeNameByHexMap`, `Get-PurgeTunedVolume`, and `Update-PurgeMatrixInConfigText` helpers (near `Invoke-ColorPatchAllFiles`) that rewrite `flush_volumes_matrix` entries in `project_settings.config` JSON using tuned rows from `$script:PurgeDict`. Wired into the anchor-file substitution loop in `Start-NextProcess` and into `Invoke-ColorPatchAllFiles` (now takes a `$NameByHex` param) for all other .3mf files in the design folder — so saving colors also writes tuned purge volumes everywhere colors get patched.

**Gotcha worth remembering:** `.GetNewClosure()` scriptblocks created *inside a function* silently resolve `$script:`-scoped variables to `$null` (and `$null.Count` quietly returns `0` rather than erroring) — always capture a direct local reference to script-scoped collections (e.g. `$capturedPurgeDict = $script:PurgeDict`) before building closures that need them, same pattern as `$capturedPurgeView`.

**Next up:**
- Re-run BOD creation on the BlackKnight folder and confirm in the Bambu Studio plate preview that copies no longer touch in Y (and still look right in X), then confirm it slices and the gcode lands in the dated Printing Queue folder.
- Smoke-test the integrated purge tuning by running "Save Colors" on a real design folder with tuned PurgeDictionary entries, and confirm `flush_volumes_matrix` values update correctly in both the anchor .3mf and other .3mf files in the folder (can then retire standalone `UpdatePurgeMatrix_worker.ps1` runs from the workflow).

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
