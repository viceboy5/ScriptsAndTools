# Normalize-DataTsvs.ps1
# ---------------------------------------------------------------------------
# One-time cleanup script: upgrades all *_Data.tsv files under $DesignsRoot
# to the current v3 format (29 cols, Date at col 5, SKU at col 3).
#
# Three upgrade paths:
#
#   OLD (28 cols, Date@4)
#       -> Insert empty SKU column at index 3.  All other data is preserved.
#
#   VERY-OLD (18 cols, Date@2)
#       -> Rebuild header (Printer/FileType/DesignName/SKU/Theme) from the
#          folder hierarchy.  Pad filament slots 5-8 with empty columns.
#          Replace any Excel formula in TotalSlotGrams with a computed value.
#          Requires TSV to be inside a proper 3-level design folder structure.
#
#   CURRENT (29 cols, Date@5)  -- already correct, skipped.
#   STUB (4 cols, cols 0-2 empty)  -- SKU placeholder only, skipped.
#   VERY-OLD outside proper folder structure  -- skipped, reported separately.
#
# SAFETY:
#   - Default mode is DRY RUN.  Nothing is written until you add -Apply.
#   - With -Apply, each modified file gets a .bak backup first (unless -NoBackup).
#   - Backup files are written alongside the originals as <filename>.bak.
#
# Usage:
#   # Preview all changes (no files modified):
#   .\Normalize-DataTsvs.ps1
#
#   # Preview a single theme:
#   .\Normalize-DataTsvs.ps1 -DesignsRoot "C:\ZB_Designs\Christmas25"
#
#   # Apply changes with backups:
#   .\Normalize-DataTsvs.ps1 -Apply
#
#   # Apply changes without creating backup files:
#   .\Normalize-DataTsvs.ps1 -Apply -NoBackup
# ---------------------------------------------------------------------------

param(
    [string]$DesignsRoot = "C:\ZB_Designs",
    [switch]$Apply,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Patterns and constants
# ---------------------------------------------------------------------------
$datePat     = '^\d{1,2}/\d{1,2}/\d{4}$'
$skuPat      = '^[^\s]{3,}$'
$TOTAL_COLS  = 29
$SLOT_PAIRS  = 8   # current format has 8 filament slots

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Detect-Format([string[]]$cols) {
    if ($cols.Count -eq 4 -and
        [string]::IsNullOrWhiteSpace($cols[0]) -and
        [string]::IsNullOrWhiteSpace($cols[1]) -and
        [string]::IsNullOrWhiteSpace($cols[2])) { return 'STUB' }
    if ($cols.Count -ge 6 -and $cols[5] -match $datePat)  { return 'CURRENT'  }
    if ($cols.Count -ge 5 -and $cols[4] -match $datePat)  { return 'OLD'      }
    if ($cols.Count -ge 3 -and $cols[2] -match $datePat)  { return 'VERY-OLD' }
    return 'UNKNOWN'
}

# Converts OLD (28 cols) -> CURRENT (29 cols) by inserting empty SKU at col 3.
function Convert-OldToCurrent([string[]]$cols) {
    return $cols[0..2] + @('') + $cols[3..($cols.Count - 1)]
}

# Converts VERY-OLD (18 cols) -> CURRENT (29 cols).
# $tsvPath = full path of the TSV file (used to derive folder context).
function Convert-VeryOldToCurrent([string[]]$cols, [string]$tsvPath) {
    # Derive header values from folder hierarchy:
    #   DesignFolder  = parent of TSV   e.g. P2S_AbeLincoln_StarsAndStripes
    #   TypeFolder    = grandparent     e.g. P2S_Standard_StarsAndStripes
    $designFolder = Split-Path $tsvPath -Parent
    $typeFolder   = Split-Path $designFolder -Parent

    $designLeaf = Split-Path $designFolder -Leaf
    $typeLeaf   = Split-Path $typeFolder   -Leaf

    $typeParts  = $typeLeaf -split '_'
    $printer    = $typeParts[0]
    $theme      = $typeParts[-1]
    $fileType   = if ($typeParts.Count -ge 3) { $typeParts[1..($typeParts.Count-2)] -join '_' } else { 'Standard' }

    $designIdx  = $designLeaf.IndexOf('_')
    $designName = if ($designIdx -ge 0) { $designLeaf.Substring($designIdx + 1) } else { $designLeaf }

    # Preserve date / H / M from old cols 2-4
    $date = $cols[2]
    $h    = $cols[3]
    $m    = $cols[4]

    # Old format had 4 filament slots (cols 5..12 = 8 values = 4 grams+color pairs)
    $oldSlots = $cols[5..12]   # exactly 8 values

    # Pad to 8 slot pairs: add 4 more empty pairs (8 empty strings)
    $paddedSlots = $oldSlots + @('','','','','','','','')

    # Tail stats from old format:
    #   old[13] = ColorSwaps
    #   old[14] = ObjCount
    #   old[15] = ModelGrams
    #   old[16] = TotalSlotGrams  (may contain =SUM formula -- replace with computed value)
    #   old[17] = TimeAdd
    $colorSwaps = $cols[13]
    $objCount   = $cols[14]
    $modelGrams = $cols[15]
    $timeAdd    = $cols[17]

    # Compute TotalSlotGrams from actual slot data (replaces any Excel formula)
    $computedTotal = 0.0
    for ($si = 0; $si -le 6; $si += 2) {
        $g = 0.0
        if ([double]::TryParse($oldSlots[$si], [ref]$g)) { $computedTotal += $g }
    }
    $totalSlotGrams = [math]::Round($computedTotal, 2)

    # Build the new 29-col row
    return @($printer, $fileType, $designName, '', $theme, $date, $h, $m) +
           $paddedSlots +
           @($colorSwaps, $objCount, $modelGrams, $totalSlotGrams, $timeAdd)
}

# Returns $true if the TSV file is inside a valid 3-level design folder structure.
# Structure required: ThemeRoot/{P}_{T}/{P}_{FT}_{T}/{P}_{D}_{T}/file.tsv
# Validated by confirming the type folder (grandparent) has at least 2 underscores.
function Test-ValidDesignFolder([string]$tsvPath) {
    $typeFolderLeaf = Split-Path (Split-Path (Split-Path $tsvPath -Parent) -Parent) -Leaf
    return ($typeFolderLeaf -split '_').Count -ge 3
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if (-not (Test-Path $DesignsRoot)) {
    Write-Host "[!] DesignsRoot not found: $DesignsRoot" -ForegroundColor Red
    exit 1
}

$mode = if ($Apply) { "APPLY" } else { "DRY RUN" }
Write-Host ""
Write-Host "Normalize-DataTsvs  --  $mode" -ForegroundColor Cyan
Write-Host "Scanning: $DesignsRoot"
Write-Host ""

# Gather all TSV files
$allTsvs = @(Get-ChildItem $DesignsRoot -Recurse -Filter "*_Data.tsv" -ErrorAction SilentlyContinue)
Write-Host "Found $($allTsvs.Count) *_Data.tsv files.  Classifying..." -ForegroundColor DarkCyan
Write-Host ""

# Classify
$toConvertOld     = [System.Collections.Generic.List[string]]::new()
$toConvertVeryOld = [System.Collections.Generic.List[string]]::new()
$skippedCurrent   = 0
$skippedStub      = 0
$skippedBadPath   = [System.Collections.Generic.List[string]]::new()
$skippedUnknown   = [System.Collections.Generic.List[string]]::new()

foreach ($tsv in $allTsvs) {
    $line = Get-Content $tsv.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
    if ([string]::IsNullOrWhiteSpace($line)) { $skippedUnknown.Add($tsv.FullName); continue }

    $cols = $line -split "`t"
    $fmt  = Detect-Format $cols

    switch ($fmt) {
        'CURRENT'  { $skippedCurrent++ }
        'STUB'     { $skippedStub++ }
        'OLD'      { $toConvertOld.Add($tsv.FullName) }
        'VERY-OLD' {
            if (Test-ValidDesignFolder $tsv.FullName) {
                $toConvertVeryOld.Add($tsv.FullName)
            } else {
                $skippedBadPath.Add($tsv.FullName)
            }
        }
        default    { $skippedUnknown.Add($tsv.FullName) }
    }
}

# ---------------------------------------------------------------------------
# Report what will change
# ---------------------------------------------------------------------------
Write-Host "SUMMARY" -ForegroundColor White
Write-Host "-------" -ForegroundColor White
Write-Host "  Already current    : $skippedCurrent" -ForegroundColor Green
Write-Host "  Stub (SKU-only)    : $skippedStub  (skipped, no slice data)" -ForegroundColor DarkGray
Write-Host "  OLD -> CURRENT     : $($toConvertOld.Count)  (insert empty SKU column)" -ForegroundColor Yellow
Write-Host "  VERY-OLD -> CURRENT: $($toConvertVeryOld.Count)  (rebuild header + pad slots)" -ForegroundColor Yellow
if ($skippedBadPath.Count -gt 0) {
    Write-Host "  VERY-OLD, bad path : $($skippedBadPath.Count)  (skipped -- not in 3-level folder structure)" -ForegroundColor Red
}
if ($skippedUnknown.Count -gt 0) {
    Write-Host "  Unknown format     : $($skippedUnknown.Count)  (skipped)" -ForegroundColor Red
}
Write-Host ""

# Show OLD conversions (brief)
if ($toConvertOld.Count -gt 0) {
    Write-Host "OLD files to upgrade ($($toConvertOld.Count)):" -ForegroundColor Yellow
    foreach ($f in $toConvertOld) {
        $line = Get-Content $f | Select-Object -Last 1
        $cols = $line -split "`t"
        $label = "  $($cols[0]) | $($cols[1]) | $($cols[2]) | [SKU->empty] | $($cols[3]) | $($cols[4])"
        Write-Host $label -ForegroundColor DarkYellow
    }
    Write-Host ""
}

# Show VERY-OLD conversions
if ($toConvertVeryOld.Count -gt 0) {
    Write-Host "VERY-OLD files to upgrade ($($toConvertVeryOld.Count)):" -ForegroundColor Yellow
    foreach ($f in $toConvertVeryOld) {
        $designLeaf = Split-Path (Split-Path $f -Parent) -Leaf
        $typeLeaf   = Split-Path (Split-Path (Split-Path $f -Parent) -Parent) -Leaf
        $typeParts  = $typeLeaf -split '_'
        $printer    = $typeParts[0]
        $theme      = $typeParts[-1]
        $ftParts    = if ($typeParts.Count -ge 3) { $typeParts[1..($typeParts.Count-2)] -join '_' } else { 'Standard' }
        $dIdx       = $designLeaf.IndexOf('_')
        $dName      = if ($dIdx -ge 0) { $designLeaf.Substring($dIdx + 1) } else { $designLeaf }
        Write-Host "  $printer | $ftParts | $dName | [SKU->empty] | $theme  <--  $(Split-Path $f -Leaf)" -ForegroundColor DarkYellow
    }
    Write-Host ""
}

# Skipped bad-path files
if ($skippedBadPath.Count -gt 0) {
    Write-Host "VERY-OLD files SKIPPED (not in 3-level folder structure):" -ForegroundColor Red
    foreach ($f in $skippedBadPath) {
        Write-Host "  $($f.Replace($DesignsRoot,''))" -ForegroundColor DarkRed
    }
    Write-Host "  -> These need manual review and folder restructuring." -ForegroundColor DarkRed
    Write-Host ""
}

# Unknown-format files
if ($skippedUnknown.Count -gt 0) {
    Write-Host "Files with UNKNOWN format (skipped):" -ForegroundColor Red
    foreach ($f in $skippedUnknown) {
        Write-Host "  $($f.Replace($DesignsRoot,''))" -ForegroundColor DarkRed
    }
    Write-Host ""
}

$totalToFix = $toConvertOld.Count + $toConvertVeryOld.Count

if ($totalToFix -eq 0) {
    Write-Host "Nothing to do.  All eligible files are already in the current format." -ForegroundColor Green
    exit 0
}

if (-not $Apply) {
    Write-Host "DRY RUN complete.  $totalToFix file(s) would be upgraded." -ForegroundColor Cyan
    Write-Host "Run with -Apply to write changes.  Backups will be created unless -NoBackup is also specified." -ForegroundColor Cyan
    exit 0
}

# ---------------------------------------------------------------------------
# Apply changes
# ---------------------------------------------------------------------------
Write-Host "Applying changes..." -ForegroundColor Cyan
$successOld = 0; $successVeryOld = 0; $errors = 0

foreach ($f in $toConvertOld) {
    try {
        $line    = Get-Content $f -Raw -ErrorAction Stop
        $cols    = $line.TrimEnd("`r","`n") -split "`t"
        $newCols = Convert-OldToCurrent $cols

        if (-not $NoBackup) { Copy-Item $f "$f.bak" -Force }
        Set-Content -Path $f -Value ($newCols -join "`t") -Encoding UTF8 -NoNewline
        Write-Host "  [OLD->v3]  $(Split-Path $f -Leaf)" -ForegroundColor Green
        $successOld++
    } catch {
        Write-Host "  [ERROR]    $(Split-Path $f -Leaf) -- $_" -ForegroundColor Red
        $errors++
    }
}

foreach ($f in $toConvertVeryOld) {
    try {
        $line    = Get-Content $f -Raw -ErrorAction Stop
        $cols    = $line.TrimEnd("`r","`n") -split "`t"
        $newCols = Convert-VeryOldToCurrent $cols $f

        if (-not $NoBackup) { Copy-Item $f "$f.bak" -Force }
        Set-Content -Path $f -Value ($newCols -join "`t") -Encoding UTF8 -NoNewline
        Write-Host "  [v1->v3]   $(Split-Path $f -Leaf)" -ForegroundColor Green
        $successVeryOld++
    } catch {
        Write-Host "  [ERROR]    $(Split-Path $f -Leaf) -- $_" -ForegroundColor Red
        $errors++
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Write-Host "  OLD -> v3  : $successOld upgraded" -ForegroundColor Green
Write-Host "  v1  -> v3  : $successVeryOld upgraded" -ForegroundColor Green
if ($errors -gt 0) {
    Write-Host "  Errors     : $errors  (original files unchanged where error occurred)" -ForegroundColor Red
}
if (-not $NoBackup -and ($successOld + $successVeryOld) -gt 0) {
    Write-Host "  Backups    : .bak files created alongside each modified file" -ForegroundColor DarkGray
}
