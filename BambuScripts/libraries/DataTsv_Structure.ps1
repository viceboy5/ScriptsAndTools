# DataTsv_Structure.ps1
# ---------------------------------------------------------------------------
# Reference definition for the Data TSV format used across all BambuScripts.
# Dot-source this file to get column-index constants, detection helpers,
# and format documentation in one place.
#
# Usage:
#   . "$PSScriptRoot\..\libraries\DataTsv_Structure.ps1"
#   $cols[$TSV_SKU]          # read SKU from a split line
#   Get-TsvFormat $cols      # returns "CURRENT" | "OLD" | "VERY-OLD" | "STUB" | "UNKNOWN"
# ---------------------------------------------------------------------------


# ===========================================================================
# CURRENT FORMAT  (v3)  --  29 columns  --  Date is at index 5
# ===========================================================================
#
#  Idx  Name              Example / Notes
#  ---  ----------------  --------------------------------------------------
#   0   Printer           X1C | P2S | P1S | A1
#   1   FileType          Standard | Multicolor | Big Wig | ...
#   2   DesignName        AbeLincoln_StarsAndStripes
#   3   SKU               sns-sw-2608  (may be empty string)
#   4   Theme             StarsAndStripes
#   5   Date              M/d/yyyy  ->  4/27/2026
#   6   H                 print hours  (integer)
#   7   M                 print minutes (integer)
#   8   Slot1Grams        grams from AMS slot 1 (empty if slot unused)
#   9   Slot1Color        filament label  e.g. "Esun Cold White"
#  10   Slot2Grams
#  11   Slot2Color
#  12   Slot3Grams
#  13   Slot3Color
#  14   Slot4Grams
#  15   Slot4Color
#  16   Slot5Grams
#  17   Slot5Color
#  18   Slot6Grams
#  19   Slot6Color
#  20   Slot7Grams
#  21   Slot7Color
#  22   Slot8Grams
#  23   Slot8Color
#  24   ColorSwaps        filament-change count during print
#  25   ObjCount          number of objects in the slice
#  26   ModelGrams        model-only filament weight (no supports/waste)
#  27   TotalSlotGrams    sum of all slot grams (Slot1..Slot8)
#  28   TimeAdd           per-object time delta used for multi-object pricing
#
#  Detection: cols[5] matches date pattern  AND  cols.Count >= 29
#
# ---------------------------------------------------------------------------
# Column index constants  (current format)
# ---------------------------------------------------------------------------
$TSV_PRINTER         = 0
$TSV_FILETYPE        = 1
$TSV_DESIGNNAME      = 2
$TSV_SKU             = 3
$TSV_THEME           = 4
$TSV_DATE            = 5
$TSV_H               = 6
$TSV_M               = 7
$TSV_SLOT_START      = 8   # first slot column (Slot1Grams)
$TSV_SLOT_COUNT      = 8   # total number of filament slots
$TSV_COLOR_SWAPS     = 24
$TSV_OBJ_COUNT       = 25
$TSV_MODEL_GRAMS     = 26
$TSV_TOTAL_SLOT_GRAMS = 27
$TSV_TIME_ADD        = 28
$TSV_TOTAL_COLS      = 29

# Slot helper: given a 1-based slot number (1..8), returns the grams col index
function Get-TsvSlotGramsIndex([int]$slot) { 6 + ($slot * 2) }
# Slot helper: given a 1-based slot number (1..8), returns the color col index
function Get-TsvSlotColorIndex([int]$slot) { 7 + ($slot * 2) }


# ===========================================================================
# HISTORICAL FORMAT: OLD  (v2)  --  28 columns  --  Date at index 4
# ===========================================================================
#
#  Idx  Name              Notes
#  ---  ----------------  --------------------------------------------------
#   0   Printer
#   1   FileType
#   2   DesignName
#   3   Theme             (no SKU column -- Theme is one position earlier)
#   4   Date
#   5   H
#   6   M
#   7   Slot1Grams
#   8   Slot1Color
#   ... (same 8-slot layout as current, shifted -1)
#  22   Slot8Grams
#  23   Slot8Color
#  24   ColorSwaps        (was col 23 -- no SKU shift)
#
#  Wait -- old format still has 8 slots.  Shift from current: all cols
#  after DesignName are -1 because SKU column is missing.
#
#  Detection: cols[4] matches date pattern  AND  cols.Count == 28
#  Fix: insert '' at index 3  ->  28 cols become 29 cols (current)
#
# ===========================================================================
# HISTORICAL FORMAT: VERY-OLD  (v1)  --  18 columns  --  Date at index 2
# ===========================================================================
#
#  Idx  Name              Notes
#  ---  ----------------  -----------------------------------------------
#   0   OldName           legacy internal name, e.g. X1C_Bunny_Artic_Christmas2025
#   1   OldTheme          legacy theme label, e.g. Christmas2025 (may differ from folder)
#   2   Date
#   3   H
#   4   M
#   5   Slot1Grams
#   6   Slot1Color
#   7   Slot2Grams
#   8   Slot2Color
#   9   Slot3Grams
#  10   Slot3Color
#  11   Slot4Grams
#  12   Slot4Color        (only 4 slots in this version)
#  13   ColorSwaps
#  14   ObjCount
#  15   ModelGrams
#  16   TotalSlotGrams    often contains an Excel =SUM(INDIRECT(...)) formula
#  17   TimeAdd
#
#  Detection: cols[2] matches date pattern  AND  cols.Count == 18
#             AND folder leaf starts with a known printer prefix
#  Fix: rebuild header from folder hierarchy; pad slots 5-8 with empty cols
#
# ===========================================================================
# STUB  --  4 columns  --  SKU placeholder only
# ===========================================================================
#
#  [0]=""  [1]=""  [2]=""  [3]=SKU
#  Created manually before slicing data was available.
#  Do NOT attempt to reformat these; they have no slice data.
#
# ===========================================================================


# ---------------------------------------------------------------------------
# Shared patterns
# ---------------------------------------------------------------------------
$TSV_DATE_PATTERN = '^\d{1,2}/\d{1,2}/\d{4}$'
$TSV_SKU_PATTERN  = '^[^\s]{3,}$'

# Printer prefixes expected as the first underscore-delimited token in folder names
$TSV_KNOWN_PRINTERS = @('X1C', 'P2S', 'P1S', 'A1', 'A1M', 'P1P', 'H2D')


# ---------------------------------------------------------------------------
# Get-TsvFormat
# Given a split-tab array of columns, returns the format name.
# Returns: "CURRENT" | "OLD" | "VERY-OLD" | "STUB" | "UNKNOWN"
# ---------------------------------------------------------------------------
function Get-TsvFormat([string[]]$cols) {
    if ($null -eq $cols -or $cols.Count -eq 0) { return 'UNKNOWN' }

    # STUB: 4 cols, first three empty, last is SKU
    if ($cols.Count -eq 4 -and
        [string]::IsNullOrWhiteSpace($cols[0]) -and
        [string]::IsNullOrWhiteSpace($cols[1]) -and
        [string]::IsNullOrWhiteSpace($cols[2])) {
        return 'STUB'
    }

    # CURRENT: Date at col 5
    if ($cols.Count -ge 6 -and $cols[5] -match $TSV_DATE_PATTERN) { return 'CURRENT' }

    # OLD: Date at col 4
    if ($cols.Count -ge 5 -and $cols[4] -match $TSV_DATE_PATTERN) { return 'OLD' }

    # VERY-OLD: Date at col 2
    if ($cols.Count -ge 3 -and $cols[2] -match $TSV_DATE_PATTERN) { return 'VERY-OLD' }

    return 'UNKNOWN'
}


# ---------------------------------------------------------------------------
# Get-TsvDesignName
# Extracts design name from a split-tab column array, with fallback to
# folder path when col[2] holds a date (VERY-OLD format).
# $folderPath = the design folder (parent directory of the TSV file)
# ---------------------------------------------------------------------------
function Get-TsvDesignName([string[]]$cols, [string]$folderPath) {
    if ($null -ne $cols -and $cols.Count -ge 3 -and
        -not [string]::IsNullOrWhiteSpace($cols[2]) -and
        $cols[2].Trim() -notmatch $TSV_DATE_PATTERN) {
        return $cols[2].Trim()
    }
    # Fallback: strip printer prefix from folder leaf
    $leaf = Split-Path $folderPath -Leaf
    $idx  = $leaf.IndexOf('_')
    return if ($idx -ge 0) { $leaf.Substring($idx + 1) } else { $leaf }
}
