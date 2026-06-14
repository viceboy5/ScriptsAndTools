param(
    [Parameter(Mandatory=$true)]
    [string]$ThemeRoot   # e.g. C:\ZB_Designs\Foodz
)

# Sync-DesignSKUs.ps1
# -------------------------------------------------------------------
# Scans all *_Data.tsv files under $ThemeRoot, groups them by design
# name (col 2), and propagates the SKU (col 3) from any file that has
# one to all sibling files with the same design name that don't.
#
# Works for any number of printer prefixes (X1C, P2S, P1S, A1, etc.)
# because the match key is the design name, not the printer prefix.
#
# Usage:
#   .\Sync-DesignSKUs.ps1 -ThemeRoot "C:\ZB_Designs\Foodz"
# -------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$datePattern = '^\d{1,2}/\d{1,2}/\d{4}$'
$skuPattern  = '^[\w][\w-]{1,}$'

if (-not (Test-Path $ThemeRoot)) {
    Write-Host "[!] ThemeRoot not found: $ThemeRoot" -ForegroundColor Red
    exit 1
}

# --- 1. Collect all Data TSV files and parse their current state ---
$records = [System.Collections.Generic.List[hashtable]]::new()

Get-ChildItem $ThemeRoot -Recurse -Filter "*_Data.tsv" | ForEach-Object {
    $raw  = Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($raw)) { return }
    $cols = $raw.TrimEnd("`r","`n") -split "`t"
    if ($cols.Count -lt 4) { return }

    $isOldFormat = $cols.Count -gt 4 -and $cols[4] -match $datePattern
    $designName  = $cols[2].Trim()
    $sku         = if (-not $isOldFormat -and $cols[3].Trim() -match $skuPattern) { $cols[3].Trim() } else { '' }

    $records.Add(@{
        Path        = $_.FullName
        Cols        = $cols
        IsOldFormat = $isOldFormat
        DesignName  = $designName
        SKU         = $sku
    })
}

if ($records.Count -eq 0) {
    Write-Host "No *_Data.tsv files found under: $ThemeRoot" -ForegroundColor Yellow
    exit 0
}

# --- 2. Build design-name -> SKU lookup from files that already have one ---
$skuLookup = @{}
foreach ($r in $records) {
    if ($r.SKU -ne '' -and -not $skuLookup.ContainsKey($r.DesignName)) {
        $skuLookup[$r.DesignName] = $r.SKU
    }
}

if ($skuLookup.Count -eq 0) {
    Write-Host "No SKUs found in any Data TSV under $ThemeRoot - nothing to propagate." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($skuLookup.Count) unique design SKUs. Propagating..." -ForegroundColor Cyan

# --- 3. Propagate to files that are missing the SKU ---
$updated = 0; $skipped = 0; $noMapping = 0

foreach ($r in $records) {
    if ($r.SKU -ne '') {
        # Already has SKU - nothing to do
        $skipped++
        continue
    }

    if (-not $skuLookup.ContainsKey($r.DesignName)) {
        Write-Host "  [?] No SKU known for: $($r.DesignName)  ($($r.Path))" -ForegroundColor Yellow
        $noMapping++
        continue
    }

    $sku  = $skuLookup[$r.DesignName]
    $cols = $r.Cols

    if ($r.IsOldFormat) {
        # Insert SKU column at index 3, shifting Theme onward
        $newCols = $cols[0..2] + $sku + $cols[3..($cols.Count - 1)]
        Set-Content -Path $r.Path -Value ($newCols -join "`t") -Encoding UTF8 -NoNewline
    } else {
        # New format - fill the empty SKU slot
        $cols[3] = $sku
        Set-Content -Path $r.Path -Value ($cols -join "`t") -Encoding UTF8 -NoNewline
    }

    Write-Host "  [+] $($r.DesignName)  ->  $sku  ($(Split-Path $r.Path -Leaf))" -ForegroundColor Green
    $updated++
}

Write-Host "`nDone.  Propagated: $updated  |  Already set: $skipped  |  No mapping: $noMapping" -ForegroundColor Cyan
