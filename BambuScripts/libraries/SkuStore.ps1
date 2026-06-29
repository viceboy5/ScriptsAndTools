# --- Durable per-design SKU store --------------------------------------------
# The SKU is the ONE piece of business data in a design folder that cannot be
# regenerated (everything else - gcode, images, the _Data.tsv metrics - is
# derived and can be rebuilt). Historically the SKU lived in column 3 of the
# regenerable <design>_Data.tsv, so the re-nester / merge / slice pipeline (which
# rewrites or deletes that file) could wipe it.
#
# This module stores the SKU in a separate sidecar file INSIDE the design folder:
#   - The pipeline never touches it, so re-nesting can't lose it.
#   - It is keyed by physical location, not by name, so renaming the design
#     (the editor renames the folder in place) carries the SKU along with zero
#     key maintenance.
#   - DataExtract_worker reads it at extraction time and stamps it into the TSV.
#
# The canonical source of truth is still upstream (the web DB -> CSV import ->
# export to Google Sheets); this sidecar is a durable LOCAL cache so a SKU
# assigned in the editor survives until it is exported, regardless of how many
# times the design is re-nested or renamed in between.
#
# Dot-source from any worker that reads/writes SKUs:
#   . (Join-Path $PSScriptRoot "..\libraries\SkuStore.ps1")

# Shared SKU validation pattern (3+ non-whitespace chars). Defined only if the
# host script hasn't already set it, so CardQueueEditor's existing value wins.
if (-not $script:SkuPattern) { $script:SkuPattern = '^[^\s]{3,}$' }

# Fixed sidecar filename. Fixed (not design-named) so it is rename-proof and the
# editor's per-file rename pass leaves it alone.
$script:SkuSidecarName = 'SKU.txt'

function Get-SkuSidecarPath([string]$folderPath) {
    if ([string]::IsNullOrWhiteSpace($folderPath)) { return $null }
    return (Join-Path $folderPath $script:SkuSidecarName)
}

# Returns the stored SKU for a design folder, or "" if none / invalid.
function Read-SkuSidecar([string]$folderPath) {
    $p = Get-SkuSidecarPath $folderPath
    if ($null -eq $p -or -not (Test-Path -LiteralPath $p)) { return "" }
    try {
        foreach ($line in (Get-Content -LiteralPath $p -ErrorAction Stop)) {
            $v = "$line".Trim()
            if ($v -ne "" -and $v -match $script:SkuPattern) { return $v }
        }
    } catch {}
    return ""
}

# Writes the SKU sidecar for a design folder. Refuses to write junk/empty values
# (so it never clobbers a good SKU with a blank), and marks the file read-only so
# a casual delete/edit in Explorer prompts a warning. Returns $true on success.
function Write-SkuSidecar([string]$folderPath, [string]$sku) {
    $p = Get-SkuSidecarPath $folderPath
    if ($null -eq $p) { return $false }
    $val = "$sku".Trim()
    if ($val -eq "" -or $val -notmatch $script:SkuPattern) { return $false }
    try {
        # Clear read-only first so an existing sidecar can be updated.
        if (Test-Path -LiteralPath $p) { try { (Get-Item -LiteralPath $p -Force).IsReadOnly = $false } catch {} }
        Set-Content -LiteralPath $p -Value $val -Encoding UTF8 -Force
        try { (Get-Item -LiteralPath $p -Force).IsReadOnly = $true } catch {}
        return $true
    } catch {
        return $false
    }
}
