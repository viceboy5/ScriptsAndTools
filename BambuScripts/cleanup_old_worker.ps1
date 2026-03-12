param(
    [string]$TargetDir
)
$ErrorActionPreference = 'SilentlyContinue'

if (-not (Test-Path $TargetDir)) { exit 0 }

# Find all files/folders containing 'old' in the name, OR any file with a .txt extension
$items = Get-ChildItem -Path $TargetDir -Recurse | Where-Object {
    $_.Name -match '(?i)old' -or $_.Extension -match '(?i)^\.txt$'
} | Sort-Object -Property @{Expression={$_.FullName.Length}; Ascending=$true}

if ($items.Count -eq 0) { exit 0 }

# Condense the list so we don't spam the console with files that are inside a folder we are already listing
$roots = @()
foreach ($item in $items) {
    $isChild = $false
    foreach ($root in $roots) {
        if ($item.FullName.StartsWith($root.FullName + "\")) {
            $isChild = $true
            break
        }
    }
    if (-not $isChild) { $roots += $item }
}

Write-Host "`n==============================================================" -ForegroundColor Magenta
Write-Host " PRE-FLIGHT CLEANUP: 'OLD' OR '.TXT' ITEMS DETECTED" -ForegroundColor Magenta
Write-Host "==============================================================" -ForegroundColor Magenta

foreach ($root in $roots) {
    $type = if ($root.PSIsContainer) { "[FOLDER]" } else { "[FILE]  " }
    Write-Host "  $type $($root.Name)" -ForegroundColor Yellow
}

Write-Host ""
$ans = Read-Host "Do you want to permanently delete these items? (Y/N)"
if ($ans -match '(?i)^y') {
    foreach ($root in $roots) {
        Remove-Item -Path $root.FullName -Recurse -Force
    }
    Write-Host "  [+] Items successfully deleted." -ForegroundColor Green
} else {
    Write-Host "  [-] Deletion skipped." -ForegroundColor DarkGray
}
Write-Host ""
exit 0