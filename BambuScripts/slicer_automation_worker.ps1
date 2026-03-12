param(
    [string]$InputPath,
    [string]$IsolatedPath = "",
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe"
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $BambuPath)) {
    Write-Host "  [!] ERROR: Bambu Studio not found at: $BambuPath" -ForegroundColor Red
    exit 1
}

$workDir = Split-Path $InputPath -Parent
$baseName = (Get-Item $InputPath).BaseName
$slicedOut = Join-Path $workDir "$baseName.gcode.3mf"

# Delay to ensure network drives/Synology drop their locks
Start-Sleep -Seconds 3

Write-Host "  -> Slicing Full Plate: $baseName" -ForegroundColor Cyan
$proc1 = Start-Process -FilePath $BambuPath -ArgumentList "--debug 3 --no-check --slice 1 --min-save --export-3mf `"$slicedOut`" `"$InputPath`"" -Wait -NoNewWindow -PassThru

if (-not (Test-Path $slicedOut)) {
    Write-Host "  [!] ERROR: Bambu Studio failed to generate $slicedOut" -ForegroundColor Red
    exit 1
}

if ($IsolatedPath -ne "" -and (Test-Path $IsolatedPath)) {
    $isoBase = (Get-Item $IsolatedPath).BaseName
    $isoSlicedOut = Join-Path $workDir "$isoBase.gcode.3mf"

    Write-Host "  -> Slicing Isolated Object: $isoBase" -ForegroundColor Cyan
    $proc2 = Start-Process -FilePath $BambuPath -ArgumentList "--debug 3 --no-check --slice 1 --min-save --export-3mf `"$isoSlicedOut`" `"$IsolatedPath`"" -Wait -NoNewWindow -PassThru

    if (-not (Test-Path $isoSlicedOut)) {
        Write-Host "  [!] WARNING: Isolated object failed to slice." -ForegroundColor Yellow
    }
}

Write-Host "  -> Slicing Automation Complete!" -ForegroundColor Green
exit 0