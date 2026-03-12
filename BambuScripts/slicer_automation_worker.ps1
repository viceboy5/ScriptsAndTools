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

# Create two separate log files to satisfy PowerShell's strict redirection rules
$logOut = Join-Path $env:TEMP "slice_log_out_$baseName.txt"
$logErr = Join-Path $env:TEMP "slice_log_err_$baseName.txt"

# Delay to ensure network drives/Synology drop their locks
Start-Sleep -Seconds 3

Write-Host "  -> Slicing Full Plate: $baseName" -ForegroundColor Cyan

$args1 = "--debug 3 --no-check --slice 1 --min-save --export-3mf `"$slicedOut`" `"$InputPath`""

# Native PowerShell launch: Prevents buffer freezes AND ignores ghost background threads
Start-Process -FilePath $BambuPath -ArgumentList $args1 -Wait -RedirectStandardOutput $logOut -RedirectStandardError $logErr -WindowStyle Hidden

if (-not (Test-Path $slicedOut)) {
    Write-Host "  [!] ERROR: Bambu Studio failed to generate $slicedOut" -ForegroundColor Red
    Write-Host "  ==================== BAMBU STUDIO LOG ====================" -ForegroundColor DarkGray
    if (Test-Path $logOut) { Get-Content $logOut | Select-Object -Last 10 | Write-Host -ForegroundColor DarkGray }
    if (Test-Path $logErr) { Get-Content $logErr | Select-Object -Last 10 | Write-Host -ForegroundColor DarkGray }
    Write-Host "  ==========================================================" -ForegroundColor DarkGray
    exit 1
}

if ($IsolatedPath -ne "" -and (Test-Path $IsolatedPath)) {
    $isoBase = (Get-Item $IsolatedPath).BaseName
    $isoSlicedOut = Join-Path $workDir "$isoBase.gcode.3mf"

    Write-Host "  -> Slicing Isolated Object: $isoBase" -ForegroundColor Cyan

    $args2 = "--debug 3 --no-check --slice 1 --min-save --export-3mf `"$isoSlicedOut`" `"$IsolatedPath`""
    Start-Process -FilePath $BambuPath -ArgumentList $args2 -Wait -RedirectStandardOutput $logOut -RedirectStandardError $logErr -WindowStyle Hidden

    if (-not (Test-Path $isoSlicedOut)) {
        Write-Host "  [!] WARNING: Isolated object failed to slice." -ForegroundColor Yellow
    }
}

# Clean up the log files
if (Test-Path $logOut) { Remove-Item $logOut -Force -ErrorAction SilentlyContinue }
if (Test-Path $logErr) { Remove-Item $logErr -Force -ErrorAction SilentlyContinue }

Write-Host "  -> Slicing Automation Complete!" -ForegroundColor Green
exit 0