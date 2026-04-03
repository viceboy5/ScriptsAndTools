param(
    [string]$InputPath = "",
    [string[]]$InputPaths = @(),
    [string]$IsolatedPath = "",
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe"
)
$ErrorActionPreference = 'Stop'

# Build the full list of files to slice.
# Master-Controller passes -InputPath (single file) - that path is honoured as-is.
# Standalone / Slice.bat passes -InputPaths (array) for multi-file drag-drop.
$allInputs = [System.Collections.Generic.List[string]]::new()
if ($InputPath -ne "") { $allInputs.Add($InputPath) }
foreach ($p in $InputPaths) { if ($p -ne "") { $allInputs.Add($p) } }

if ($allInputs.Count -eq 0) {
    Write-Host "  [!] ERROR: No input file specified." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $BambuPath)) {
    Write-Host "  [!] ERROR: Bambu Studio not found at: $BambuPath" -ForegroundColor Red
    exit 1
}

function Invoke-SliceFile([string]$filePath, [string]$label) {
    $workDir   = Split-Path $filePath -Parent
    $baseName  = (Get-Item $filePath).BaseName
    $slicedOut = Join-Path $workDir "$baseName.gcode.3mf"
    $logOut    = Join-Path $env:TEMP "slice_log_out_$baseName.txt"
    $logErr    = Join-Path $env:TEMP "slice_log_err_$baseName.txt"

    Write-Host "  -> Slicing $label`: $baseName " -ForegroundColor Cyan -NoNewline

    # Added --uptodate and --allow-newer-file to bypass all version mismatch prompts
    $procArgs = "--debug 3 --no-check --uptodate --allow-newer-file --slice 1 --min-save --export-3mf `"$slicedOut`" `"$filePath`""
    $proc = Start-Process -FilePath $BambuPath -ArgumentList $procArgs -RedirectStandardOutput $logOut -RedirectStandardError $logErr -WindowStyle Hidden -PassThru

    while (-not $proc.HasExited) {
        Write-Host "." -ForegroundColor Cyan -NoNewline
        Start-Sleep -Seconds 3
    }
    Write-Host " [DONE]" -ForegroundColor Green

    if (-not (Test-Path $slicedOut)) {
        Write-Host "`n  [!] ERROR: Bambu Studio failed to generate $slicedOut" -ForegroundColor Red
        if (Test-Path $logOut) { Get-Content $logOut | Select-Object -Last 10 | Write-Host -ForegroundColor DarkGray }
        if (Test-Path $logErr) { Get-Content $logErr | Select-Object -Last 10 | Write-Host -ForegroundColor DarkGray }
        if (Test-Path $logOut) { Remove-Item $logOut -Force -ErrorAction SilentlyContinue }
        if (Test-Path $logErr) { Remove-Item $logErr -Force -ErrorAction SilentlyContinue }
        return $false
    }

    if (Test-Path $logOut) { Remove-Item $logOut -Force -ErrorAction SilentlyContinue }
    if (Test-Path $logErr) { Remove-Item $logErr -Force -ErrorAction SilentlyContinue }
    return $true
}

# --- Slice all input files ---
$total = $allInputs.Count
for ($i = 0; $i -lt $total; $i++) {
    $label = if ($total -gt 1) { "[$($i+1)/$total]" } else { "Full Plate" }
    Invoke-SliceFile $allInputs[$i] $label | Out-Null
}

# --- Optionally slice the isolated Final file (Master-Controller path) ---
if ($IsolatedPath -ne "" -and (Test-Path $IsolatedPath)) {
    Invoke-SliceFile $IsolatedPath "Isolated Object" | Out-Null
}

Write-Host "  -> Slicing Automation Complete!`n" -ForegroundColor Green
exit 0