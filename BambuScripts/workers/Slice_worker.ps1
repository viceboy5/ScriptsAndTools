param(
    [string]$InputPath = "",
    [string[]]$InputPaths = @(),
    [string]$IsolatedPath = "",
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe",
    [string]$StatusFile = ""
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

# Write a "SLICING... XX%" progress line to the shared status file.
# PhaseStart/PhaseEnd let the caller map this file's 0-100 into a slice of overall progress.
function Write-SliceProgress([string]$sf, [int]$phasePct, [int]$phaseStart, [int]$phaseEnd) {
    if ($sf -eq "") { return }
    $overall = $phaseStart + [int](($phaseEnd - $phaseStart) * $phasePct / 100)
    try { Set-Content -Path $sf -Value "SLICING... $overall%" -Force -ErrorAction SilentlyContinue } catch {}
}

function Invoke-SliceFile([string]$filePath, [string]$label, [string]$sf, [int]$phaseStart, [int]$phaseEnd) {
    $workDir   = Split-Path $filePath -Parent
    $baseName  = (Get-Item $filePath).BaseName
    $slicedOut = Join-Path $workDir "$baseName.gcode.3mf"
    $logOut    = Join-Path $env:TEMP "slice_log_out_$baseName.txt"
    $logErr    = Join-Path $env:TEMP "slice_log_err_$baseName.txt"

    Write-Host "  -> Slicing $label`: $baseName " -ForegroundColor Cyan -NoNewline

    # Typical Bambu Studio slice: ~60-120 s.  Cap time-estimate at 95 % to leave room for finalisation.
    $estSeconds = 90
    $startTime  = Get-Date
    Write-SliceProgress $sf 0 $phaseStart $phaseEnd

    # Added --uptodate and --allow-newer-file to bypass all version mismatch prompts
    $procArgs = "--debug 3 --no-check --uptodate --allow-newer-file --slice 1 --min-save --export-3mf `"$slicedOut`" `"$filePath`""
    $proc = Start-Process -FilePath $BambuPath -ArgumentList $procArgs -RedirectStandardOutput $logOut -RedirectStandardError $logErr -WindowStyle Hidden -PassThru

    while (-not $proc.HasExited) {
        # Try to parse a progress percentage from the Bambu Studio log first.
        $parsedPct = $null
        if (Test-Path $logOut) {
            try {
                $lines = Get-Content $logOut -ErrorAction SilentlyContinue | Select-Object -Last 30
                for ($li = $lines.Count - 1; $li -ge 0; $li--) {
                    $line = $lines[$li]
                    # e.g. "progress 0.450" or "progress: 0.45"
                    if ($line -match 'progress[:\s]+([0-9]*\.?[0-9]+)') {
                        $val = [double]$matches[1]
                        if ($val -le 1.0) { $parsedPct = [int]($val * 100); break }
                        elseif ($val -le 100.0) { $parsedPct = [int]$val; break }
                    }
                    # e.g. "45%" anywhere on the line
                    if ($line -match '\b([0-9]{1,3})%') {
                        $parsedPct = [int]$matches[1]; break
                    }
                }
            } catch {}
        }

        if ($null -ne $parsedPct) {
            $filePct = [Math]::Min(95, $parsedPct)
        } else {
            $elapsed  = ((Get-Date) - $startTime).TotalSeconds
            $filePct  = [Math]::Min(95, [int]($elapsed / $estSeconds * 100))
        }

        Write-SliceProgress $sf $filePct $phaseStart $phaseEnd
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

# --- Determine phase split between Full and Isolated slices ---
$hasIsolated = ($IsolatedPath -ne "" -and (Test-Path $IsolatedPath))
$totalFiles  = $allInputs.Count + [int]$hasIsolated
# Allocate phase ranges proportionally across all files (0-100 overall)
$phaseSize   = if ($totalFiles -gt 0) { [int](100 / $totalFiles) } else { 100 }

# --- Slice all input files ---
$total = $allInputs.Count
for ($i = 0; $i -lt $total; $i++) {
    $label      = if ($total -gt 1) { "[$($i+1)/$total]" } else { "Full Plate" }
    $pStart     = $i * $phaseSize
    $pEnd       = if ($i -lt $total - 1 -or $hasIsolated) { ($i + 1) * $phaseSize } else { 100 }
    Invoke-SliceFile $allInputs[$i] $label $StatusFile $pStart $pEnd | Out-Null
}

# --- Optionally slice the isolated Final file (Master-Controller path) ---
if ($hasIsolated) {
    $pStart = $total * $phaseSize
    Invoke-SliceFile $IsolatedPath "Isolated Object" $StatusFile $pStart 100 | Out-Null
}

Write-Host "  -> Slicing Automation Complete!`n" -ForegroundColor Green
exit 0