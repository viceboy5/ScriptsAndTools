param(
    [string]$InputPath = "",
    [string[]]$InputPaths = @(),
    [string]$IsolatedPath = "",
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe",
    [string]$StatusFile = "",
    [switch]$DiagnoseFiles   # Watch for any files Bambu creates/deletes during slicing
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

# ---------------------------------------------------------------------------
# -DiagnoseFiles helpers
# Snapshot-diff approach: no background threads, no runspaces, no delegates.
# Take a file inventory before Bambu starts, poll every tick of the existing
# progress loop to catch transient files, then diff again after exit.
# ---------------------------------------------------------------------------
function Start-FileSnapshot([string]$workDir, [string]$bambuDir) {
    $watchDirs = @(
        $workDir,
        $env:TEMP,
        "$env:APPDATA\BambuStudio",
        $bambuDir,
        "$env:LOCALAPPDATA\Temp"
    ) | Where-Object { $_ -ne "" -and (Test-Path $_) } | Select-Object -Unique

    $snapshot = @{}
    foreach ($dir in $watchDirs) {
        try {
            Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                $snapshot[$_.FullName] = $_.LastWriteTimeUtc
            }
        } catch {}
    }
    return @{ WatchDirs = $watchDirs; Snapshot = $snapshot; Seen = [ordered]@{} }
}

# Called inside the progress loop — catches files that appear and disappear
# between the before and after snapshots.
function Update-FileSnapshot($state) {
    foreach ($dir in $state.WatchDirs) {
        try {
            Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                $fp = $_.FullName
                if (-not $state.Snapshot.ContainsKey($fp) -and -not $state.Seen.Contains($fp)) {
                    $ts = (Get-Date).ToString("HH:mm:ss")
                    $state.Seen[$fp] = "[$ts] Created  $fp"
                }
            }
        } catch {}
    }
}

# Called after process exits — final diff merged with any transient hits
function Stop-FileSnapshot($state) {
    $events = [System.Collections.Generic.List[string]]::new()
    foreach ($dir in $state.WatchDirs) {
        try {
            Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                $fp  = $_.FullName
                $lwt = $_.LastWriteTimeUtc
                if (-not $state.Snapshot.ContainsKey($fp)) {
                    if (-not $state.Seen.Contains($fp)) {
                        $ts = (Get-Date).ToString("HH:mm:ss")
                        $events.Add("[$ts] Created  $fp")
                    }
                    # Already in Seen — will be merged below
                } elseif ($state.Snapshot[$fp] -ne $lwt) {
                    $ts = (Get-Date).ToString("HH:mm:ss")
                    $events.Add("[$ts] Modified $fp")
                }
            }
        } catch {}
    }
    # Merge transient hits from polling
    foreach ($line in $state.Seen.Values) { $events.Add($line) }

    return @($events | Sort-Object -Unique)
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

    # Snapshot all watched directories BEFORE launching Bambu
    $diagState = $null
    if ($DiagnoseFiles) {
        Write-Host "`n  [DIAG] Snapshotting file system..." -ForegroundColor Yellow
        $diagState = Start-FileSnapshot $workDir (Split-Path $BambuPath -Parent)
        Write-Host "  [DIAG] Watching $($diagState.WatchDirs.Count) directories for new files." -ForegroundColor Yellow
    }

    # Added --uptodate and --allow-newer-file to bypass all version mismatch prompts.
    # WorkingDirectory is set explicitly so Bambu writes result.json into the design
    # folder rather than wherever the calling shell's CWD happens to be.
    $procArgs = "--debug 3 --no-check --uptodate --allow-newer-file --slice 1 --min-save --export-3mf `"$slicedOut`" `"$filePath`""
    $proc = Start-Process -FilePath $BambuPath -ArgumentList $procArgs `
        -WorkingDirectory $workDir `
        -RedirectStandardOutput $logOut -RedirectStandardError $logErr -WindowStyle Hidden -PassThru

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
        # Poll for new files every tick so transient files aren't missed
        if ($null -ne $diagState) { Update-FileSnapshot $diagState }
        Start-Sleep -Seconds 3
    }
    Write-Host " [DONE]" -ForegroundColor Green

    # Final snapshot diff + report
    if ($null -ne $diagState) {
        $diagLog = Stop-FileSnapshot $diagState
        $diagReportPath = Join-Path $workDir "${baseName}_DiagFiles.txt"
        Write-Host "`n  [DIAG] New/modified files during slice ($($diagLog.Count) total):" -ForegroundColor Yellow
        if ($diagLog.Count -gt 0) {
            $diagLog | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
        } else {
            Write-Host "    (none detected)" -ForegroundColor DarkGray
        }
        # Also check stdout for JSON in case Bambu pipes the result there
        if (Test-Path $logOut) {
            $stdoutContent = Get-Content $logOut -Raw -ErrorAction SilentlyContinue
            if ($stdoutContent -match '^\s*\{') {
                Write-Host "`n  [DIAG] Stdout contains JSON!" -ForegroundColor Cyan
                $diagLog += ""
                $diagLog += "=== STDOUT JSON ==="
                $diagLog += $stdoutContent
            }
        }
        $diagLog | Set-Content -Path $diagReportPath -Encoding UTF8
        Write-Host "  [DIAG] Report saved: $diagReportPath" -ForegroundColor Yellow
    }

    if (-not (Test-Path $slicedOut)) {
        Write-Host "`n  [!] ERROR: Bambu Studio failed to generate $slicedOut" -ForegroundColor Red
        if (Test-Path $logOut) { Get-Content $logOut | Select-Object -Last 10 | Write-Host -ForegroundColor DarkGray }
        if (Test-Path $logErr) { Get-Content $logErr | Select-Object -Last 10 | Write-Host -ForegroundColor DarkGray }
        if (Test-Path $logOut) { Remove-Item $logOut -Force -ErrorAction SilentlyContinue }
        if (Test-Path $logErr) { Remove-Item $logErr -Force -ErrorAction SilentlyContinue }
        return $false
    }

    # Bambu Studio writes result.json to its working directory after a successful slice.
    # Rename it to {baseName}_result.json so it stays associated with this specific file
    # and doesn't get overwritten when the Isolated file is sliced next.
    $genericResult = Join-Path $workDir "result.json"
    $namedResult   = Join-Path $workDir "${baseName}_result.json"
    if (Test-Path $genericResult) {
        Move-Item -Path $genericResult -Destination $namedResult -Force -ErrorAction SilentlyContinue
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