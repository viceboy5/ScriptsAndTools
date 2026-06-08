# ============================================================
# ResetPurgeMatrix_worker.ps1  -  Reset purge (flush) volumes to Bambu defaults
#
# For each .3mf found (recursively, under each path passed in):
#   1. Removes the custom 'flush_volumes_matrix' entry from
#      Metadata/project_settings.config (in-place zip rewrite)
#   2. Re-exports/re-saves the file through the Bambu Studio CLI
#      (--export-3mf, no slicing) so Bambu recomputes the matrix
#      from its own built-in defaults and writes it back.
#
# .gcode.3mf (sliced output) files are skipped - they're
# regenerated fresh on the next slice anyway.
#
# Usage: ResetPurgeMatrix_worker.ps1 <folder-or-file> [...]
# ============================================================
param(
    [string[]]$Paths = @(),
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe"
)
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$CONFIG_ENTRY = "Metadata/project_settings.config"

# ------------------------------------------------------------
# PS5.1 compatible: PSCustomObject tree -> hashtable
# ------------------------------------------------------------
function ConvertTo-Hashtable {
    param([Parameter(ValueFromPipeline)]$InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @($InputObject | ForEach-Object { ConvertTo-Hashtable $_ })
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $ht = @{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $ht[$prop.Name] = ConvertTo-Hashtable $prop.Value
        }
        return $ht
    }
    return $InputObject
}

# ------------------------------------------------------------
# Step 1: strip flush_volumes_matrix out of project_settings.config
# ------------------------------------------------------------
function Remove-FlushMatrix([string]$FilePath) {
    $tmpPath = [System.IO.Path]::GetTempFileName() + ".3mf"
    try {
        $srcStream = [System.IO.File]::OpenRead($FilePath)
        $srcZip    = [System.IO.Compression.ZipArchive]::new($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)

        $configEntry = $srcZip.GetEntry($CONFIG_ENTRY)
        if ($null -eq $configEntry) {
            Write-Host "  SKIP: '$CONFIG_ENTRY' not found - not a Bambu project .3mf?"
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        $reader   = [System.IO.StreamReader]::new($configEntry.Open())
        $jsonText = $reader.ReadToEnd()
        $reader.Dispose()

        $data = ConvertTo-Hashtable (ConvertFrom-Json $jsonText)

        if (-not $data.ContainsKey('flush_volumes_matrix')) {
            Write-Host "  No flush_volumes_matrix present - nothing to remove."
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        $data.Remove('flush_volumes_matrix') | Out-Null

        $dstStream = [System.IO.File]::Open($tmpPath, [System.IO.FileMode]::Create)
        $dstZip    = [System.IO.Compression.ZipArchive]::new($dstStream, [System.IO.Compression.ZipArchiveMode]::Create)

        foreach ($entry in $srcZip.Entries) {
            $srcEntryStream = $entry.Open()
            $dstEntry       = $dstZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $dstEntryStream = $dstEntry.Open()

            if ($entry.FullName -eq $CONFIG_ENTRY) {
                $newJson  = $data | ConvertTo-Json -Depth 20 -Compress:$false
                $newBytes = [System.Text.Encoding]::UTF8.GetBytes($newJson)
                $dstEntryStream.Write($newBytes, 0, $newBytes.Length)
            } else {
                $srcEntryStream.CopyTo($dstEntryStream)
            }

            $dstEntryStream.Dispose()
            $srcEntryStream.Dispose()
        }

        $dstZip.Dispose(); $dstStream.Dispose()
        $srcZip.Dispose(); $srcStream.Dispose()

        [System.IO.File]::Copy($tmpPath, $FilePath, $true)
        Write-Host "  Removed flush_volumes_matrix from project_settings.config."
        return $true
    }
    catch {
        Write-Host "  ERROR removing matrix: $_" -ForegroundColor Red
        return $false
    }
    finally {
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force -ErrorAction SilentlyContinue }
    }
}

# ------------------------------------------------------------
# Step 2: re-export/re-save through Bambu Studio so it
# recomputes the matrix from its own defaults
# ------------------------------------------------------------
function Resave-WithBambu([string]$FilePath) {
    if (-not (Test-Path $BambuPath)) {
        Write-Host "  SKIP re-export: Bambu Studio not found at $BambuPath" -ForegroundColor Yellow
        return $false
    }
    $tag     = [guid]::NewGuid().ToString('N').Substring(0,8)
    $tempOut = $FilePath + ".resave.tmp.3mf"
    $logOut  = Join-Path $env:TEMP "reset_purge_resave_out_$tag.txt"
    $logErr  = Join-Path $env:TEMP "reset_purge_resave_err_$tag.txt"
    $procArgs = "--debug 3 --no-check --uptodate --allow-newer-file --export-3mf `"$tempOut`" `"$FilePath`""

    $proc = Start-Process -FilePath $BambuPath -ArgumentList $procArgs `
                          -WorkingDirectory $env:TEMP `
                          -RedirectStandardOutput $logOut -RedirectStandardError $logErr `
                          -WindowStyle Hidden -PassThru
    $proc.WaitForExit()
    foreach ($log in @($logOut, $logErr)) { if (Test-Path $log) { Remove-Item $log -Force -ErrorAction SilentlyContinue } }

    if (Test-Path $tempOut) {
        Move-Item $tempOut $FilePath -Force
        Write-Host "  Re-exported via Bambu Studio - matrix regenerated with defaults."
        return $true
    } else {
        Write-Host "  ERROR: Bambu re-export produced no output (file left with matrix removed)." -ForegroundColor Red
        return $false
    }
}

# ------------------------------------------------------------
# Collect .3mf targets (skip sliced .gcode.3mf outputs)
# ------------------------------------------------------------
$candidates = @()
foreach ($p in $Paths) {
    if (Test-Path $p -PathType Container) {
        $candidates += Get-ChildItem -Path $p -Filter "*.3mf" -Recurse | Select-Object -ExpandProperty FullName
    } elseif (Test-Path $p -PathType Leaf) {
        $candidates += (Resolve-Path $p).Path
    } else {
        Write-Host "Warning: skipping '$p' (not found)"
    }
}
$targets = @($candidates | Where-Object { $_ -like "*.3mf" -and $_ -notlike "*.gcode.3mf" } | Sort-Object -Unique)

if ($targets.Count -eq 0) {
    Write-Host "No project .3mf files found (gcode.3mf outputs are skipped)."
    exit 1
}

Write-Host "Found $($targets.Count) project .3mf file(s) to reset."
$ok = 0
foreach ($f in $targets) {
    Write-Host ""
    Write-Host ("-" * 60)
    Write-Host "Processing: $(Split-Path $f -Leaf)"
    if (Remove-FlushMatrix -FilePath $f) {
        if (Resave-WithBambu -FilePath $f) { $ok++ }
    } else {
        Write-Host "  Skipped re-export (matrix already absent or not a project file)."
    }
}

Write-Host ""
Write-Host ("=" * 60)
Write-Host "Done. $ok / $($targets.Count) file(s) reset and re-exported."
