# ============================================================
# ResetMetadata_worker.ps1  -  Strip and reset .3mf metadata
#
# For each project .3mf found (recursive, skips .gcode.3mf):
#   1. Removes flush_volumes_matrix from Metadata/project_settings.config
#   2. Deletes all embedded .png thumbnails (plate_*.png, pick_*.png,
#      top_*.png, etc.) - Bambu Studio regenerates these on next open/slice
#   [add more resets here as needed]
#   3. Re-exports/re-saves through the Bambu Studio CLI so Bambu
#      recomputes defaults and writes a clean file
#
# Usage: ResetMetadata_worker.ps1 <folder-or-file> [...]
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
# Process a single .3mf: strip metadata, rewrite zip in place
# Returns $true if any changes were made (and Bambu resave needed)
# ------------------------------------------------------------
function Reset-Metadata([string]$FilePath) {
    $tmpPath = [System.IO.Path]::GetTempFileName() + ".3mf"
    $changes = [System.Collections.Generic.List[string]]::new()

    try {
        $srcStream = [System.IO.File]::OpenRead($FilePath)
        $srcZip    = [System.IO.Compression.ZipArchive]::new($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)

        $configEntry = $srcZip.GetEntry($CONFIG_ENTRY)
        if ($null -eq $configEntry) {
            Write-Host "  SKIP: '$CONFIG_ENTRY' not found - not a Bambu project .3mf?"
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        # Read and parse project_settings.config
        $reader   = [System.IO.StreamReader]::new($configEntry.Open())
        $jsonText = $reader.ReadToEnd()
        $reader.Dispose()
        $data = ConvertTo-Hashtable (ConvertFrom-Json $jsonText)

        # --- Reset 1: flush_volumes_matrix ---
        if ($data.ContainsKey('flush_volumes_matrix')) {
            $data.Remove('flush_volumes_matrix') | Out-Null
            $changes.Add('flush_volumes_matrix removed')
        }

        # --- Add more resets here as needed ---
        # e.g. $data.Remove('some_other_key') | Out-Null

        # List all PNG entries to strip
        $pngEntries = @($srcZip.Entries | Where-Object { $_.FullName -match '\.png$' })
        $pngNames   = $pngEntries | ForEach-Object { $_.FullName }
        if ($pngNames.Count -gt 0) {
            $changes.Add("$($pngNames.Count) PNG thumbnail(s) removed")
        }

        if ($changes.Count -eq 0) {
            Write-Host "  Nothing to reset - file already clean."
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        # Rewrite zip, skipping PNG entries and updating config
        $dstStream = [System.IO.File]::Open($tmpPath, [System.IO.FileMode]::Create)
        $dstZip    = [System.IO.Compression.ZipArchive]::new($dstStream, [System.IO.Compression.ZipArchiveMode]::Create)

        foreach ($entry in $srcZip.Entries) {
            # Skip PNG thumbnails
            if ($entry.FullName -match '\.png$') { continue }

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
        foreach ($c in $changes) { Write-Host "  $c" }
        return $true
    }
    catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        return $false
    }
    finally {
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force -ErrorAction SilentlyContinue }
    }
}

# ------------------------------------------------------------
# Re-export/re-save a single file through the Bambu Studio CLI
# ------------------------------------------------------------
function Resave-WithBambu([string]$FilePath) {
    if (-not (Test-Path $BambuPath)) {
        Write-Host "  SKIP re-export: Bambu Studio not found at $BambuPath" -ForegroundColor Yellow
        return $false
    }
    $tag      = [guid]::NewGuid().ToString('N').Substring(0,8)
    $tempOut  = $FilePath + ".resave.tmp.3mf"
    $logOut   = Join-Path $env:TEMP "reset_meta_resave_out_$tag.txt"
    $logErr   = Join-Path $env:TEMP "reset_meta_resave_err_$tag.txt"
    $procArgs = "--debug 3 --no-check --uptodate --allow-newer-file --export-3mf `"$tempOut`" `"$FilePath`""

    $proc = Start-Process -FilePath $BambuPath -ArgumentList $procArgs `
                          -WorkingDirectory $env:TEMP `
                          -RedirectStandardOutput $logOut -RedirectStandardError $logErr `
                          -WindowStyle Hidden -PassThru
    $proc.WaitForExit()
    foreach ($log in @($logOut, $logErr)) { if (Test-Path $log) { Remove-Item $log -Force -ErrorAction SilentlyContinue } }

    if (Test-Path $tempOut) {
        Move-Item $tempOut $FilePath -Force
        Write-Host "  Re-exported via Bambu Studio."
        return $true
    } else {
        Write-Host "  ERROR: Bambu re-export produced no output." -ForegroundColor Red
        return $false
    }
}

# ------------------------------------------------------------
# Collect .3mf targets
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
    if (Reset-Metadata -FilePath $f) {
        if (Resave-WithBambu -FilePath $f) { $ok++ }
    } else {
        Write-Host "  Skipped re-export (nothing changed)."
    }
}

Write-Host ""
Write-Host ("=" * 60)
Write-Host "Done. $ok / $($targets.Count) file(s) reset and re-exported."
