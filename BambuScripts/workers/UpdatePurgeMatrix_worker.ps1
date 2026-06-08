# ============================================================
# UpdatePurgeMatrix_worker.ps1  -  Bambu .3mf purge matrix updater
#
# For each .3mf dropped on the caller:
#   1. Reads filament_colour from Metadata/project_settings.config
#   2. Matches each hex color to a filament name via FilamentLibrary.csv
#   3. Looks up (source, target) pairs in PurgeDictionary.csv
#   4. Updates flush_volumes_matrix entries where Tuned_Volume is present
#      - entries with no tuned value are left unchanged.
#
# Drag and drop files or folders onto UpdatePurgeMatrix.bat to run.
# ============================================================

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$SCRIPT_DIR   = Split-Path $MyInvocation.MyCommand.Path -Parent
$LIB_DIR      = Join-Path $SCRIPT_DIR "..\libraries"
$FILAMENT_LIB = Join-Path $LIB_DIR "FilamentLibrary.csv"
$PURGE_DICT   = Join-Path $LIB_DIR "PurgeDictionary.csv"
$CONFIG_ENTRY = "Metadata/project_settings.config"

# ------------------------------------------------------------
# Load FilamentLibrary: hex -> name, plus list for fuzzy match
# ------------------------------------------------------------
$ColorToName   = @{}
$FilamentColors = @()

foreach ($row in (Import-Csv $FILAMENT_LIB -Header Name,R,G,B)) {
    if ($row.Name -and $row.Name -ne 'N/A') {
        try {
            $r = [int]$row.R; $g = [int]$row.G; $b = [int]$row.B
            $hex = ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b).ToUpper()
            if (-not $ColorToName.ContainsKey($hex)) {
                $ColorToName[$hex] = $row.Name
            }
            $FilamentColors += [PSCustomObject]@{ Name=$row.Name; R=$r; G=$g; B=$b }
        } catch {}
    }
}

function Resolve-FilamentName {
    param([string]$Hex)
    # Strip alpha channel if present (#RRGGBBAA -> #RRGGBB)
    $clean = $Hex.ToUpper() -replace '^(#[0-9A-F]{6})[0-9A-F]{0,2}$', '$1'
    if ($ColorToName.ContainsKey($clean)) { return $ColorToName[$clean] }

    # Fuzzy: nearest by Euclidean RGB distance
    $hexClean = $clean.TrimStart('#')
    if ($hexClean.Length -lt 6) { return $null }
    $r = [Convert]::ToInt32($hexClean.Substring(0,2), 16)
    $g = [Convert]::ToInt32($hexClean.Substring(2,2), 16)
    $b = [Convert]::ToInt32($hexClean.Substring(4,2), 16)

    $best = $null; $bestDist = [double]::MaxValue
    foreach ($fc in $FilamentColors) {
        $dist = [Math]::Sqrt(($r-$fc.R)*($r-$fc.R) + ($g-$fc.G)*($g-$fc.G) + ($b-$fc.B)*($b-$fc.B))
        if ($dist -lt $bestDist) { $bestDist = $dist; $best = $fc.Name }
    }
    return $best
}

# ------------------------------------------------------------
# Load PurgeDictionary: "Source|Target" -> Tuned_Volume (col E only)
# ------------------------------------------------------------
$PurgeDict = @{}

foreach ($row in (Import-Csv $PURGE_DICT -Delimiter "`t")) {
    $src = $row.Source_Filament.Trim()
    $tgt = $row.Target_Filament.Trim()
    $vol = $row.Tuned_Volume.Trim()
    if ($src -and $tgt -and $vol -ne '') {
        $PurgeDict["$src|$tgt"] = $vol
    }
}

Write-Host ""
Write-Host "Library loaded: $($FilamentColors.Count) filament colors, $($PurgeDict.Count) tuned purge entries."

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
# Update purge matrix in a single .3mf
# ------------------------------------------------------------
function Update-3mf {
    param([string]$FilePath)

    $fileName = Split-Path $FilePath -Leaf
    Write-Host ""
    Write-Host ("-" * 60)
    Write-Host "Processing: $fileName"

    $tmpPath = [System.IO.Path]::GetTempFileName() + ".3mf"

    try {
        $srcStream = [System.IO.File]::OpenRead($FilePath)
        $srcZip    = [System.IO.Compression.ZipArchive]::new($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)

        $configEntry = $srcZip.GetEntry($CONFIG_ENTRY)
        if ($null -eq $configEntry) {
            Write-Host "  SKIP: '$CONFIG_ENTRY' not found - not a Bambu .3mf?"
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        $reader   = [System.IO.StreamReader]::new($configEntry.Open())
        $jsonText = $reader.ReadToEnd()
        $reader.Dispose()

        $data = ConvertTo-Hashtable (ConvertFrom-Json $jsonText)

        if (-not $data.ContainsKey('filament_colour') -or -not $data.ContainsKey('flush_volumes_matrix')) {
            Write-Host "  SKIP: missing filament_colour or flush_volumes_matrix keys."
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        $colors = @($data['filament_colour'])
        $matrix = @($data['flush_volumes_matrix'])
        $N      = $colors.Count

        if ($matrix.Count -ne ($N * $N)) {
            Write-Host "  SKIP: matrix size $($matrix.Count) doesn't match $N x $N = $($N*$N)."
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        # Resolve each slot's color to a filament name
        $names = @()
        for ($i = 0; $i -lt $N; $i++) {
            $hex  = $colors[$i]
            $name = Resolve-FilamentName $hex
            $names += $name
            $nameLabel = if ($name) { $name } else { '(no match)' }
            Write-Host ("  Slot {0}: {1} -> {2}" -f ($i+1), $hex, $nameLabel)
        }

        # Apply tuned values from PurgeDictionary
        $changes = 0
        for ($i = 0; $i -lt $N; $i++) {
            for ($j = 0; $j -lt $N; $j++) {
                if ($i -eq $j) { continue }
                $srcName = $names[$i]; $tgtName = $names[$j]
                if (-not $srcName -or -not $tgtName) { continue }

                $key = "$srcName|$tgtName"
                if ($PurgeDict.ContainsKey($key)) {
                    $idx    = $i * $N + $j
                    $oldVal = $matrix[$idx]
                    $newVal = $PurgeDict[$key]
                    if ($oldVal -ne $newVal) {
                        $matrix[$idx] = $newVal
                        $changes++
                        Write-Host ("    [{0}->{1}] {2} -> {3}: {4} -> {5}" -f ($i+1), ($j+1), $srcName, $tgtName, $oldVal, $newVal)
                    }
                }
            }
        }

        if ($changes -eq 0) {
            Write-Host "  No matrix changes needed."
            $srcZip.Dispose(); $srcStream.Dispose()
            return $false
        }

        $data['flush_volumes_matrix'] = $matrix

        # Write updated zip to temp file then replace original
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

        $word = if ($changes -eq 1) { 'entry' } else { 'entries' }
        Write-Host "  Saved. $changes matrix $word updated."
        return $true
    }
    catch {
        Write-Host "  ERROR: $_"
        return $false
    }
    finally {
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force }
    }
}

# ------------------------------------------------------------
# Collect .3mf files from drag-and-drop args
# ------------------------------------------------------------
$candidates = @()

if ($args.Count -eq 0) {
    $candidates = Get-ChildItem -Path (Get-Location) -Filter "*.3mf" -Recurse |
                  Select-Object -ExpandProperty FullName
} else {
    foreach ($arg in $args) {
        if (Test-Path $arg -PathType Container) {
            $candidates += Get-ChildItem -Path $arg -Filter "*.3mf" -Recurse |
                           Select-Object -ExpandProperty FullName
        } elseif (Test-Path $arg -PathType Leaf) {
            $candidates += (Resolve-Path $arg).Path
        } else {
            Write-Host "Warning: skipping '$arg' (not found)"
        }
    }
}

$targets = @($candidates | Where-Object { $_ -like "*.3mf" })

if ($targets.Count -eq 0) {
    Write-Host ""
    Write-Host "No .3mf files found. Drop a .3mf file or a folder containing .3mf files."
    exit 1
}

Write-Host "Found $($targets.Count) .3mf file(s) to process."

$changed = 0
foreach ($f in $targets) {
    if (Update-3mf -FilePath $f) { $changed++ }
}

Write-Host ""
Write-Host ("=" * 60)
Write-Host "Done. $changed / $($targets.Count) file(s) updated."
