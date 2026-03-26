# ============================================================
# patch_3mf.ps1  -  Bambu .3mf settings patcher
#
# Only processes files ending in "Final.3mf" or "Full.3mf".
# Drag and drop a folder, single file, or multiple files
# onto patch_3mf.bat to run.
#
# Settings applied to Metadata/project_settings.config:
#
#   additional_cooling_fan_speed -> 40  (auxiliary part cooling fan)
#   fan_min_speed                -> 40  (minimum fan speed threshold)
#   fan_max_speed                -> 60  (max fan speed threshold)
#
#   Also registers all of the above in different_settings_to_system
#   so Bambu Studio recognises them as active overrides.
# ============================================================

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$CONFIG_ENTRY = "Metadata/project_settings.config"

# Only filenames ending with these suffixes are processed
$ALLOWED_SUFFIXES = @("Final.3mf", "Full.3mf")

# Keys that go in the filament 8-element array (Standard slots only = even indices)
# Value = what to set at even indices; odd indices stay "nil"
$FILAMENT_OVERRIDE_PATCHES = [ordered]@{}

# Keys that use a simple 4-element array (one per filament, all slots set)
$PLATE_TEMP_PATCHES = [ordered]@{
    "additional_cooling_fan_speed" = "40"
    "fan_min_speed"                = "40"
    "fan_max_speed"                = "60"
}

# All override keys that must be registered in different_settings_to_system
$OVERRIDE_KEYS_TO_REGISTER = @(
    "additional_cooling_fan_speed",
    "fan_min_speed",
    "fan_max_speed"
)

# ------------------------------------------------------------
# Helper: check filename against allowed suffixes
# ------------------------------------------------------------
function Is-Allowed {
    param([string]$FilePath)
    $name = Split-Path $FilePath -Leaf
    foreach ($suffix in $ALLOWED_SUFFIXES) {
        if ($name.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

# ------------------------------------------------------------
# PS5.1 compatible: convert PSCustomObject tree to hashtable
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
# Apply all patches; return list of change descriptions
# ------------------------------------------------------------
function Apply-Patches {
    param([hashtable]$Data)
    $changes = @()

    # --- 8-element filament override arrays (even = Standard, odd = High Flow) ---
    foreach ($key in $FILAMENT_OVERRIDE_PATCHES.Keys) {
        $newVal = $FILAMENT_OVERRIDE_PATCHES[$key]
        if (-not $Data.ContainsKey($key)) { continue }

        $arr = @($Data[$key])
        $changed = $false
        for ($i = 0; $i -lt $arr.Count; $i++) {
            if ($i % 2 -eq 0) {
                # Standard slot - set the value
                if ($arr[$i] -ne $newVal) {
                    $arr[$i] = $newVal
                    $changed = $true
                }
            } else {
                # High Flow slot - ensure it stays nil
                if ($arr[$i] -ne "nil") {
                    $arr[$i] = "nil"
                    $changed = $true
                }
            }
        }
        if ($changed) {
            $Data[$key] = $arr
            $changes += "  ${key}: even indices -> '$newVal', odd indices -> 'nil'"
        }
    }

    # --- 4-element plate temp arrays (all slots) ---
    foreach ($key in $PLATE_TEMP_PATCHES.Keys) {
        $newVal = $PLATE_TEMP_PATCHES[$key]
        if (-not $Data.ContainsKey($key)) { continue }

        $arr = @($Data[$key])
        $oldFirst = $arr[0]
        $newArr = @($newVal) * $arr.Count
        if (($arr | Where-Object { $_ -ne $newVal }).Count -gt 0) {
            $Data[$key] = $newArr
            $changes += "  ${key}: '$oldFirst' -> '$newVal'  (x$($arr.Count))"
        }
    }

    # --- Register overrides in different_settings_to_system ---
    if ($Data.ContainsKey("different_settings_to_system")) {
        $dsts = @($Data["different_settings_to_system"])
        # Element [0] is process overrides, last element is empty string
        # Elements [1] through [len-2] are per-filament override lists
        $dsts_changed = $false
        for ($fi = 1; $fi -lt ($dsts.Count - 1); $fi++) {
            $existing = $dsts[$fi]
            $keySet = if ($existing -ne "") { $existing -split ";" } else { @() }
            $keySetList = [System.Collections.Generic.List[string]]$keySet

            foreach ($ok in $OVERRIDE_KEYS_TO_REGISTER) {
                if (-not $keySetList.Contains($ok)) {
                    $keySetList.Add($ok)
                    $dsts_changed = $true
                }
            }

            # Re-sort alphabetically to match Bambu Studio's ordering
            $sorted = ($keySetList | Sort-Object) -join ";"
            if ($sorted -ne $existing) {
                $dsts[$fi] = $sorted
                $dsts_changed = $true
            }
        }
        if ($dsts_changed) {
            $Data["different_settings_to_system"] = $dsts
            $changes += "  different_settings_to_system: registered override keys for all filament slots"
        }
    }

    return $changes
}

# ------------------------------------------------------------
# Patch a single .3mf file in place
# ------------------------------------------------------------
function Patch-3mf {
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
            Write-Host "  ERROR: '$CONFIG_ENTRY' not found - is this a Bambu .3mf?"
            $srcZip.Dispose()
            $srcStream.Dispose()
            return $false
        }

        $reader   = [System.IO.StreamReader]::new($configEntry.Open())
        $jsonText = $reader.ReadToEnd()
        $reader.Dispose()

        $data    = ConvertTo-Hashtable (ConvertFrom-Json $jsonText)
        $changes = Apply-Patches -Data $data

        $anyChanges = $changes.Count -gt 0

        if ($anyChanges) {
            Write-Host "  Changes:"
            $changes | ForEach-Object { Write-Host $_ }
        }
        else {
            Write-Host "  All target settings already at desired values."
        }

        $dstStream = [System.IO.File]::Open($tmpPath, [System.IO.FileMode]::Create)
        $dstZip    = [System.IO.Compression.ZipArchive]::new($dstStream, [System.IO.Compression.ZipArchiveMode]::Create)

        foreach ($entry in $srcZip.Entries) {
            $srcEntryStream = $entry.Open()
            $dstEntry       = $dstZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $dstEntryStream = $dstEntry.Open()

            if ($entry.FullName -eq $CONFIG_ENTRY -and $anyChanges) {
                $newJson  = $data | ConvertTo-Json -Depth 20 -Compress:$false
                $newBytes = [System.Text.Encoding]::UTF8.GetBytes($newJson)
                $dstEntryStream.Write($newBytes, 0, $newBytes.Length)
            }
            else {
                $srcEntryStream.CopyTo($dstEntryStream)
            }

            $dstEntryStream.Dispose()
            $srcEntryStream.Dispose()
        }

        $dstZip.Dispose()
        $dstStream.Dispose()
        $srcZip.Dispose()
        $srcStream.Dispose()

        if ($anyChanges) {
            [System.IO.File]::Copy($tmpPath, $FilePath, $true)
            Write-Host "  Saved OK."
        }
    }
    catch {
        Write-Host "  ERROR: $_"
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force }
        return $false
    }
    finally {
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force }
    }

    return $anyChanges
}

# ------------------------------------------------------------
# Collect candidate files from drag-and-drop args
# ------------------------------------------------------------
$candidates = @()

if ($args.Count -eq 0) {
    $candidates = Get-ChildItem -Path (Get-Location) -Filter "*.3mf" -Recurse |
                  Select-Object -ExpandProperty FullName
}
else {
    foreach ($arg in $args) {
        if (Test-Path $arg -PathType Container) {
            $candidates += Get-ChildItem -Path $arg -Filter "*.3mf" -Recurse |
                           Select-Object -ExpandProperty FullName
        }
        elseif (Test-Path $arg -PathType Leaf) {
            $candidates += (Resolve-Path $arg).Path
        }
        else {
            Write-Host "Warning: skipping '$arg' (not found)"
        }
    }
}

$targets = @($candidates | Where-Object { Is-Allowed $_ })
$skipped = @($candidates | Where-Object { -not (Is-Allowed $_) })

if ($skipped.Count -gt 0) {
    Write-Host ""
    Write-Host "Skipped ($($skipped.Count) file(s) don't end in 'Final.3mf' or 'Full.3mf'):"
    $skipped | ForEach-Object { Write-Host "  $(Split-Path $_ -Leaf)" }
}

if ($targets.Count -eq 0) {
    Write-Host ""
    Write-Host "No matching files found. Only files ending in 'Final.3mf' or 'Full.3mf' are processed."
    exit 1
}

Write-Host ""
Write-Host "Found $($targets.Count) matching file(s) to process."
Write-Host ""
Write-Host "Target settings:"
Write-Host "  additional_cooling_fan_speed -> 40  (auxiliary part cooling fan)"
Write-Host "  fan_min_speed                -> 40  (minimum fan speed threshold)"
Write-Host "  fan_max_speed                -> 60  (max fan speed threshold)"

$changed = 0
foreach ($f in $targets) {
    $result = Patch-3mf -FilePath $f
    if ($result) { $changed++ }
}

Write-Host ""
Write-Host ("=" * 60)
Write-Host "Done. $changed/$($targets.Count) file(s) updated."