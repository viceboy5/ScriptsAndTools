param(
    [string]$InputPath,                                                                    # 3mf to read from
    [string]$OutputPath,                                                                   # BOD.3mf destination
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe",
    [int]$PairCount    = 5,                                                                # How many pairs to keep (Full mode)
    [string]$Mode      = "Full"                                                            # "Full" or "Grid"
)
$ErrorActionPreference = 'Stop'

# ================================================================================
#  create_bod_worker.ps1
#  Takes a merged Full.3mf, keeps the $PairCount pairs closest to the plate
#  centre (128, 128), removes lone objects and all other pairs, and repacks
#  to $OutputPath ready for slicing.
#
#  "Pairs" are identified by the name tag set by merge_3mf_worker.ps1:
#    MergedGroup_N  ->  pair / group (keep candidates)
#    *_Lone_*       ->  lone / single (always removed)
#    text / version ->  ignored label (always removed)
#  Objects with no recognisable name tag are treated as pairs (safe fallback).
# ================================================================================

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

function Parse-Tx([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return [double[]](1,0,0, 0,1,0, 0,0,1, 0,0,0) }
    return [double[]]($s.Trim() -split '\s+')
}
function Save-Xml([xml]$doc, [string]$path) {
    $ws = New-Object System.Xml.XmlWriterSettings
    $ws.Encoding = New-Object System.Text.UTF8Encoding($false)
    $ws.Indent   = $true
    $w = [System.Xml.XmlWriter]::Create($path, $ws)
    $doc.Save($w); $w.Close()
}
function Find-File([string]$base, [string]$rel) {
    $norm = $rel -replace '\\', '/'
    $p = Join-Path $base ($norm -replace '/', [System.IO.Path]::DirectorySeparatorChar)
    if (Test-Path $p) { return $p }
    $p2 = Join-Path $base $rel
    if (Test-Path $p2) { return $p2 }
    return $null
}

if (-not (Test-Path $InputPath)) {
    Write-Host "[BOD] ERROR: InputPath not found: $InputPath" -ForegroundColor Red
    exit 1
}

# ================================================================================
#  GRID MODE  -  Final.3mf -> 15-copy grid BOD
# ================================================================================
if ($Mode -eq 'Grid') {

    $PLATE_W   = 256.0
    $PLATE_H   = 256.0
    $MARGIN    = 15.0
    $COPIES    = 15
    $MIN_GAP   = 3.0

    Write-Host "[BOD-Grid] Extracting Final.3mf..."
    $tempDir = Join-Path $env:TEMP "bod_grid_$([guid]::NewGuid().ToString().Substring(0,8))"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    try { [System.IO.Compression.ZipFile]::ExtractToDirectory($InputPath, $tempDir) }
    catch { Write-Host "[BOD-Grid] ERROR extracting: $_" -ForegroundColor Red; Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue; exit 1 }

    # -- Read plate_1.json for object bounding box --------------------------------
    $plate1Path = Find-File $tempDir 'Metadata/plate_1.json'
    if ($null -eq $plate1Path) { Write-Host "[BOD-Grid] ERROR: plate_1.json not found." -ForegroundColor Red; Remove-Item $tempDir -Recurse -Force; exit 1 }
    $plate1 = ConvertFrom-Json ([System.IO.File]::ReadAllText($plate1Path, [System.Text.Encoding]::UTF8))

    $designEntry = $plate1.bbox_objects | Where-Object { $_.name -ne 'wipe_tower' } | Select-Object -First 1
    if ($null -eq $designEntry) { Write-Host "[BOD-Grid] ERROR: no design object found in plate_1.json." -ForegroundColor Red; Remove-Item $tempDir -Recurse -Force; exit 1 }

    $bboxMinX = [double]$designEntry.bbox[0]; $bboxMinY = [double]$designEntry.bbox[1]
    $bboxMaxX = [double]$designEntry.bbox[2]; $bboxMaxY = [double]$designEntry.bbox[3]
    $bboxW    = $bboxMaxX - $bboxMinX
    $bboxD    = $bboxMaxY - $bboxMinY
    $origCX   = ($bboxMinX + $bboxMaxX) / 2.0
    $origCY   = ($bboxMinY + $bboxMaxY) / 2.0

    Write-Host ("[BOD-Grid] Design footprint: {0:F2} x {1:F2} mm  (centre {2:F2}, {3:F2})" -f $bboxW, $bboxD, $origCX, $origCY)

    # -- Choose best grid layout (3, 4, or 5 columns for 15 items) ---------------
    $avail    = $PLATE_W - 2.0 * $MARGIN
    $best     = $null; $bestMinGap = -1

    foreach ($cols in @(3, 4, 5)) {
        $rows = [int][Math]::Ceiling($COPIES / $cols)
        if ($cols -lt 2 -or $rows -lt 2) { continue }
        $gX = ($avail - $cols * $bboxW) / ($cols - 1)
        $gY = ($avail - $rows * $bboxD) / ($rows - 1)
        $mg = [Math]::Min($gX, $gY)
        if ($gX -ge $MIN_GAP -and $gY -ge $MIN_GAP -and $mg -gt $bestMinGap) {
            $bestMinGap = $mg; $best = @{ Cols=$cols; Rows=$rows; GapX=$gX; GapY=$gY }
        }
    }
    if ($null -eq $best) {
        Write-Host "[BOD-Grid] ERROR: design too large to fit $COPIES copies on a ${PLATE_W}x${PLATE_H} plate with ${MARGIN}mm margins." -ForegroundColor Red
        Remove-Item $tempDir -Recurse -Force; exit 1
    }
    Write-Host ("[BOD-Grid] Layout: {0} cols x {1} rows  gap {2:F1}mm x {3:F1}mm" -f $best.Cols, $best.Rows, $best.GapX, $best.GapY)

    # -- Parse 3dmodel.model -------------------------------------------------------
    $modelFile = (Get-ChildItem -Path $tempDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
    [xml]$xml  = [System.IO.File]::ReadAllText($modelFile, [System.Text.Encoding]::UTF8)
    $xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

    $buildNode     = $xml.SelectSingleNode('//m:build', $xns)
    $origBuildItems = @($buildNode.SelectNodes('m:item', $xns))
    Write-Host "[BOD-Grid] Original build items: $($origBuildItems.Count)"

    # Store original transforms indexed by position
    $origTxList = $origBuildItems | ForEach-Object { Parse-Tx ($_.GetAttribute('transform')) }

    # Remove all existing build items (we'll re-add them as copies)
    foreach ($item in $origBuildItems) { $buildNode.RemoveChild($item) | Out-Null }

    # -- Build 15 copies ----------------------------------------------------------
    $idx = 0
    for ($row = 0; $row -lt $best.Rows; $row++) {
        for ($col = 0; $col -lt $best.Cols; $col++) {
            if ($idx -ge $COPIES) { break }

            $newCX  = $MARGIN + $col * ($bboxW + $best.GapX) + $bboxW / 2.0
            $newCY  = $MARGIN + $row * ($bboxD + $best.GapY) + $bboxD / 2.0
            $deltaX = $newCX - $origCX
            $deltaY = $newCY - $origCY

            for ($i = 0; $i -lt $origBuildItems.Count; $i++) {
                $orig    = $origBuildItems[$i]
                $tx      = [double[]]$origTxList[$i].Clone()
                $tx[9]  += $deltaX
                $tx[10] += $deltaY

                $newItem = $xml.CreateElement('item', $nsCore)
                $newItem.SetAttribute('objectid', $orig.GetAttribute('objectid'))

                # Preserve p:UUID-style attributes with a fresh UUID per copy
                $uuidAttr = $orig.GetAttributeNode('UUID', $nsProd)
                if ($null -ne $uuidAttr) {
                    $newItem.SetAttribute('UUID', $nsProd, [guid]::NewGuid().ToString('D'))
                }

                $txStr = ($tx | ForEach-Object { $_.ToString('G9') }) -join ' '
                $newItem.SetAttribute('transform', $txStr)
                $newItem.SetAttribute('printable', '1')
                $buildNode.AppendChild($newItem) | Out-Null
            }
            $idx++
        }
    }

    # -- Update model_settings.config ---------------------------------------------
    $settPath = Find-File $tempDir 'Metadata/model_settings.config'
    if ($null -ne $settPath) {
        [xml]$sett  = [System.IO.File]::ReadAllText($settPath, [System.Text.Encoding]::UTF8)
        $plateNode  = $sett.SelectSingleNode('//plate')

        if ($null -ne $plateNode) {
            # Find primary objectid from the existing model_instance
            $existingMI = $plateNode.SelectSingleNode('model_instance')
            $primaryId  = if ($null -ne $existingMI) {
                $existingMI.SelectSingleNode('metadata[@key="object_id"]').GetAttribute('value')
            } else { $origBuildItems[0].GetAttribute('objectid') }

            # Remove all existing model_instances
            foreach ($mi in @($plateNode.SelectNodes('model_instance'))) { $plateNode.RemoveChild($mi) | Out-Null }

            # Add one model_instance per copy
            for ($c = 0; $c -lt $COPIES; $c++) {
                $mi = $sett.CreateElement('model_instance')
                $addMeta = {
                    param($key, $val)
                    $m = $sett.CreateElement('metadata')
                    $m.SetAttribute('key', $key); $m.SetAttribute('value', $val)
                    $mi.AppendChild($m) | Out-Null
                }
                & $addMeta 'object_id'   $primaryId
                & $addMeta 'instance_id' "$c"
                & $addMeta 'identify_id' "$([int64]50000 + $c)"
                $plateNode.AppendChild($mi) | Out-Null
            }
        }
        Save-Xml $sett $settPath
    }

    # -- Save model, strip plate images, repack -----------------------------------
    Save-Xml $xml $modelFile

    foreach ($f in (Get-ChildItem -Path (Join-Path $tempDir 'Metadata') -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '^(plate|pick|top)_.*\.(png|json)$' })) {
        Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
    try {
        Get-ChildItem $tempDir -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($tempDir.Length).TrimStart('\','/').Replace('\','/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
        }
    } finally { $zip.Dispose(); [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers() }
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

    # -- Optional Bambu resave ----------------------------------------------------
    if (Test-Path $BambuPath) {
        Write-Host "[BOD-Grid] Running Bambu resave..." -NoNewline
        $tempResave = $OutputPath + ".resave.tmp.3mf"
        $logOut = Join-Path $env:TEMP "bambu_bod_out.txt"; $logErr = Join-Path $env:TEMP "bambu_bod_err.txt"
        $proc = Start-Process -FilePath $BambuPath -ArgumentList "--debug 3 --no-check --uptodate --allow-newer-file --export-3mf `"$tempResave`" `"$OutputPath`"" `
                              -RedirectStandardOutput $logOut -RedirectStandardError $logErr -WindowStyle Hidden -PassThru
        $proc.WaitForExit()
        foreach ($log in @($logOut,$logErr)) { if (Test-Path $log) { Remove-Item $log -Force -ErrorAction SilentlyContinue } }
        if (Test-Path $tempResave) { Move-Item $tempResave $OutputPath -Force; Write-Host " done." } else { Write-Host " skipped." }
    }

    Write-Host "[BOD-Grid] BOD.3mf created: $OutputPath"
    exit 0
}

# -- Extract Full.3mf to a temp directory ------------------------------------
$tempDir = Join-Path $env:TEMP "bod_work_$([guid]::NewGuid().ToString().Substring(0,8))"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($InputPath, $tempDir)
} catch {
    Write-Host "[BOD] ERROR extracting Full.3mf: $_" -ForegroundColor Red
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

# -- Locate files in the extracted tree -------------------------------------
$modelFile   = (Get-ChildItem -Path $tempDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
$objectsDir  = Join-Path $tempDir '3D/Objects'
$relsPath    = Find-File $tempDir '3D/_rels/3dmodel.model.rels'
$settingsPath = Find-File $tempDir 'Metadata/model_settings.config'
$cutInfoPath  = Find-File $tempDir 'Metadata/cut_information.xml'

if ([string]::IsNullOrWhiteSpace($modelFile) -or -not (Test-Path $modelFile)) {
    Write-Host "[BOD] ERROR: 3dmodel.model not found in extracted archive." -ForegroundColor Red
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

[xml]$xml = [System.IO.File]::ReadAllText($modelFile, [System.Text.Encoding]::UTF8)
$xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

$buildItems = @($xml.SelectNodes('//m:build/m:item', $xns))
$objById    = @{}
foreach ($o in $xml.SelectNodes('//m:resources/m:object', $xns)) {
    $objById[$o.GetAttribute('id')] = $o
}

$hasSettings = ($null -ne $settingsPath) -and (Test-Path $settingsPath)
[xml]$settings = $null; $settObjById = @{}
if ($hasSettings) {
    $settings = [xml][System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)
    foreach ($node in $settings.config.ChildNodes) {
        if ($node.LocalName -eq 'object') { $settObjById[$node.GetAttribute('id')] = $node }
    }
}

# -- Classify each build item -----------------------------------------------
# Returns 'pair', 'lone', or 'ignored'
function Get-ItemClass([string]$objectId) {
    if (-not $hasSettings) { return 'pair' }
    $sObj = $settObjById[$objectId]
    if ($null -eq $sObj) { return 'pair' }

    $nameNode = $sObj.SelectSingleNode('metadata[@key="name"]')
    if ($null -eq $nameNode) { return 'pair' }
    $nameVal = $nameNode.GetAttribute('value')

    # Bambu-injected label (text / version) that is NOT a real mesh file
    if ($nameVal -match '(?i)text|version' -and $nameVal -notmatch '(?i)\.(stl|3mf|obj|step|stp)$') {
        return 'ignored'
    }
    # Lone item tagged by merge worker
    if ($nameVal -match '_Lone_\d+$') { return 'lone' }
    # Merged group tagged by merge worker
    if ($nameVal -match '^MergedGroup_') { return 'pair' }
    # Unknown name -- treat as pair (safe)
    return 'pair'
}

$pairItems   = [System.Collections.Generic.List[object]]::new()
$removeItems = [System.Collections.Generic.List[object]]::new()

foreach ($item in $buildItems) {
    $cls = Get-ItemClass ($item.GetAttribute('objectid'))
    if ($cls -eq 'pair') { $pairItems.Add($item) } else { $removeItems.Add($item) }
}

Write-Host "[BOD] Pairs found: $($pairItems.Count) | Removing (lones/ignored): $($removeItems.Count)"

if ($pairItems.Count -eq 0) {
    Write-Host "[BOD] WARNING: No pairs found in Full.3mf. Nothing to do." -ForegroundColor Yellow
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

# -- Pick the $PairCount pairs closest to plate centre (128, 128) -----------
$centreX = 128.0; $centreY = 128.0

$ranked = $pairItems | ForEach-Object {
    $tx = Parse-Tx ($_.GetAttribute('transform'))
    $dx = $tx[9] - $centreX; $dy = $tx[10] - $centreY
    [PSCustomObject]@{ Item = $_; Dist2 = $dx*$dx + $dy*$dy }
} | Sort-Object Dist2

$keepCount  = [Math]::Min($PairCount, $ranked.Count)
$keepItems  = [System.Collections.Generic.HashSet[object]]::new()
for ($i = 0; $i -lt $keepCount; $i++) { $keepItems.Add($ranked[$i].Item) | Out-Null }

# Everything not in the keep set gets removed
foreach ($item in $pairItems) {
    if (-not $keepItems.Contains($item)) { $removeItems.Add($item) }
}

Write-Host "[BOD] Keeping $keepCount pair(s), removing $($removeItems.Count) total item(s)."

# -- Remove unwanted build items and their object definitions ---------------
$killedIds = New-Object System.Collections.Generic.HashSet[string]

foreach ($item in $removeItems) {
    $id = $item.GetAttribute('objectid')
    $killedIds.Add($id) | Out-Null
    $item.ParentNode.RemoveChild($item) | Out-Null
    $obj = $objById[$id]
    if ($null -ne $obj) { $obj.ParentNode.RemoveChild($obj) | Out-Null }
}

# Also cascade-remove any component sub-objects that belong only to killed items
# (re-use the protectedIds pattern from merge_3mf_worker)
$remainingBuildItems = @($xml.SelectNodes('//m:build/m:item', $xns))
$protectedIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($item in $remainingBuildItems) { $protectedIds.Add($item.GetAttribute('objectid')) | Out-Null }

$added = $true
while ($added) {
    $added = $false
    foreach ($id in @($protectedIds)) {
        $obj = $objById[$id]
        if ($null -ne $obj) {
            foreach ($comp in $obj.SelectNodes('.//m:component', $xns)) {
                if ($protectedIds.Add($comp.GetAttribute('objectid'))) { $added = $true }
            }
        }
    }
}

foreach ($objId in @($objById.Keys)) {
    if (-not $protectedIds.Contains($objId)) {
        $killedIds.Add($objId) | Out-Null
        $obj = $objById[$objId]
        if ($null -ne $obj -and $null -ne $obj.ParentNode) {
            $obj.ParentNode.RemoveChild($obj) | Out-Null
        }
    }
}

# -- Prune model_settings.config --------------------------------------------
if ($hasSettings) {
    # Remove plate model_instance entries for killed ids
    if ($plate = $settings.SelectSingleNode('//plate')) {
        foreach ($inst in @($plate.SelectNodes('model_instance'))) {
            $metaId = $inst.SelectSingleNode('metadata[@key="object_id"]')
            if ($null -ne $metaId -and $killedIds.Contains($metaId.GetAttribute('value'))) {
                $inst.ParentNode.RemoveChild($inst) | Out-Null
            }
        }
    }

    # Remove orphaned <object> metadata blocks
    foreach ($kId in $killedIds) {
        $sObj = $settObjById[$kId]
        if ($null -ne $sObj -and $null -ne $sObj.ParentNode) {
            $sObj.ParentNode.RemoveChild($sObj) | Out-Null
        }
        # Also remove assemble_item entries
        foreach ($node in @($settings.SelectNodes("//*[@object_id='$kId']"))) {
            $node.ParentNode.RemoveChild($node) | Out-Null
        }
    }
}

# -- Prune cut_information.xml (if present) ---------------------------------
if (($null -ne $cutInfoPath) -and (Test-Path $cutInfoPath)) {
    [xml]$cutXml = [System.IO.File]::ReadAllText($cutInfoPath, [System.Text.Encoding]::UTF8)
    $cutModified = $false
    foreach ($cutObj in @($cutXml.SelectNodes('//*[local-name()="object"]'))) {
        if ($killedIds.Contains($cutObj.GetAttribute('id'))) {
            $cutObj.ParentNode.RemoveChild($cutObj) | Out-Null
            $cutModified = $true
        }
    }
    if ($cutModified) {
        $ws = New-Object System.Xml.XmlWriterSettings
        $ws.Encoding = New-Object System.Text.UTF8Encoding($false); $ws.Indent = $true
        $w = [System.Xml.XmlWriter]::Create($cutInfoPath, $ws); $cutXml.Save($w); $w.Close()
    }
}

# -- Clean unused .model geometry files and rebuild .rels -------------------
$preservedModelPaths = New-Object System.Collections.Generic.HashSet[string]
foreach ($id in $protectedIds) {
    $obj = $objById[$id]
    if ($null -eq $obj) { continue }
    foreach ($c in $obj.SelectNodes('m:components/m:component', $xns)) {
        $path = $c.GetAttribute('path', $nsProd)
        if ([string]::IsNullOrEmpty($path)) { $path = $c.GetAttribute('p:path') }
        if ([string]::IsNullOrEmpty($path)) { $path = $c.GetAttribute('path') }
        if (-not [string]::IsNullOrEmpty($path)) {
            if (-not $path.StartsWith('/')) { $path = '/' + $path }
            $preservedModelPaths.Add($path) | Out-Null
        }
    }
}

if (Test-Path $objectsDir) {
    foreach ($f in (Get-ChildItem -Path $objectsDir -Filter '*.model')) {
        $checkPath = '/3D/Objects/' + $f.Name
        if (-not $preservedModelPaths.Contains($checkPath)) {
            Remove-Item $f.FullName -Force
        }
    }
}

if ($null -ne $relsPath) {
    $relsXml = "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
    $relIdx = 1
    foreach ($path in $preservedModelPaths) {
        $relsXml += "<Relationship Target='$path' Id='rel-bod-$relIdx' Type='http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel'/>"
        $relIdx++
    }
    $relsXml += "</Relationships>"
    [System.IO.File]::WriteAllText($relsPath, $relsXml, (New-Object System.Text.UTF8Encoding($false)))
}

# -- Save modified XMLs -----------------------------------------------------
Save-Xml $xml $modelFile
if ($hasSettings) { Save-Xml $settings $settingsPath }

# -- Delete plate image/pick files so Bambu re-slices cleanly ---------------
foreach ($f in (Get-ChildItem -Path (Join-Path $tempDir 'Metadata') -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^(plate|pick)_.*\.(png|json)$' })) {
    Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
}

# -- Repack to OutputPath ---------------------------------------------------
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
try {
    Get-ChildItem $tempDir -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($tempDir.Length).TrimStart('\','/').Replace('\','/')
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
    }
} finally {
    $zip.Dispose()
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}

Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# -- Optional Bambu resave to clean stale metadata --------------------------
if (Test-Path $BambuPath) {
    Write-Host "[BOD] Running Bambu resave to clean metadata..." -NoNewline
    $tempResave = $OutputPath + ".resave.tmp.3mf"
    $logOut     = Join-Path $env:TEMP "bambu_bod_out.txt"
    $logErr     = Join-Path $env:TEMP "bambu_bod_err.txt"
    $procArgs   = "--debug 3 --no-check --uptodate --allow-newer-file --export-3mf `"$tempResave`" `"$OutputPath`""
    $proc       = Start-Process -FilePath $BambuPath -ArgumentList $procArgs `
                                -RedirectStandardOutput $logOut -RedirectStandardError $logErr `
                                -WindowStyle Hidden -PassThru
    $proc.WaitForExit()
    foreach ($log in @($logOut, $logErr)) {
        if (Test-Path $log) { Remove-Item $log -Force -ErrorAction SilentlyContinue }
    }
    if (Test-Path $tempResave) { Move-Item $tempResave $OutputPath -Force; Write-Host " done." }
    else                       { Write-Host " skipped (no output)." }
}

Write-Host "[BOD] BOD.3mf created: $OutputPath"
