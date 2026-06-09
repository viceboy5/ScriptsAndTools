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

# -- Geometry helpers for computing a true assembled-mesh footprint -------------
# (plate_1.json's bbox can be stale for objects that were isolated/regrouped after
#  it was last written - e.g. multi-part "colorcut" assemblies - so we walk the
#  actual mesh geometry through the component transform chain instead.)
function Combine-Tx([double[]]$inner, [double[]]$outer) {
    # Equivalent to applying $inner then $outer to a row-vector point: v * inner * outer
    $r = New-Object double[] 12
    for ($i = 0; $i -lt 3; $i++) {
        for ($j = 0; $j -lt 3; $j++) {
            $r[$i * 3 + $j] = $inner[$i*3+0]*$outer[0*3+$j] + $inner[$i*3+1]*$outer[1*3+$j] + $inner[$i*3+2]*$outer[2*3+$j]
        }
    }
    for ($j = 0; $j -lt 3; $j++) {
        $r[9 + $j] = $inner[9]*$outer[0*3+$j] + $inner[10]*$outer[1*3+$j] + $inner[11]*$outer[2*3+$j] + $outer[9+$j]
    }
    return [double[]]$r
}
function Apply-Tx([double[]]$t, [double]$x, [double]$y, [double]$z) {
    return @(
        ($x*$t[0] + $y*$t[3] + $z*$t[6] + $t[9]),
        ($x*$t[1] + $y*$t[4] + $z*$t[7] + $t[10])
    )
}
function Get-MeshVertices([string]$modelPath) {
    $cache = @{}
    [xml]$mxml = [System.IO.File]::ReadAllText($modelPath, [System.Text.Encoding]::UTF8)
    foreach ($obj in $mxml.SelectNodes('//*[local-name()="object"]')) {
        $verts = New-Object System.Collections.Generic.List[double[]]
        foreach ($v in $obj.SelectNodes('.//*[local-name()="vertex"]')) {
            $verts.Add([double[]]@([double]$v.GetAttribute('x'), [double]$v.GetAttribute('y'), [double]$v.GetAttribute('z')))
        }
        $cache[$obj.GetAttribute('id')] = $verts
    }
    return $cache
}
function Update-Bounds([hashtable]$b, [double]$x, [double]$y) {
    if ($x -lt $b.MinX) { $b.MinX = $x }
    if ($x -gt $b.MaxX) { $b.MaxX = $x }
    if ($y -lt $b.MinY) { $b.MinY = $y }
    if ($y -gt $b.MaxY) { $b.MaxY = $y }
}
function Walk-FootprintObject([System.Xml.XmlNode]$objNode, [double[]]$accTx, [hashtable]$bounds, [hashtable]$meshCache, [xml]$xml, $xns, [string]$tempDir, [string]$nsProd) {
    $components = @($objNode.SelectNodes('m:components/m:component', $xns))
    if ($components.Count -gt 0) {
        foreach ($c in $components) {
            $path = $c.GetAttribute('path', $nsProd)
            if ([string]::IsNullOrEmpty($path)) { $path = $c.GetAttribute('p:path') }
            if ([string]::IsNullOrEmpty($path)) { $path = $c.GetAttribute('path') }
            $compTx     = Parse-Tx ($c.GetAttribute('transform'))
            $childId    = $c.GetAttribute('objectid')
            $combinedTx = Combine-Tx $compTx $accTx

            if ([string]::IsNullOrEmpty($path)) {
                $childNode = $xml.SelectSingleNode("//m:object[@id='$childId']", $xns)
                if ($null -ne $childNode) { Walk-FootprintObject $childNode $combinedTx $bounds $meshCache $xml $xns $tempDir $nsProd }
            } else {
                $resolvedPath = $path.TrimStart('/')
                if (-not $meshCache.ContainsKey($resolvedPath)) {
                    $resourcePath = Find-File $tempDir $resolvedPath
                    if ($null -eq $resourcePath) { continue }
                    $meshCache[$resolvedPath] = Get-MeshVertices $resourcePath
                }
                $verts = $meshCache[$resolvedPath][$childId]
                if ($null -eq $verts) { continue }
                foreach ($v in $verts) {
                    $p = Apply-Tx $combinedTx $v[0] $v[1] $v[2]
                    Update-Bounds $bounds $p[0] $p[1]
                }
            }
        }
    } else {
        foreach ($v in $objNode.SelectNodes('.//*[local-name()="vertex"]')) {
            $p = Apply-Tx $accTx ([double]$v.GetAttribute('x')) ([double]$v.GetAttribute('y')) ([double]$v.GetAttribute('z'))
            Update-Bounds $bounds $p[0] $p[1]
        }
    }
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
    $TARGET_GAP = 3.0   # fixed spacing between adjacent copies (mm)

    Write-Host "[BOD-Grid] Extracting Final.3mf..."
    $tempDir = Join-Path $env:TEMP "bod_grid_$([guid]::NewGuid().ToString().Substring(0,8))"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    try { [System.IO.Compression.ZipFile]::ExtractToDirectory($InputPath, $tempDir) }
    catch { Write-Host "[BOD-Grid] ERROR extracting: $_" -ForegroundColor Red; Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue; exit 1 }

    # -- Parse 3dmodel.model -------------------------------------------------------
    $modelFile = (Get-ChildItem -Path $tempDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
    [xml]$xml  = [System.IO.File]::ReadAllText($modelFile, [System.Text.Encoding]::UTF8)
    $xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

    $buildNode     = $xml.SelectSingleNode('//m:build', $xns)
    $origBuildItems = @($buildNode.SelectNodes('m:item', $xns))
    Write-Host "[BOD-Grid] Original build items: $($origBuildItems.Count)"

    # Store original transforms indexed by position
    $origTxList = @($origBuildItems | ForEach-Object { , (Parse-Tx ($_.GetAttribute('transform'))) })

    # -- Compute the true assembled footprint by walking the mesh geometry --------
    # plate_1.json's bbox can be stale (e.g. left over from a larger merged plate before
    # isolate_final_worker repositioned/regrouped the design), which under/over-states the
    # real XY extent of multi-part "colorcut" assemblies. Walk every build item's object
    # through its component transform chain and union the transformed mesh vertex bounds.
    $bounds = @{ MinX = [double]::MaxValue; MaxX = [double]::MinValue; MinY = [double]::MaxValue; MaxY = [double]::MinValue }
    $meshCache = @{}
    for ($bi = 0; $bi -lt $origBuildItems.Count; $bi++) {
        $rootObj = $xml.SelectSingleNode("//m:object[@id='$($origBuildItems[$bi].GetAttribute('objectid'))']", $xns)
        if ($null -ne $rootObj) { Walk-FootprintObject $rootObj $origTxList[$bi] $bounds $meshCache $xml $xns $tempDir $nsProd }
    }
    if ($bounds.MinX -eq [double]::MaxValue) {
        Write-Host "[BOD-Grid] ERROR: could not compute design footprint from mesh geometry." -ForegroundColor Red
        Remove-Item $tempDir -Recurse -Force; exit 1
    }

    $bboxW  = $bounds.MaxX - $bounds.MinX
    $bboxD  = $bounds.MaxY - $bounds.MinY
    $origCX = ($bounds.MinX + $bounds.MaxX) / 2.0
    $origCY = ($bounds.MinY + $bounds.MaxY) / 2.0

    Write-Host ("[BOD-Grid] Design footprint: {0:F2} x {1:F2} mm  (centre {2:F2}, {3:F2})" -f $bboxW, $bboxD, $origCX, $origCY)

    # -- Choose grid layout: lay out at a fixed $TARGET_GAP (don't stretch to fill the plate) --
    # Pick the cols x rows split that wastes the fewest empty cells while fitting in $avail at
    # $TARGET_GAP spacing; ties broken toward the most square-ish shape, then toward wider/fewer
    # rows (e.g. 5 cols x 3 rows over 3 cols x 5 rows). Leftover plate space just becomes extra
    # margin once the block is centred below - it does not inflate the gaps.
    $avail = $PLATE_W - 2.0 * $MARGIN
    $best  = $null; $bestWaste = [int]::MaxValue; $bestSquareness = [double]::MaxValue

    for ($cols = 1; $cols -le $COPIES; $cols++) {
        $rows = [int][Math]::Ceiling($COPIES / $cols)
        $gridWNeeded = $cols * $bboxW + [Math]::Max(0, $cols - 1) * $TARGET_GAP
        $gridHNeeded = $rows * $bboxD + [Math]::Max(0, $rows - 1) * $TARGET_GAP
        if ($gridWNeeded -gt $avail -or $gridHNeeded -gt $avail) { continue }

        $waste = ($cols * $rows) - $COPIES
        $squareness = [Math]::Abs($cols - $rows)
        $better = $false
        if ($waste -lt $bestWaste) { $better = $true }
        elseif ($waste -eq $bestWaste -and $squareness -lt $bestSquareness) { $better = $true }
        elseif ($waste -eq $bestWaste -and $squareness -eq $bestSquareness -and $null -ne $best -and $cols -gt $best.Cols) { $better = $true }

        if ($better) {
            $bestWaste = $waste; $bestSquareness = $squareness
            $best = @{ Cols = $cols; Rows = $rows; GapX = $TARGET_GAP; GapY = $TARGET_GAP }
        }
    }
    if ($null -eq $best) {
        Write-Host "[BOD-Grid] ERROR: design too large to fit $COPIES copies on a ${PLATE_W}x${PLATE_H} plate with ${MARGIN}mm margins at ${TARGET_GAP}mm spacing." -ForegroundColor Red
        Remove-Item $tempDir -Recurse -Force; exit 1
    }
    Write-Host ("[BOD-Grid] Layout: {0} cols x {1} rows  gap {2:F1}mm x {3:F1}mm  ({4} empty cell(s))" -f $best.Cols, $best.Rows, $best.GapX, $best.GapY, $bestWaste)

    # -- Centre the grid block on the plate (instead of anchoring to the margin corner) --
    $gridW   = $best.Cols * $bboxW + ($best.Cols - 1) * $best.GapX
    $gridH   = $best.Rows * $bboxD + ($best.Rows - 1) * $best.GapY
    $offsetX = ($PLATE_W - $gridW) / 2.0
    $offsetY = ($PLATE_H - $gridH) / 2.0

    # -- Nudge the grid clear of the printer's bed exclusion zone (bottom-left corner) --
    # X1C/X1/P1/P2/A1-family beds reserve a rectangle at the plate origin (e.g. 18 x 28 mm on the X1C)
    # for the wiper/handle. Only the corner cell can ever overlap it - if it would, push the whole
    # grid just far enough along the cheaper axis to clear it (rectangles no longer overlap once
    # separated on either axis), rather than uniformly inflating the margin and wasting plate space.
    $EXCLUDE_W = 18.0; $EXCLUDE_H = 28.0
    if ($offsetX -lt $EXCLUDE_W -and $offsetY -lt $EXCLUDE_H) {
        $shiftX = $EXCLUDE_W - $offsetX
        $shiftY = $EXCLUDE_H - $offsetY
        if ($shiftX -le $shiftY) { $offsetX = $EXCLUDE_W } else { $offsetY = $EXCLUDE_H }
        Write-Host ("[BOD-Grid] Corner cell would overlap the bed exclusion zone - nudged offset to ({0:F1}, {1:F1})" -f $offsetX, $offsetY) -ForegroundColor Yellow
    }
    Write-Host ("[BOD-Grid] Grid block: {0:F1} x {1:F1} mm, offset ({2:F1}, {3:F1})" -f $gridW, $gridH, $offsetX, $offsetY)

    # Remove all existing build items (we'll re-add them as copies)
    foreach ($item in $origBuildItems) { $buildNode.RemoveChild($item) | Out-Null }

    # -- Build 15 copies ----------------------------------------------------------
    $idx = 0
    for ($row = 0; $row -lt $best.Rows; $row++) {
        for ($col = 0; $col -lt $best.Cols; $col++) {
            if ($idx -ge $COPIES) { break }

            $newCX  = $offsetX + $col * ($bboxW + $best.GapX) + $bboxW / 2.0
            $newCY  = $offsetY + $row * ($bboxD + $best.GapY) + $bboxD / 2.0
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

    # -- Reposition the prime tower to the right of the grid ----------------------
    # Place it just to the right of the grid block, near the bottom of the plate
    # (above calibration lines), and clamp it so it never overruns the right edge.
    $TOWER_GAP          = 5.0    # mm clearance between grid right edge and tower left edge
    $TOWER_Y            = 15.0   # mm from plate front - above calibration lines
    $TOWER_EDGE_MARGIN  = 5.0    # minimum clearance from the right plate edge
    $psCfgPath = Find-File $tempDir 'Metadata/project_settings.config'
    if ($null -ne $psCfgPath) {
        $psCfgText = [System.IO.File]::ReadAllText($psCfgPath, [System.Text.Encoding]::UTF8)

        # Read prime_tower_width from the config (default 35mm if key absent)
        $towerWidth = 35.0
        if ($psCfgText -match '"prime_tower_width"\s*:\s*"([0-9.]+)"') {
            try { $towerWidth = [double]$Matches[1] } catch {}
        }

        $gridRightEdge = $offsetX + $gridW
        $towerX = $gridRightEdge + $TOWER_GAP

        # Clamp: ensure tower right edge stays $TOWER_EDGE_MARGIN away from the plate edge
        $towerMaxX = $PLATE_W - $TOWER_EDGE_MARGIN - $towerWidth
        if ($towerX -gt $towerMaxX) {
            Write-Host ("[BOD-Grid] Prime tower clamped left: X {0:F1} -> {1:F1} (tower {2:F0}mm wide, plate {3:F0}mm)" -f $towerX, $towerMaxX, $towerWidth, $PLATE_W) -ForegroundColor Yellow
            $towerX = $towerMaxX
        }

        $towerXStr = $towerX.ToString('F4', [System.Globalization.CultureInfo]::InvariantCulture)
        $towerYStr = $TOWER_Y.ToString('F4', [System.Globalization.CultureInfo]::InvariantCulture)

        # Patch the two keys in-place; handles both compact and indented JSON
        $psCfgText = $psCfgText -replace '"wipe_tower_x"\s*:\s*\[[^\]]*\]', ('"wipe_tower_x": ["' + $towerXStr + '"]')
        $psCfgText = $psCfgText -replace '"wipe_tower_y"\s*:\s*\[[^\]]*\]', ('"wipe_tower_y": ["' + $towerYStr + '"]')

        [System.IO.File]::WriteAllText($psCfgPath, $psCfgText, [System.Text.Encoding]::UTF8)
        Write-Host ("[BOD-Grid] Prime tower -> X={0:F1}  Y={1:F1}  width={2:F0}mm  right edge at {3:F1}mm" -f $towerX, $TOWER_Y, $towerWidth, ($towerX + $towerWidth))
    } else {
        Write-Host "[BOD-Grid] WARNING: project_settings.config not found - prime tower position not updated." -ForegroundColor Yellow
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
