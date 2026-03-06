param(
    [string]$WorkDir,
    [string]$InputPath,
    [string]$OutputPath,
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  merge_3mf_worker.ps1 - AUTO-PLAN N-WAY MERGE (PICK BUFFER IMAGE PURGE)
# ════════════════════════════════════════════════════════════════════════════════

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

# ── Transform math ────────────────────────────────────────────────────────────
function Parse-Tx([string]$s) { if ([string]::IsNullOrWhiteSpace($s)) { return [double[]](1,0,0, 0,1,0, 0,0,1, 0,0,0) }; return [double[]]($s.Trim() -split '\s+') }
function Fmt-Tx([double[]]$v) { return ($v | ForEach-Object { $_.ToString('G15') }) -join ' ' }
function Fmt-Matrix([double[]]$t) {
    $r00 = $t[0].ToString('G15'); $r10 = $t[1].ToString('G15'); $r20 = $t[2].ToString('G15')
    $r01 = $t[3].ToString('G15'); $r11 = $t[4].ToString('G15'); $r21 = $t[5].ToString('G15')
    $r02 = $t[6].ToString('G15'); $r12 = $t[7].ToString('G15'); $r22 = $t[8].ToString('G15')
    $tx  = $t[9].ToString('G15'); $ty  = $t[10].ToString('G15'); $tz = $t[11].ToString('G15')
    return "$r00 $r01 $r02 $tx $r10 $r11 $r12 $ty $r20 $r21 $r22 $tz 0 0 0 1"
}

function Mul-Tx([double[]]$A, [double[]]$B) {
    $a00=$A[0]; $a01=$A[3]; $a02=$A[6]; $atx=$A[9]; $a10=$A[1]; $a11=$A[4]; $a12=$A[7]; $aty=$A[10]; $a20=$A[2]; $a21=$A[5]; $a22=$A[8]; $atz=$A[11]
    $b00=$B[0]; $b01=$B[3]; $b02=$B[6]; $btx=$B[9]; $b10=$B[1]; $b11=$B[4]; $b12=$B[7]; $bty=$B[10]; $b20=$B[2]; $b21=$B[5]; $b22=$B[8]; $btz=$B[11]
    $c00 = ($a00 * $b00) + ($a01 * $b10) + ($a02 * $b20); $c10 = ($a10 * $b00) + ($a11 * $b10) + ($a12 * $b20); $c20 = ($a20 * $b00) + ($a21 * $b10) + ($a22 * $b20)
    $c01 = ($a00 * $b01) + ($a01 * $b11) + ($a02 * $b21); $c11 = ($a10 * $b01) + ($a11 * $b11) + ($a12 * $b21); $c21 = ($a20 * $b01) + ($a21 * $b11) + ($a22 * $b21)
    $c02 = ($a00 * $b02) + ($a01 * $b12) + ($a02 * $b22); $c12 = ($a10 * $b02) + ($a11 * $b12) + ($a12 * $b22); $c22 = ($a20 * $b02) + ($a21 * $b12) + ($a22 * $b22)
    $ctx = ($a00 * $btx) + ($a01 * $bty) + ($a02 * $btz) + $atx; $cty = ($a10 * $btx) + ($a11 * $bty) + ($a12 * $btz) + $aty; $ctz = ($a20 * $btx) + ($a21 * $bty) + ($a22 * $btz) + $atz
    return [double[]] @($c00, $c10, $c20, $c01, $c11, $c21, $c02, $c12, $c22, $ctx, $cty, $ctz)
}
function Inv-Tx([double[]]$A) {
    $ir00=$A[0]; $ir01=$A[1]; $ir02=$A[2]; $ir10=$A[3]; $ir11=$A[4]; $ir12=$A[5]; $ir20=$A[6]; $ir21=$A[7]; $ir22=$A[8]; $tx=$A[9]; $ty=$A[10]; $tz=$A[11]
    return [double[]]($ir00,$ir10,$ir20, $ir01,$ir11,$ir21, $ir02,$ir12,$ir22, -($ir00*$tx + $ir01*$ty + $ir02*$tz), -($ir10*$tx + $ir11*$ty + $ir12*$tz), -($ir20*$tx + $ir21*$ty + $ir22*$tz))
}
function Save-Xml([xml]$doc, [string]$path) {
    $settings = New-Object System.Xml.XmlWriterSettings; $settings.Encoding = New-Object System.Text.UTF8Encoding($false); $settings.Indent = $true
    $w = [System.Xml.XmlWriter]::Create($path, $settings); $doc.Save($w); $w.Close()
}
function Find-File([string]$base, [string]$rel) {
    $p = Join-Path $base $rel; if (Test-Path $p) { return $p }; $p2 = Join-Path $base ($rel -replace '\\','/'); if (Test-Path $p2) { return $p2 }; return $null
}

# ── Locate Files ──────────────────────────────────────────────────────────────
$modelFile = (Get-ChildItem -Path $WorkDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
$objectsDir = Join-Path $WorkDir '3D/Objects'
$relsPath = Find-File $WorkDir '3D/_rels/3dmodel.model.rels'
$settingsPath = Find-File $WorkDir 'Metadata/model_settings.config'
$cutInfoPath = Find-File $WorkDir 'Metadata/cut_information.xml'
$sliceInfoPath = Find-File $WorkDir 'Metadata/slice_info.config'

$allModelFiles = @(Get-ChildItem -Path $objectsDir -Filter '*.model' | Sort-Object { [int]($_.BaseName -replace 'object_','') })

# ── Parse Documents ───────────────────────────────────────────────────────────
[xml]$xml = [System.IO.File]::ReadAllText($modelFile, [System.Text.Encoding]::UTF8)
$xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

$buildItems = @($xml.SelectNodes('//m:build/m:item', $xns))
$objById = @{}; foreach ($o in $xml.SelectNodes('//m:resources/m:object', $xns)) { $objById[$o.GetAttribute('id')] = $o }

$hasSettings = ($null -ne $settingsPath) -and (Test-Path $settingsPath)
[xml]$settings = $null; $settObjById = @{}
if ($hasSettings) {
    $settings = [xml][System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)
    foreach ($node in $settings.config.ChildNodes) { if ($node.LocalName -eq 'object') { $settObjById[$node.GetAttribute('id')] = $node } }
}

$report = New-Object System.Collections.Generic.List[string]

# ── Outlier Detection (Isolate Version Text) ──────────────────────────────────
$mergeItems   = @()
$ignoredItems = @()
$fcMap = @{}

foreach ($item in $buildItems) {
    $id = $item.GetAttribute('objectid')
    $fc = "unknown"
    if ($hasSettings -and $null -ne $settObjById[$id]) {
        $fcNode = $settObjById[$id].SelectSingleNode('metadata[@face_count]')
        if ($null -ne $fcNode) { $fc = $fcNode.GetAttribute('face_count') }
    }
    if (-not $fcMap.Contains($fc)) { $fcMap[$fc] = 0 }
    $fcMap[$fc]++
}
$majorityFc = ($fcMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Name

foreach ($item in $buildItems) {
    $id = $item.GetAttribute('objectid')
    $isTarget = $true

    if ($hasSettings -and $null -ne $settObjById[$id]) {
        $fcNode = $settObjById[$id].SelectSingleNode('metadata[@face_count]')
        $fc = if ($null -ne $fcNode) { $fcNode.GetAttribute('face_count') } else { "unknown" }
        if ($fc -ne $majorityFc) { $isTarget = $false }

        $nameNode = $settObjById[$id].SelectSingleNode('metadata[@key="name"]')
        if ($null -ne $nameNode -and $nameNode.GetAttribute('value') -match '(?i)text|version') { $isTarget = $false }
    }
    if ($isTarget) { $mergeItems += $item } else { $ignoredItems += $item }
}

# ── Compute merge plan ────────────────────────────────────────────────────────
function Get-MergePlan([int]$total) {
    $lone = if ($total % 2 -eq 0) { 2 } else { 1 }
    $pool = $total - $lone
    $maxSlots = 64 - $lone - $ignoredItems.Count

    if ($pool -le 0) { return @() }

    for ($g = 2; $g -le $pool; $g++) {
        $period = $g + 1; $bestA = -1; $bestB = -1
        for ($a0 = 0; $a0 -le $g; $a0++) {
            $rem = $pool - $a0 * $g
            if ($rem -lt 0) { break }
            if ($rem % ($g + 1) -ne 0) { continue }
            $a = $a0
            while ($a -le [int]($pool / $g)) {
                $b = ($pool - $a * $g) / ($g + 1)
                if (($a + $b) -le $maxSlots) { if ($a -gt $bestA) { $bestA = $a; $bestB = $b } }
                $a += $period
            }
            break
        }
        if ($bestA -ge 0) {
            $plan = @()
            for ($i = 0; $i -lt $bestA; $i++) { $plan += $g }
            for ($i = 0; $i -lt $bestB; $i++) { $plan += ($g + 1) }
            return $plan
        }
    }
    return @($pool)
}

$totalItems = $mergeItems.Count
$lone       = if ($totalItems % 2 -eq 0) { 2 } else { 1 }
$mergePlan  = @( Get-MergePlan $totalItems )

$report.Add("Total objects on plate: $($buildItems.Count)")
$report.Add("Ignored items (Text/Outliers): $($ignoredItems.Count)")
$report.Add("Items to merge: $totalItems")
$report.Add("Lone (untouched) target items: $lone")
$report.Add("Merge plan: $($mergePlan -join ', ') ($($mergePlan.Count) groups)")

# ── SHARED MESH FILE INITIALIZATION (1-to-1 Memory Optimization) ──────────────
$sharedModelPath = "/3D/Objects/object_1.model"
$sharedDiskPath = Join-Path $objectsDir "object_1.model"

$emptyShell = "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`n<model unit=`"millimeter`" xml:lang=`"en-US`" xmlns=`"http://schemas.microsoft.com/3dmanufacturing/core/2015/02`" xmlns:p=`"http://schemas.microsoft.com/3dmanufacturing/production/2015/06`" requiredextensions=`"p`">`n  <metadata name=`"BambuStudio:3mfVersion`">1</metadata>`n  <resources>`n  </resources>`n  <build/>`n</model>"

$usedModelPaths = New-Object System.Collections.Generic.HashSet[string]
$usedModelPaths.Add($sharedModelPath) | Out-Null

$sourceRelPath = $null
if ($mergeItems.Count -gt 0) {
    $firstId = $mergeItems[0].GetAttribute('objectid')
    $firstObj = $objById[$firstId]
    if ($null -ne $firstObj) {
        $firstComps = $firstObj.SelectNodes('m:components/m:component', $xns)
        if ($firstComps.Count -gt 0) {
            $sourceRelPath = $firstComps[0].GetAttribute('path', $nsProd)
            if ([string]::IsNullOrWhiteSpace($sourceRelPath)) { $sourceRelPath = $firstComps[0].GetAttribute('p:path') }
            if ([string]::IsNullOrWhiteSpace($sourceRelPath)) { $sourceRelPath = $firstComps[0].GetAttribute('path') }
        }
    }
}

if (-not [string]::IsNullOrWhiteSpace($sourceRelPath)) {
    $sourceLocalPath = ($sourceRelPath.TrimStart('/')).Replace('/', '\')
    $sourceDiskPath = Join-Path $WorkDir $sourceLocalPath
    if (Test-Path $sourceDiskPath) {
        Copy-Item -Path $sourceDiskPath -Destination $sharedDiskPath -Force
    }
}

$modelFileCounter = 1

# ── Dynamic Merge Loop ────────────────────────────────────────────────────────
$cursor = 0; $groupCount = 1
$killedIds = New-Object System.Collections.Generic.HashSet[string]
$survivorFaces = @{}
$survivorNames = @{}

$identifyIdCounter = 0
$survivorIdentifyIds = @{}
$objUuidCounter = 1

foreach ($groupSize in $mergePlan) {
    $groupItems = @(); $groupObjs = @(); $groupTxs = @()

    for ($k = 0; $k -lt $groupSize; $k++) {
        $item = $mergeItems[$cursor + $k]
        $id   = $item.GetAttribute('objectid')
        $groupItems += $item; $groupObjs += $objById[$id]; $groupTxs += ,( Parse-Tx ($item.GetAttribute('transform')) )
    }
    $cursor += $groupSize
    $idSurvivor = $groupItems[0].GetAttribute('objectid')

    if ($modelFileCounter -gt 1) {
        $newModelName = "object_${modelFileCounter}.model"
        $newModelPath = "/3D/Objects/$newModelName"
        $newDiskPath = Join-Path $objectsDir $newModelName
        [System.IO.File]::WriteAllText($newDiskPath, $emptyShell, (New-Object System.Text.UTF8Encoding($false)))
        $usedModelPaths.Add($newModelPath) | Out-Null
    }
    $modelFileCounter++

    # Centroid
    [double]$sumX = 0; [double]$sumY = 0; [double]$sumZ = 0
    foreach ($tx in $groupTxs) { $sumX += $tx[9]; $sumY += $tx[10]; $sumZ += $tx[11] }
    [double]$cX = $sumX / $groupSize; [double]$cY = $sumY / $groupSize; [double]$cZ = $sumZ / $groupSize
    [double[]]$txNew = (1,0,0, 0,1,0, 0,0,1, $cX, $cY, $cZ)
    $txNewStr = Fmt-Tx $txNew
    $groupItems[0].SetAttribute('transform', $txNewStr)
    [double[]]$txNew = Parse-Tx $txNewStr
    [double[]]$invTxNew = Inv-Tx $txNew

    $normalComps = @(); $specialComps = @()
    $normalParts = @(); $specialParts = @()

    $compBaseSuffix = "-" + [guid]::NewGuid().ToString().Substring(9)
    $compIndex = $objUuidCounter * 65536

    for ($k = 0; $k -lt $groupSize; $k++) {
        $obj = $groupObjs[$k]; [double[]]$origTx = $groupTxs[$k]
        $memberId = $groupItems[$k].GetAttribute('objectid')

        $cList = $obj.SelectNodes('m:components/m:component', $xns)
        $pList = @()

        if ($hasSettings -and $null -ne $settObjById[$memberId]) {
            $sMember = $settObjById[$memberId]
            $pList = @($sMember.SelectNodes('part'))
        }

        for ($i = 0; $i -lt $cList.Count; $i++) {
            $c = $cList[$i]
            $isSpecial = ($c.GetAttribute('objectid') -eq '40')

            [double[]]$compTx  = Parse-Tx ($c.GetAttribute('transform'))
            [double[]]$bakedTx = Mul-Tx $invTxNew (Mul-Tx $origTx $compTx)

            $newComp = $xml.CreateElement('component', $nsCore)
            $newComp.SetAttribute('path', $nsProd, $sharedModelPath)
            $newComp.SetAttribute('objectid', $c.GetAttribute('objectid'))

            $compUuid = $compIndex.ToString("x8") + $compBaseSuffix
            $newComp.SetAttribute('UUID', $nsProd, $compUuid)
            $compIndex++

            $newComp.SetAttribute('transform', (Fmt-Tx $bakedTx))

            $newPart = $null
            if ($hasSettings -and $i -lt $pList.Count) {
                $newPart = $settings.ImportNode($pList[$i], $true)
                $matNode = $newPart.SelectSingleNode('metadata[@key="matrix"]')
                if ($null -eq $matNode) {
                    $matNode = $settings.CreateElement('metadata')
                    $matNode.SetAttribute('key', 'matrix')
                    $newPart.AppendChild($matNode) | Out-Null
                }
                # FIXED: Writes the true 16-value 4x4 format for Slicer BBox
                $matNode.SetAttribute('value', (Fmt-Matrix $bakedTx))
            }

            if ($isSpecial) {
                $specialComps += $newComp
                if ($null -ne $newPart) { $specialParts += $newPart }
            } else {
                $normalComps += $newComp
                if ($null -ne $newPart) { $normalParts += $newPart }
            }
        }
    }

    [int]$totalFaces = 0
    foreach ($p in $normalParts) {
        $pfcNode = $p.SelectSingleNode('mesh_stat')
        if ($null -ne $pfcNode) {
            [int]$pfc = 0
            if ([int]::TryParse($pfcNode.GetAttribute('face_count'), [ref]$pfc)) { $totalFaces += $pfc }
        }
    }
    foreach ($p in $specialParts) {
        $pfcNode = $p.SelectSingleNode('mesh_stat')
        if ($null -ne $pfcNode) {
            [int]$pfc = 0
            if ([int]::TryParse($pfcNode.GetAttribute('face_count'), [ref]$pfc)) { $totalFaces += $pfc }
        }
    }

    if ($totalFaces -eq 0) {
        for ($k = 0; $k -lt $groupSize; $k++) {
            $memberId = $groupItems[$k].GetAttribute('objectid')
            if ($hasSettings -and $null -ne $settObjById[$memberId]) {
                $mfcNode = $settObjById[$memberId].SelectSingleNode('metadata[@face_count]')
                if ($null -ne $mfcNode) {
                    [int]$mfc = 0
                    if ([int]::TryParse($mfcNode.GetAttribute('face_count'), [ref]$mfc)) { $totalFaces += $mfc }
                }
            }
        }
    }

    $partCount = $normalParts.Count + $specialParts.Count
    if ($partCount -eq 0) { $partCount = 36 * $groupSize }
    $idGap = [math]::Round(442 * ($partCount / 36))
    $identifyIdCounter += $idGap
    $survivorIdentifyIds[$idSurvivor] = $identifyIdCounter

    $survivorFaces[$idSurvivor] = $totalFaces
    $survivorNames[$idSurvivor] = "MergedGroup_$groupSize"

    $newCompsEl = $xml.CreateElement('components', $nsCore)
    foreach ($c in $normalComps) { $newCompsEl.AppendChild($c) | Out-Null }
    foreach ($c in $specialComps) { $newCompsEl.AppendChild($c) | Out-Null }

    if ($null -ne ($oldC = $groupObjs[0].SelectSingleNode('m:components', $xns))) {
        $groupObjs[0].ReplaceChild($newCompsEl, $oldC) | Out-Null
    } else {
        $groupObjs[0].AppendChild($newCompsEl) | Out-Null
    }

    $objUuidStr = $objUuidCounter.ToString("x8") + "-71cb-4c03-9d28-80fed5dfa1dc"
    $groupObjs[0].SetAttribute('UUID', $nsProd, $objUuidStr)
    $objUuidCounter++

    if ($hasSettings) {
        $sSurvivor = $settObjById[$idSurvivor]
        if ($null -ne $sSurvivor) {
            $nameNode = $sSurvivor.SelectSingleNode('metadata[@key="name"]')
            if ($null -ne $nameNode) { $nameNode.SetAttribute('value', $survivorNames[$idSurvivor]) }

            $fcNode = $sSurvivor.SelectSingleNode('metadata[@face_count]')
            if ($null -ne $fcNode) { $fcNode.SetAttribute('face_count', $totalFaces.ToString()) }

            $defaultExtruder = '1'
            $objExtNode = $sSurvivor.SelectSingleNode('metadata[@key="extruder"]')
            if ($null -ne $objExtNode) { $defaultExtruder = $objExtNode.GetAttribute('value') }

            foreach ($ep in @($sSurvivor.SelectNodes('part'))) { $ep.ParentNode.RemoveChild($ep) | Out-Null }

            $partIdCounter = 1
            foreach ($p in $normalParts) {
                $originalSubtype = $p.GetAttribute('subtype')

                $p.SetAttribute('id', $partIdCounter.ToString())
                $p.SetAttribute('subtype', 'normal_part')

                if ($originalSubtype -eq 'modifier_part') {
                    $extNode = $p.SelectSingleNode('metadata[@key="extruder"]')
                    if ($null -eq $extNode) {
                        $extNode = $settings.CreateElement('metadata')
                        $extNode.SetAttribute('key', 'extruder')
                        $p.AppendChild($extNode) | Out-Null
                    }
                    $extNode.SetAttribute('value', $defaultExtruder)
                }

                $pNameNode = $p.SelectSingleNode('metadata[@key="name"]')
                if ($null -ne $pNameNode) { $pNameNode.SetAttribute('value', "MergedPart_$partIdCounter") }
                $sSurvivor.AppendChild($p) | Out-Null
                $partIdCounter++
            }
            foreach ($p in $specialParts) {
                $originalSubtype = $p.GetAttribute('subtype')

                $p.SetAttribute('id', $partIdCounter.ToString())
                $p.SetAttribute('subtype', 'normal_part')

                if ($originalSubtype -eq 'modifier_part') {
                    $extNode = $p.SelectSingleNode('metadata[@key="extruder"]')
                    if ($null -eq $extNode) {
                        $extNode = $settings.CreateElement('metadata')
                        $extNode.SetAttribute('key', 'extruder')
                        $p.AppendChild($extNode) | Out-Null
                    }
                    $extNode.SetAttribute('value', $defaultExtruder)
                }

                $pNameNode = $p.SelectSingleNode('metadata[@key="name"]')
                if ($null -ne $pNameNode) { $pNameNode.SetAttribute('value', "MergedPart_$partIdCounter") }
                $sSurvivor.AppendChild($p) | Out-Null
                $partIdCounter++
            }

            for ($k = 1; $k -lt $groupSize; $k++) {
                $memberId = $groupItems[$k].GetAttribute('objectid')
                $killedIds.Add($memberId) | Out-Null
                $sMember = $settObjById[$memberId]
                if ($null -ne $sMember) { $sMember.ParentNode.RemoveChild($sMember) | Out-Null }
            }

            $assemble = $settings.SelectSingleNode('//assemble')
            if ($null -ne $assemble) {
                $asmSurv = $assemble.SelectSingleNode("assemble_item[@object_id='$idSurvivor']")
                if ($null -ne $asmSurv) { $asmSurv.SetAttribute('transform', "1 0 0 0 1 0 0 0 1 $($txNew[9]) $($txNew[10]) $($txNew[11])") }
                for ($k = 1; $k -lt $groupSize; $k++) {
                    $memberId = $groupItems[$k].GetAttribute('objectid')
                    $asmMember = $assemble.SelectSingleNode("assemble_item[@object_id='$memberId']")
                    if ($null -ne $asmMember) { $assemble.RemoveChild($asmMember) | Out-Null }
                }
            }
        }
    }

    for ($k = 1; $k -lt $groupSize; $k++) {
        $groupItems[$k].ParentNode.RemoveChild($groupItems[$k]) | Out-Null
        $groupObjs[$k].ParentNode.RemoveChild($groupObjs[$k])   | Out-Null
    }
    $groupCount++
}

# ── Ensure Lone Items receive proper Tracking and UUIDs ───────────────────────
$loneCounter = 1
for ($li = ($mergeItems.Count - $lone); $li -lt $mergeItems.Count; $li++) {
    $loneId = $mergeItems[$li].GetAttribute('objectid')
    $loneObj = $objById[$loneId]

    $objUuidStr = $objUuidCounter.ToString("x8") + "-71cb-4c03-9d28-80fed5dfa1dc"
    if ($null -ne $loneObj) { $loneObj.SetAttribute('UUID', $nsProd, $objUuidStr) }
    $objUuidCounter++

    $identifyIdCounter += 442
    $survivorIdentifyIds[$loneId] = $identifyIdCounter

    if ($hasSettings -and ($null -ne ($sLone = $settObjById[$loneId]))) {
        $loneNameNode = $sLone.SelectSingleNode('metadata[@key="name"]')
        if ($null -ne $loneNameNode) {
            $loneNameNode.SetAttribute('value', "$($loneNameNode.GetAttribute('value'))_Lone")
        }
    }

    if ($modelFileCounter -gt 1) {
        $newModelName = "object_${modelFileCounter}.model"
        $newModelPath = "/3D/Objects/$newModelName"
        $newDiskPath = Join-Path $objectsDir $newModelName
        [System.IO.File]::WriteAllText($newDiskPath, $emptyShell, (New-Object System.Text.UTF8Encoding($false)))
        $usedModelPaths.Add($newModelPath) | Out-Null
    }
    $modelFileCounter++

    $loneComps = $loneObj.SelectNodes('m:components/m:component', $xns)
    foreach ($c in $loneComps) {
        $c.SetAttribute('path', $nsProd, $sharedModelPath)
    }

    $loneCounter++
}

# ── Remove killed model_instances BEFORE renumbering (avoids ID collision) ────
if ($hasSettings -and ($plate = $settings.SelectSingleNode('//plate'))) {
    foreach ($inst in @($plate.SelectNodes('model_instance'))) {
        $metaId = $inst.SelectSingleNode('metadata[@key="object_id"]')
        if ($null -ne $metaId -and $killedIds.Contains($metaId.GetAttribute('value'))) {
            $inst.ParentNode.RemoveChild($inst) | Out-Null
        }
    }
}

# ── GLOBAL OBJECT ID RENUMBERING ──────────────────────────────────────────────
$survivingObjects = @($xml.SelectNodes('//m:resources/m:object', $xns)) | Sort-Object { [int]$_.GetAttribute('id') }

$minId = 999999
foreach ($obj in $survivingObjects) {
    $curr = [int]$obj.GetAttribute('id')
    if ($curr -lt $minId) { $minId = $curr }
}

$idMap = @{}
$newIdCounter = $minId

foreach ($obj in $survivingObjects) {
    $oldId = $obj.GetAttribute('id')
    $newId = $newIdCounter.ToString()
    $idMap[$oldId] = $newId
    $obj.SetAttribute('id', $newId)
    $newIdCounter++
}

foreach ($item in @($xml.SelectNodes('//m:build/m:item', $xns))) {
    $oldId = $item.GetAttribute('objectid')
    if ($idMap.Contains($oldId)) { $item.SetAttribute('objectid', $idMap[$oldId]) }
}

if ($hasSettings) {
    foreach ($obj in @($settings.SelectNodes('//*[local-name()="object"]'))) {
        $oldId = $obj.GetAttribute('id')
        if ($idMap.Contains($oldId)) { $obj.SetAttribute('id', $idMap[$oldId]) }
    }
    foreach ($asm in @($settings.SelectNodes('//assemble/assemble_item'))) {
        $oldId = $asm.GetAttribute('object_id')
        if ($idMap.Contains($oldId)) { $asm.SetAttribute('object_id', $idMap[$oldId]) }
    }
    if ($plate = $settings.SelectSingleNode('//plate')) {
        foreach ($meta in @($plate.SelectNodes('model_instance/metadata[@key="object_id"]'))) {
            $oldId = $meta.GetAttribute('value')
            if ($idMap.Contains($oldId)) { $meta.SetAttribute('value', $idMap[$oldId]) }
        }
    }
}

# ── Finalize Metadata ─────────────────────────────────────────────────────────
if ($hasSettings -and ($plate = $settings.SelectSingleNode('//plate'))) {
    foreach ($inst in @($plate.SelectNodes('model_instance'))) {
        $metaId = $inst.SelectSingleNode('metadata[@key="object_id"]')
        if ($null -ne $metaId) {
            $instObjId = $metaId.GetAttribute('value')
            $originalId = $null
            foreach ($key in $idMap.Keys) { if ($idMap[$key] -eq $instObjId) { $originalId = $key; break } }

            if ($null -ne $originalId) {
                if ($killedIds.Contains($originalId)) {
                    $inst.ParentNode.RemoveChild($inst) | Out-Null
                } elseif ($survivorIdentifyIds.Contains($originalId)) {
                    $identifyIdNode = $inst.SelectSingleNode('metadata[@key="identify_id"]')
                    if ($null -ne $identifyIdNode) {
                        $identifyIdNode.SetAttribute('value', $survivorIdentifyIds[$originalId].ToString())
                    }
                }
            }
        }
    }
}

# --- TRIM AND RENUMBER CUT_INFORMATION.XML ARRAY ---
if (($null -ne $cutInfoPath) -and (Test-Path $cutInfoPath)) {
    [xml]$cutXml = [System.IO.File]::ReadAllText($cutInfoPath, [System.Text.Encoding]::UTF8)
    $cutObjs = @($cutXml.SelectNodes('//*[local-name()="object"]'))
    $survivorCount = $survivingObjects.Count

    if ($cutObjs.Count -gt $survivorCount) {
        for ($i = $survivorCount; $i -lt $cutObjs.Count; $i++) {
            $cutObjs[$i].ParentNode.RemoveChild($cutObjs[$i]) | Out-Null
        }
    }

    $cIdx = 1
    foreach ($co in @($cutXml.SelectNodes('//*[local-name()="object"]'))) {
        $co.SetAttribute('id', $cIdx.ToString())
        $cIdx++
    }

    $ws = New-Object System.Xml.XmlWriterSettings; $ws.Encoding = New-Object System.Text.UTF8Encoding($false); $ws.Indent = $true
    $w = [System.Xml.XmlWriter]::Create($cutInfoPath, $ws); $cutXml.Save($w); $w.Close()
}

# --- SYNCHRONIZING THE SLICE CACHE ---
if (($null -ne $sliceInfoPath) -and (Test-Path $sliceInfoPath)) {
    [xml]$sliceXml = [System.IO.File]::ReadAllText($sliceInfoPath, [System.Text.Encoding]::UTF8)
    $sliceModified = $false

    foreach ($sObj in @($sliceXml.SelectNodes('//*[local-name()="object"]'))) {
        $oldId = $sObj.GetAttribute('id')
        if ($killedIds.Contains($oldId)) {
            $sObj.ParentNode.RemoveChild($sObj) | Out-Null
            $sliceModified = $true
        } elseif ($idMap.Contains($oldId)) {
            $newMappedId = $idMap[$oldId]
            $sObj.SetAttribute('id', $newMappedId)
            $sliceModified = $true

            if ($survivorFaces.Contains($oldId)) {
                $fcNode = $sObj.SelectSingleNode('*[local-name()="metadata" and @face_count]')
                if ($null -ne $fcNode) { $fcNode.SetAttribute('face_count', $survivorFaces[$oldId].ToString()) }
                $nameNode = $sObj.SelectSingleNode('*[local-name()="metadata" and @key="name"]')
                if ($null -ne $nameNode) { $nameNode.SetAttribute('value', $survivorNames[$oldId]) }

                if ($hasSettings) {
                    $matchingSettingsObj = $settings.SelectSingleNode("//*[local-name()='object' and @id='$newMappedId']")
                    if ($null -ne $matchingSettingsObj) {
                        foreach ($sp in @($sObj.SelectNodes('*[local-name()="part"]'))) {
                            $sp.ParentNode.RemoveChild($sp) | Out-Null
                        }
                        foreach ($mp in @($matchingSettingsObj.SelectNodes('*[local-name()="part"]'))) {
                            $importedPart = $sliceXml.ImportNode($mp, $true)
                            $sObj.AppendChild($importedPart) | Out-Null
                        }
                    }
                }
            }
        }
    }

    foreach ($inst in @($sliceXml.SelectNodes('//*[local-name()="model_instance"]'))) {
        $metaId = $inst.SelectSingleNode('*[local-name()="metadata" and @key="object_id"]')
        if ($null -ne $metaId) {
            $oldId = $metaId.GetAttribute('value')
            if ($killedIds.Contains($oldId)) {
                $inst.ParentNode.RemoveChild($inst) | Out-Null
                $sliceModified = $true
            } elseif ($idMap.Contains($oldId)) {
                $metaId.SetAttribute('value', $idMap[$oldId])
                $sliceModified = $true
                if ($survivorIdentifyIds.Contains($oldId)) {
                    $identifyIdNode = $inst.SelectSingleNode('*[local-name()="metadata" and @key="identify_id"]')
                    if ($null -ne $identifyIdNode) {
                        $identifyIdNode.SetAttribute('value', $survivorIdentifyIds[$oldId].ToString())
                    }
                }
            }
        }
    }
    if ($sliceModified) { Save-Xml $sliceXml $sliceInfoPath }
}

# --- REFRESH DICTIONARY BEFORE GC BUG FIX ---
$objById = @{}
foreach ($o in $xml.SelectNodes('//m:resources/m:object', $xns)) {
    $objById[$o.GetAttribute('id')] = $o
}

# ── GEOMETRY GARBAGE COLLECTION & RELS REBUILD ────────────────────────────────
$allModelFiles = Get-ChildItem -Path $objectsDir -Filter '*.model'
foreach ($f in $allModelFiles) {
    $checkPath = "/3D/Objects/" + $f.Name
    if (-not $usedModelPaths.Contains($checkPath)) {
        Remove-Item $f.FullName -Force
    }
}

# ── PURGE STALE UI CACHES (The Pick Buffer Fix) ───────────────────────────────
Get-ChildItem -Path (Join-Path $WorkDir "Metadata") -Filter "pick_*.png" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -Path (Join-Path $WorkDir "Metadata") -Filter "plate_*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

Save-Xml $xml $modelFile
if ($hasSettings) { Save-Xml $settings $settingsPath }

$relsLines = @('<?xml version="1.0" encoding="UTF-8"?>')
$relsLines += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
$relIdx = 1
foreach ($path in $usedModelPaths) {
    $relsLines += "  <Relationship Target=`"$path`" Id=`"rel-$relIdx`" Type=`"http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel`"/>"
    $relIdx++
}
$relsLines += '</Relationships>'
[System.IO.File]::WriteAllText($relsPath, ($relsLines -join "`r`n"), (New-Object System.Text.UTF8Encoding($false)))

# ── Repack ────────────────────────────────────────────────────────────────────
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
Get-ChildItem $WorkDir -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($WorkDir.Length).TrimStart('\','/').Replace('\','/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel)
}
$zip.Dispose()

$finalCount = $mergePlan.Count + $lone
$report.Add("Final object count: $finalCount")
$report | Set-Content -Path $ReportPath -Encoding UTF8
Write-Host "Success! Ignored: $($ignoredItems.Count). Merged: $groupCount groups. Final Target Count: $finalCount."