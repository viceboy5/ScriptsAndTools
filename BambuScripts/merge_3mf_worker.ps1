param(
    [string]$WorkDir,
    [string]$InputPath,
    [string]$OutputPath,
    [string]$ReportPath,
    [string]$DoColors = "0"
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  merge_3mf_worker.ps1 - BASELINE REVERT (PURGE & MERGE ONLY)
# ════════════════════════════════════════════════════════════════════════════════

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

# ── Transform math ────────────────────────────────────────────────────────────
function Parse-Tx([string]$s) { if ([string]::IsNullOrWhiteSpace($s)) { return [double[]](1,0,0, 0,1,0, 0,0,1, 0,0,0) }; return [double[]]($s.Trim() -split '\s+') }
function Fmt-Tx([double[]]$v) { return ($v | ForEach-Object { $_.ToString('G15') }) -join ' ' }
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

# ── PURGE OFF-PLATE & ORPHANED OBJECTS ────────────────────────────────────────
$validBuildItems = @()
$killedIds = New-Object System.Collections.Generic.HashSet[string]

foreach ($item in $buildItems) {
    $id = $item.GetAttribute('objectid')
    [double[]]$tx = Parse-Tx ($item.GetAttribute('transform'))
    $x = $tx[9]; $y = $tx[10]

    if ($x -lt -50 -or $x -gt 300 -or $y -lt -50 -or $y -gt 300) {
        $item.ParentNode.RemoveChild($item) | Out-Null
    } else {
        $validBuildItems += $item
    }
}
$buildItems = $validBuildItems

$protectedIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($item in $buildItems) { $protectedIds.Add($item.GetAttribute('objectid')) | Out-Null }

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
        $obj.ParentNode.RemoveChild($obj) | Out-Null
        $objById.Remove($objId)

        if ($hasSettings) {
            if ($null -ne $settObjById[$objId]) {
                $settObjById[$objId].ParentNode.RemoveChild($settObjById[$objId]) | Out-Null
            }
            foreach ($node in $settings.SelectNodes("//*[@object_id='$objId']")) {
                $node.ParentNode.RemoveChild($node) | Out-Null
            }
        }
    }
}

# ── ID Diagnostic Output ────────────────────────────────────────────────────
$survivingIds = @($buildItems | ForEach-Object { $_.GetAttribute('objectid') }) | Sort-Object { [int]$_ }
$killedList   = @($killedIds)  | Sort-Object { [int]$_ }

Write-Host ""
Write-Host "  -- ID REPORT ----------------------------------------------------------"
Write-Host "  Removed (off-plate and orphaned) [$($killedList.Count)]: $($killedList -join ', ')"
Write-Host "  Surviving on-plate IDs [$($survivingIds.Count)]: $($survivingIds -join ', ')"
Write-Host "  -----------------------------------------------------------------------"
Write-Host ""

# ── Outlier Detection (Isolate Version Text) ──────────────────────────────────
$report = New-Object System.Collections.Generic.List[string]
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

# ── DYNAMIC MESH FILE TRACKER ─────────────────────────────────────────────────
$emptyShell = "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`n<model unit=`"millimeter`" xml:lang=`"en-US`" xmlns=`"http://schemas.microsoft.com/3dmanufacturing/core/2015/02`" xmlns:p=`"http://schemas.microsoft.com/3dmanufacturing/production/2015/06`" requiredextensions=`"p`">`n  <metadata name=`"BambuStudio:3mfVersion`">1</metadata>`n  <resources>`n  </resources>`n  <build/>`n</model>"

$sourceToMasterMap = @{}
$usedModelPaths = New-Object System.Collections.Generic.HashSet[string]
$modelFileCounter = 1

# ── Dynamic Merge Loop ────────────────────────────────────────────────────────
$cursor = 0; $groupCount = 1
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

    $firstItemComps = $groupObjs[0].SelectNodes('m:components/m:component', $xns)
    $sourceRelPath = $firstItemComps[0].GetAttribute('path', $nsProd)
    if ([string]::IsNullOrWhiteSpace($sourceRelPath)) { $sourceRelPath = $firstItemComps[0].GetAttribute('p:path') }
    if ([string]::IsNullOrWhiteSpace($sourceRelPath)) { $sourceRelPath = $firstItemComps[0].GetAttribute('path') }

    $sourceLocalPath = ($sourceRelPath.TrimStart('/')).Replace('/', '\')
    $sourceDiskPath = Join-Path $WorkDir $sourceLocalPath

    $newModelName = "object_${modelFileCounter}.model"
    $newModelPath = "/3D/Objects/$newModelName"
    $newDiskPath = Join-Path $objectsDir $newModelName

    if (-not $sourceToMasterMap.Contains($sourceRelPath)) {
        $normSource = [System.IO.Path]::GetFullPath($sourceDiskPath)
        $normDest = [System.IO.Path]::GetFullPath($newDiskPath)
        if ($normSource -ne $normDest) {
            Copy-Item -Path $normSource -Destination $normDest -Force
        }
        $sourceToMasterMap[$sourceRelPath] = $newModelPath
        $masterPathForThisGroup = $newModelPath
    } else {
        [System.IO.File]::WriteAllText($newDiskPath, $emptyShell, (New-Object System.Text.UTF8Encoding($false)))
        $masterPathForThisGroup = $sourceToMasterMap[$sourceRelPath]
    }
    $usedModelPaths.Add($newModelPath) | Out-Null
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

    $mergedComps = @()
    $mergedParts = @()

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

            [double[]]$compTx  = Parse-Tx ($c.GetAttribute('transform'))
            [double[]]$bakedTx = Mul-Tx $invTxNew (Mul-Tx $origTx $compTx)

            $newComp = $xml.CreateElement('component', $nsCore)
            $newComp.SetAttribute('path', $nsProd, $masterPathForThisGroup) | Out-Null
            $newComp.SetAttribute('objectid', $c.GetAttribute('objectid'))

            $compUuid = $compIndex.ToString("x8") + $compBaseSuffix
            $newComp.SetAttribute('UUID', $nsProd, $compUuid) | Out-Null
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
                $matNode.SetAttribute('value', (Fmt-Tx $bakedTx))
            }

            $mergedComps += $newComp
            if ($null -ne $newPart) { $mergedParts += $newPart }
        }
    }

    [int]$totalFaces = 0
    foreach ($p in $mergedParts) {
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

    $partCount = $mergedParts.Count
    if ($partCount -eq 0) { $partCount = 36 * $groupSize }
    $idGap = [math]::Round(442 * ($partCount / 36))
    $identifyIdCounter += $idGap
    $survivorIdentifyIds[$idSurvivor] = $identifyIdCounter

    $survivorFaces[$idSurvivor] = $totalFaces
    $survivorNames[$idSurvivor] = "MergedGroup_$groupSize"

    $newCompsEl = $xml.CreateElement('components', $nsCore)
    foreach ($c in $mergedComps) { $newCompsEl.AppendChild($c) | Out-Null }

    if ($null -ne ($oldC = $groupObjs[0].SelectSingleNode('m:components', $xns))) {
        $groupObjs[0].ReplaceChild($newCompsEl, $oldC) | Out-Null
    } else {
        $groupObjs[0].AppendChild($newCompsEl) | Out-Null
    }

    $objUuidStr = $objUuidCounter.ToString("x8") + "-71cb-4c03-9d28-80fed5dfa1dc"
    $groupObjs[0].SetAttribute('UUID', $nsProd, $objUuidStr) | Out-Null
    $objUuidCounter++

    if ($hasSettings) {
        $sSurvivor = $settObjById[$idSurvivor]
        if ($null -ne $sSurvivor) {
            $nameNode = $sSurvivor.SelectSingleNode('metadata[@key="name"]')
            if ($null -ne $nameNode) { $nameNode.SetAttribute('value', $survivorNames[$idSurvivor]) }

            $fcNode = $sSurvivor.SelectSingleNode('metadata[@face_count]')
            if ($null -ne $fcNode) { $fcNode.SetAttribute('face_count', $totalFaces.ToString()) }

            foreach ($ep in @($sSurvivor.SelectNodes('part'))) { $ep.ParentNode.RemoveChild($ep) | Out-Null }

            for ($pi = 0; $pi -lt $mergedParts.Count; $pi++) {
                $p = $mergedParts[$pi]
                $compObjId = $mergedComps[$pi].GetAttribute('objectid')
                $p.SetAttribute('id', $compObjId)

                $pNameNode = $p.SelectSingleNode('metadata[@key="name"]')
                if ($null -ne $pNameNode) { $pNameNode.SetAttribute('value', "MergedPart_$compObjId") }
                $sSurvivor.AppendChild($p) | Out-Null
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
                if ($null -ne $asmSurv) {
                    $sR = $groupTxs[0]
                    $asmSurv.SetAttribute('transform', "$($sR[0]) $($sR[1]) $($sR[2]) $($sR[3]) $($sR[4]) $($sR[5]) $($sR[6]) $($sR[7]) $($sR[8]) $($txNew[9]) $($txNew[10]) $($txNew[11])")
                }
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
    if ($null -ne $loneObj) { $loneObj.SetAttribute('UUID', $nsProd, $objUuidStr) | Out-Null }
    $objUuidCounter++

    if ($hasSettings -and ($null -ne ($sLone = $settObjById[$loneId]))) {
        $loneNameNode = $sLone.SelectSingleNode('metadata[@key="name"]')
        if ($null -ne $loneNameNode) {
            $loneNameNode.SetAttribute('value', "$($loneNameNode.GetAttribute('value'))_Lone_$loneCounter")
        }
    }

    $loneComps = $loneObj.SelectNodes('m:components/m:component', $xns)
    if ($loneComps.Count -gt 0) {
        $sourceRelPath = $loneComps[0].GetAttribute('path', $nsProd)
        if ([string]::IsNullOrWhiteSpace($sourceRelPath)) { $sourceRelPath = $loneComps[0].GetAttribute('p:path') }
        if ([string]::IsNullOrWhiteSpace($sourceRelPath)) { $sourceRelPath = $loneComps[0].GetAttribute('path') }

        $sourceLocalPath = ($sourceRelPath.TrimStart('/')).Replace('/', '\')
        $sourceDiskPath = Join-Path $WorkDir $sourceLocalPath

        $newModelName = "object_${modelFileCounter}.model"
        $newModelPath = "/3D/Objects/$newModelName"
        $newDiskPath = Join-Path $objectsDir $newModelName

        if (-not $sourceToMasterMap.Contains($sourceRelPath)) {
            $normSource = [System.IO.Path]::GetFullPath($sourceDiskPath)
            $normDest = [System.IO.Path]::GetFullPath($newDiskPath)
            if ($normSource -ne $normDest) {
                Copy-Item -Path $normSource -Destination $normDest -Force
            }
            $sourceToMasterMap[$sourceRelPath] = $newModelPath
            $masterPathForThisGroup = $newModelPath
        } else {
            [System.IO.File]::WriteAllText($newDiskPath, $emptyShell, (New-Object System.Text.UTF8Encoding($false)))
            $masterPathForThisGroup = $sourceToMasterMap[$sourceRelPath]
        }
        $usedModelPaths.Add($newModelPath) | Out-Null
        $modelFileCounter++

        foreach ($c in $loneComps) {
            $c.SetAttribute('path', $nsProd, $masterPathForThisGroup) | Out-Null
        }
    }

    $loneCounter++
}

# ── Clean Killed Instances From Configs (Leaves Natural IDs Alone) ────────────
if ($hasSettings -and ($plate = $settings.SelectSingleNode('//plate'))) {
    foreach ($inst in @($plate.SelectNodes('model_instance'))) {
        $metaId = $inst.SelectSingleNode('metadata[@key="object_id"]')
        if ($null -ne $metaId -and $killedIds.Contains($metaId.GetAttribute('value'))) {
            $inst.ParentNode.RemoveChild($inst) | Out-Null
        }
    }
}

# --- TRIM KILLED IDs FROM CUT_INFORMATION.XML SAFELY ---
if (($null -ne $cutInfoPath) -and (Test-Path $cutInfoPath)) {
    [xml]$cutXml = [System.IO.File]::ReadAllText($cutInfoPath, [System.Text.Encoding]::UTF8)
    $cutModified = $false

    foreach ($co in @($cutXml.SelectNodes('//*[local-name()="object"]'))) {
        $oldId = $co.GetAttribute('id')
        if ($killedIds.Contains($oldId)) {
            $co.ParentNode.RemoveChild($co) | Out-Null
            $cutModified = $true
        }
    }
    if ($cutModified) { Save-Xml $cutXml $cutInfoPath }
}

# --- TRIM KILLED IDs AND UPDATE FACES FROM SLICE CACHE ---
if (($null -ne $sliceInfoPath) -and (Test-Path $sliceInfoPath)) {
    [xml]$sliceXml = [System.IO.File]::ReadAllText($sliceInfoPath, [System.Text.Encoding]::UTF8)
    $sliceModified = $false

    foreach ($sObj in @($sliceXml.SelectNodes('//*[local-name()="object"]'))) {
        $oldId = $sObj.GetAttribute('id')
        if ($killedIds.Contains($oldId)) {
            $sObj.ParentNode.RemoveChild($sObj) | Out-Null
            $sliceModified = $true
        } elseif ($survivorFaces.Contains($oldId)) {
            $fcNode = $sObj.SelectSingleNode('*[local-name()="metadata" and @face_count]')
            if ($null -ne $fcNode) { $fcNode.SetAttribute('face_count', $survivorFaces[$oldId].ToString()) }
            $nameNode = $sObj.SelectSingleNode('*[local-name()="metadata" and @key="name"]')
            if ($null -ne $nameNode) { $nameNode.SetAttribute('value', $survivorNames[$oldId]) }

            if ($hasSettings) {
                $matchingSettingsObj = $settings.SelectSingleNode("//*[local-name()='object' and @id='$oldId']")
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
            $sliceModified = $true
        }
    }

    foreach ($inst in @($sliceXml.SelectNodes('//*[local-name()="model_instance"]'))) {
        $metaId = $inst.SelectSingleNode('*[local-name()="metadata" and @key="object_id"]')
        if ($null -ne $metaId) {
            $oldId = $metaId.GetAttribute('value')
            if ($killedIds.Contains($oldId)) {
                $inst.ParentNode.RemoveChild($inst) | Out-Null
                $sliceModified = $true
            }
        }
    }
    if ($sliceModified) { Save-Xml $sliceXml $sliceInfoPath }
}

# ── GEOMETRY GARBAGE COLLECTION & RELS REBUILD ────────────────────────────────
# FINAL SWEEP: Protect ANY file still referenced anywhere in the XML
foreach ($node in $xml.SelectNodes('//*[@*[local-name()="path"]]', $xns)) {
    $relPath = $node.GetAttribute('path', $nsProd)
    if ([string]::IsNullOrWhiteSpace($relPath)) { $relPath = $node.GetAttribute('p:path') }
    if ([string]::IsNullOrWhiteSpace($relPath)) { $relPath = $node.GetAttribute('path') }
    if (-not [string]::IsNullOrWhiteSpace($relPath)) {
        if (-not $relPath.StartsWith('/')) { $relPath = '/' + $relPath }
        $usedModelPaths.Add($relPath) | Out-Null
    }
}

$allModelFiles = Get-ChildItem -Path $objectsDir -Filter '*.model'
foreach ($f in $allModelFiles) {
    $checkPath = "/3D/Objects/" + $f.Name
    if (-not $usedModelPaths.Contains($checkPath)) {
        Remove-Item $f.FullName -Force
    }
}

# ── Renumber build-item IDs to 1..N and write VLH 1..N ──────────────────────
# Bambu maps VLH by the model_settings object_id values.
# Working files always have consecutive IDs starting at 1.
# Renumber all surviving build items + their settings entries to 1..N,
# then write VLH for object_id=1 through object_id=N.

$n = 0
$idRemap = @{}  # old objectid -> new sequential id string

foreach ($item in $buildItems) {
    $n++
    $oldId = $item.GetAttribute('objectid')
    $newId = "$n"
    $idRemap[$oldId] = $newId
}

if ($n -gt 0) {
    # --- Remap in 3dmodel.model ---
    # Build items
    foreach ($item in $buildItems) {
        $item.SetAttribute('objectid', $idRemap[$item.GetAttribute('objectid')])
    }
    # Resource objects (top-level wrappers, same IDs as build items)
    foreach ($obj in $xml.SelectNodes('//m:resources/m:object', $xns)) {
        $oid = $obj.GetAttribute('id')
        if ($idRemap.ContainsKey($oid)) {
            $obj.SetAttribute('id', $idRemap[$oid])
        }
    }

    # --- Remap in model_settings.config ---
    if ($hasSettings) {
        # <object id="X">
        foreach ($obj in $settings.SelectNodes('//object')) {
            $oid = $obj.GetAttribute('id')
            if ($idRemap.ContainsKey($oid)) { $obj.SetAttribute('id', $idRemap[$oid]) }
        }
        # <assemble_item object_id="X">
        foreach ($ai in $settings.SelectNodes('//assemble_item')) {
            $oid = $ai.GetAttribute('object_id')
            if ($idRemap.ContainsKey($oid)) { $ai.SetAttribute('object_id', $idRemap[$oid]) }
        }
        # <metadata key="object_id" value="X"> inside model_instance
        foreach ($meta in $settings.SelectNodes("//metadata[@key='object_id']")) {
            $oid = $meta.GetAttribute('value')
            if ($idRemap.ContainsKey($oid)) { $meta.SetAttribute('value', $idRemap[$oid]) }
        }
    }

    Write-Host "  Renumbered $n build-item IDs to 1..$n"

    # --- Rebuild $objById with new IDs so VLH component lookup works ---
    $objById = @{}
    foreach ($obj in $xml.SelectNodes('//m:resources/m:object', $xns)) {
        $objById[$obj.GetAttribute('id')] = $obj
    }

    # --- Write VLH for object_id=1..N (majority-vote profile) ---
    $vlhPath = Join-Path $WorkDir "Metadata\layer_heights_profile.txt"
    if (Test-Path $vlhPath) {
        $vlhRaw = [System.IO.File]::ReadAllLines($vlhPath, [System.Text.Encoding]::UTF8)
        $dataCounts = @{}
        foreach ($line in $vlhRaw) {
            if ($line -match '\|(.+)$') {
                $ds = $Matches[1].Trim()
                if ($dataCounts.ContainsKey($ds)) { $dataCounts[$ds]++ } else { $dataCounts[$ds] = 1 }
            }
        }
        if ($dataCounts.Count -gt 0) {
            $masterVlhData = ($dataCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
            $newVlhLines = 1..$n | ForEach-Object { "object_id=$_|$masterVlhData" }
            [System.IO.File]::WriteAllText(
                $vlhPath,
                ($newVlhLines -join "`r`n"),
                (New-Object System.Text.UTF8Encoding($false))
            )
            Write-Host "  VLH sync: wrote $n entries for object_id=1..$n"
        }
    }
}

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

# ── Repack Using Original Method (NO BACKSLASHES) ─────────────────────────────
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
Get-ChildItem $WorkDir -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($WorkDir.Length).TrimStart('\','/').Replace('\','/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
}
$zip.Dispose()

$finalCount = $mergePlan.Count + $lone
$report.Add("Final object count: $finalCount")
if ($ReportPath -ne "nul" -and -not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $report | Set-Content -Path $ReportPath -Encoding UTF8
}
Write-Host "Success! Ignored: $($ignoredItems.Count). Merged: $groupCount groups. Final Target Count: $finalCount."

Start-Sleep -Milliseconds 500