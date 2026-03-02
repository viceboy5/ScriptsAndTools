param(
    [string]$WorkDir,
    [string]$InputPath,
    [string]$OutputPath,
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  merge_3mf_worker.ps1 - AUTO-PLAN N-WAY MERGE (WITH GEOMETRY PRESERVATION)
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

# ── Setup New Geometry File ───────────────────────────────────────────────────
$maxExisting = ($allModelFiles | ForEach-Object { [int]($_.BaseName -replace 'object_','') } | Measure-Object -Maximum).Maximum
$newModelNum  = $maxExisting + 57
$newModelName = "object_$newModelNum.model"
$newModelPath = "/3D/Objects/$newModelName"
$newPrefix    = ([int]($newModelNum -band 0xFFFF)).ToString('x4')

$geoContent = [System.IO.File]::ReadAllText($allModelFiles[0].FullName, [System.Text.Encoding]::UTF8)
if ($geoContent -match 'p:UUID="([0-9a-fA-F]{4})') { $geoContent = $geoContent -replace $Matches[1], $newPrefix }
[System.IO.File]::WriteAllText((Join-Path $objectsDir $newModelName), $geoContent, (New-Object System.Text.UTF8Encoding($false)))

$report = New-Object System.Collections.Generic.List[string]

# ── Outlier Detection (Isolate Version Text) ──────────────────────────────────
$mergeItems   = @()
$ignoredItems = @()
$fcMap = @{}

foreach ($item in $buildItems) {
    $id = $item.GetAttribute('objectid')
    $fc = "unknown"
    if ($hasSettings -and $null -ne $settObjById[$id]) {
        $fcNode = $settObjById[$id].SelectSingleNode('metadata[@key="face_count"]')
        if ($null -ne $fcNode) { $fc = $fcNode.GetAttribute('value') }
    }
    if (-not $fcMap.Contains($fc)) { $fcMap[$fc] = 0 }
    $fcMap[$fc]++
}
$majorityFc = ($fcMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Name

foreach ($item in $buildItems) {
    $id = $item.GetAttribute('objectid')
    $isTarget = $true
    
    if ($hasSettings -and $null -ne $settObjById[$id]) {
        $fcNode = $settObjById[$id].SelectSingleNode('metadata[@key="face_count"]')
        $fc = if ($null -ne $fcNode) { $fcNode.GetAttribute('value') } else { "unknown" }
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

# ── Dynamic Merge Loop ────────────────────────────────────────────────────────
$cursor = 0; $groupCount = 0
$killedIds = New-Object System.Collections.Generic.HashSet[string]

foreach ($groupSize in $mergePlan) {
    $groupItems = @(); $groupObjs = @(); $groupTxs = @()

    for ($k = 0; $k -lt $groupSize; $k++) {
        $item = $mergeItems[$cursor + $k]
        $id   = $item.GetAttribute('objectid')
        $groupItems += $item; $groupObjs += $objById[$id]; $groupTxs += ,( Parse-Tx ($item.GetAttribute('transform')) )
    }
    $cursor += $groupSize
    $idSurvivor = $groupItems[0].GetAttribute('objectid')

    # Centroid
    [double]$sumX = 0; [double]$sumY = 0; [double]$sumZ = 0
    foreach ($tx in $groupTxs) { $sumX += $tx[9]; $sumY += $tx[10]; $sumZ += $tx[11] }
    [double]$cX = $sumX / $groupSize; [double]$cY = $sumY / $groupSize; [double]$cZ = $sumZ / $groupSize
    [double[]]$txNew = (1,0,0, 0,1,0, 0,0,1, $cX, $cY, $cZ)
    $txNewStr = Fmt-Tx $txNew
    $groupItems[0].SetAttribute('transform', $txNewStr)
    [double[]]$txNew = Parse-Tx $txNewStr
    [double[]]$invTxNew = Inv-Tx $txNew

    # Components
    $newCompsEl = $xml.CreateElement('components', $nsCore)
    foreach ($phase in @('normal', 'special')) {
        for ($k = 0; $k -lt $groupSize; $k++) {
            $obj = $groupObjs[$k]; [double[]]$origTx = $groupTxs[$k]
            foreach ($c in $obj.SelectNodes('m:components/m:component', $xns)) {
                $isSpecial = ($c.GetAttribute('objectid') -eq '40')
                if ($phase -eq 'normal' -and $isSpecial) { continue }
                if ($phase -eq 'special' -and -not $isSpecial) { continue }
                [double[]]$compTx  = Parse-Tx ($c.GetAttribute('transform'))
                [double[]]$bakedTx = Mul-Tx $invTxNew (Mul-Tx $origTx $compTx)
                $newComp = $xml.CreateElement('component', $nsCore)
                $newComp.SetAttribute('path', $nsProd, $newModelPath)
                $newComp.SetAttribute('objectid', $c.GetAttribute('objectid'))
                $newComp.SetAttribute('UUID', $nsProd, [guid]::NewGuid().ToString())
                $newComp.SetAttribute('transform', (Fmt-Tx $bakedTx))
                $newCompsEl.AppendChild($newComp) | Out-Null
            }
        }
    }
    if ($null -ne ($oldC = $groupObjs[0].SelectSingleNode('m:components', $xns))) { $groupObjs[0].ReplaceChild($newCompsEl, $oldC) | Out-Null } else { $groupObjs[0].AppendChild($newCompsEl) | Out-Null }
    $groupObjs[0].SetAttribute('UUID', $nsProd, (([int]$newModelNum).ToString('x8') + '-71cb-4c03-9d28-80fed5dfa1dc'))

    # Metadata
    if ($hasSettings) {
        $sSurvivor = $settObjById[$idSurvivor]
        if ($null -ne $sSurvivor) {
            $nameNode = $sSurvivor.SelectSingleNode('metadata[@key="name"]')
            $baseName = if ($null -ne $nameNode) { $nameNode.GetAttribute('value') } else { 'Object' }
            if ($null -ne $nameNode) { $nameNode.SetAttribute('value', "$baseName$groupSize") }

            [int]$totalFaces = 0
            $fcNode = $sSurvivor.SelectSingleNode('metadata[@key="face_count"]')
            if ($null -ne $fcNode) { [int]::TryParse($fcNode.GetAttribute('value'), [ref]$totalFaces) | Out-Null }

            for ($k = 1; $k -lt $groupSize; $k++) {
                $memberId = $groupItems[$k].GetAttribute('objectid')
                $killedIds.Add($memberId) | Out-Null 
                
                $sMember  = $settObjById[$memberId]
                if ($null -eq $sMember) { continue }
                $mfcNode = $sMember.SelectSingleNode('metadata[@key="face_count"]')
                if ($null -ne $mfcNode) { [int]$mfc = 0; [int]::TryParse($mfcNode.GetAttribute('value'), [ref]$mfc) | Out-Null; $totalFaces += $mfc }
                foreach ($p in @($sMember.SelectNodes('part'))) { $sSurvivor.AppendChild($settings.ImportNode($p, $true)) | Out-Null }
                $sMember.ParentNode.RemoveChild($sMember) | Out-Null
            }
            if ($null -ne $fcNode) { $fcNode.SetAttribute('value', $totalFaces.ToString()) }

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

    # Remove non-survivor build items
    for ($k = 1; $k -lt $groupSize; $k++) {
        $groupItems[$k].ParentNode.RemoveChild($groupItems[$k]) | Out-Null
        $groupObjs[$k].ParentNode.RemoveChild($groupObjs[$k])   | Out-Null
    }
    $groupCount++
}

# ── Finalize Metadata (Target-Locked Eradication) ─────────────────────────────
if ($hasSettings -and ($plate = $settings.SelectSingleNode('//plate'))) {
    foreach ($inst in @($plate.SelectNodes('model_instance'))) {
        $metaId = $inst.SelectSingleNode('metadata[@key="object_id"]')
        if ($null -ne $metaId -and $killedIds.Contains($metaId.GetAttribute('value'))) {
            $inst.ParentNode.RemoveChild($inst) | Out-Null
        }
    }
}

if (($null -ne $cutInfoPath) -and (Test-Path $cutInfoPath)) {
    [xml]$cutXml = [System.IO.File]::ReadAllText($cutInfoPath, [System.Text.Encoding]::UTF8)
    $cutModified = $false
    foreach ($cutObj in @($cutXml.SelectNodes('//*[local-name()="object"]'))) {
        if ($killedIds.Contains($cutObj.GetAttribute('id'))) {
            $cutObj.ParentNode.RemoveChild($cutObj) | Out-Null; $cutModified = $true
        }
    }
    if ($cutModified) { 
        $ws = New-Object System.Xml.XmlWriterSettings; $ws.Encoding = New-Object System.Text.UTF8Encoding($false); $ws.Indent = $true
        $w = [System.Xml.XmlWriter]::Create($cutInfoPath, $ws); $cutXml.Save($w); $w.Close()
    }
}

# ── Retarget lone items' components ───────────────────────────────────────────
for ($li = ($mergeItems.Count - $lone); $li -lt $mergeItems.Count; $li++) {
    $loneId  = $mergeItems[$li].GetAttribute('objectid'); $loneObj = $objById[$loneId]
    if ($null -eq $loneObj) { continue }
    foreach ($c in $loneObj.SelectNodes('m:components/m:component', $xns)) { $c.SetAttribute('path', $nsProd, $newModelPath) }
    if ($hasSettings -and ($null -ne ($sLone = $settObjById[$loneId]))) {
        $loneNameNode = $sLone.SelectSingleNode('metadata[@key="name"]')
        if ($null -ne $loneNameNode) { $loneNameNode.SetAttribute('value', "$($loneNameNode.GetAttribute('value'))1") }
    }
}

# ── GEOMETRY PRESERVATION SYSTEM (The Namespace Fix) ──────────────────────────
$preservedModelPaths = New-Object System.Collections.Generic.HashSet[string]
foreach ($item in $ignoredItems) {
    $ignId = $item.GetAttribute('objectid')
    $ignObj = $objById[$ignId]
    if ($null -ne $ignObj) {
        foreach ($c in $ignObj.SelectNodes('m:components/m:component', $xns)) {
            $path = $c.GetAttribute('path', $nsProd)
            if ([string]::IsNullOrEmpty($path)) { $path = $c.GetAttribute('p:path') }
            if (-not [string]::IsNullOrEmpty($path)) { $preservedModelPaths.Add($path) | Out-Null }
        }
    }
}

Save-Xml $xml $modelFile
if ($hasSettings) { Save-Xml $settings $settingsPath }

foreach ($f in $allModelFiles) { 
    $checkPath = "/3D/Objects/" + $f.Name
    if ($f.Name -ne $newModelName -and -not $preservedModelPaths.Contains($checkPath)) {
        Remove-Item $f.FullName -Force 
    }
}

$relsXml = "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
$relsXml += "<Relationship Target='$newModelPath' Id='rel-main' Type='http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel'/>"
$relIdx = 1
foreach ($path in $preservedModelPaths) {
    $relsXml += "<Relationship Target='$path' Id='rel-ign-$relIdx' Type='http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel'/>"
    $relIdx++
}
$relsXml += "</Relationships>"
[System.IO.File]::WriteAllText($relsPath, $relsXml, (New-Object System.Text.UTF8Encoding($false)))

# ── Repack ────────────────────────────────────────────────────────────────────
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
Get-ChildItem $WorkDir -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($WorkDir.Length).TrimStart('\','/').Replace('\','/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel)
}
$zip.Dispose()

# Update Final Count to exclusively track target objects
$finalCount = $mergePlan.Count + $lone
$report.Add("Final object count: $finalCount")
$report | Set-Content -Path $ReportPath -Encoding UTF8
Write-Host "Success! Ignored: $($ignoredItems.Count). Merged: $groupCount groups. Final Target Count: $finalCount."