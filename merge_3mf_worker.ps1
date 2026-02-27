param(
    [string]$WorkDir,
    [string]$InputPath,
    [string]$OutputPath,
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  merge_3mf_worker.ps1 - DYNAMIC MULTI-OBJECT VERSION
# ════════════════════════════════════════════════════════════════════════════════

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

# ── Transform math ────────────────────────────────────────────────────────────
function Parse-Tx([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) {
        return [double[]](1,0,0, 0,1,0, 0,0,1, 0,0,0)
    }
    return [double[]]($s.Trim() -split '\s+')
}

function Fmt-Tx([double[]]$v) {
    return ($v | ForEach-Object { $_.ToString('G15') }) -join ' '
}

function Mul-Tx([double[]]$A, [double[]]$B) {
    $a00=$A[0]; $a01=$A[3]; $a02=$A[6]; $atx=$A[9]
    $a10=$A[1]; $a11=$A[4]; $a12=$A[7]; $aty=$A[10]
    $a20=$A[2]; $a21=$A[5]; $a22=$A[8]; $atz=$A[11]

    $b00=$B[0]; $b01=$B[3]; $b02=$B[6]; $btx=$B[9]
    $b10=$B[1]; $b11=$B[4]; $b12=$B[7]; $bty=$B[10]
    $b20=$B[2]; $b21=$B[5]; $b22=$B[8]; $btz=$B[11]

    # Explicitly calculate each position to avoid op_Multiply/Object[] errors
    $c00 = ($a00 * $b00) + ($a01 * $b10) + ($a02 * $b20)
    $c10 = ($a10 * $b00) + ($a11 * $b10) + ($a12 * $b20)
    $c20 = ($a20 * $b00) + ($a21 * $b10) + ($a22 * $b20)

    $c01 = ($a00 * $b01) + ($a01 * $b11) + ($a02 * $b21)
    $c11 = ($a10 * $b01) + ($a11 * $b11) + ($a12 * $b21)
    $c21 = ($a20 * $b01) + ($a21 * $b11) + ($a22 * $b21)

    $c02 = ($a00 * $b02) + ($a01 * $b12) + ($a02 * $b22)
    $c12 = ($a10 * $b02) + ($a11 * $b12) + ($a12 * $b22)
    $c22 = ($a20 * $b02) + ($a21 * $b12) + ($a22 * $b22)

    $ctx = ($a00 * $btx) + ($a01 * $bty) + ($a02 * $btz) + $atx
    $cty = ($a10 * $btx) + ($a11 * $bty) + ($a12 * $btz) + $aty
    $ctz = ($a20 * $btx) + ($a21 * $bty) + ($a22 * $btz) + $atz

    $result = [double[]] @($c00, $c10, $c20, $c01, $c11, $c21, $c02, $c12, $c22, $ctx, $cty, $ctz)
    return $result
}

function Inv-Tx([double[]]$A) {
    $ir00=$A[0]; $ir01=$A[1]; $ir02=$A[2]; $ir10=$A[3]; $ir11=$A[4]; $ir12=$A[5]; $ir20=$A[6]; $ir21=$A[7]; $ir22=$A[8]
    $tx=$A[9]; $ty=$A[10]; $tz=$A[11]
    return [double[]]($ir00,$ir10,$ir20, $ir01,$ir11,$ir21, $ir02,$ir12,$ir22, -($ir00*$tx + $ir01*$ty + $ir02*$tz), -($ir10*$tx + $ir11*$ty + $ir12*$tz), -($ir20*$tx + $ir21*$ty + $ir22*$tz))
}

function Save-Xml([xml]$doc, [string]$path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $true
    $w = [System.Xml.XmlWriter]::Create($path, $settings)
    $doc.Save($w)
    $w.Close()
}

function Find-File([string]$base, [string]$rel) {
    $p = Join-Path $base $rel; if (Test-Path $p) { return $p }
    $p2 = Join-Path $base ($rel -replace '\\','/'); if (Test-Path $p2) { return $p2 }
    return $null
}

# ── Locate Files ──────────────────────────────────────────────────────────────
$modelFile = (Get-ChildItem -Path $WorkDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
$objectsDir = Join-Path $WorkDir '3D/Objects'
$relsPath = Find-File $WorkDir '3D/_rels/3dmodel.model.rels'
$settingsPath = Find-File $WorkDir 'Metadata/model_settings.config'
$cutInfoPath = Find-File $WorkDir 'Metadata/cut_information.xml'
$layerProfPath = Find-File $WorkDir 'Metadata/layer_heights_profile.txt'

$allModelFiles = @(Get-ChildItem -Path $objectsDir -Filter '*.model' | Sort-Object { [int]($_.BaseName -replace 'object_','') })

# ── Parse Documents ───────────────────────────────────────────────────────────
[xml]$xml = [System.IO.File]::ReadAllText($modelFile, [System.Text.Encoding]::UTF8)
$xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

$buildItems = @($xml.SelectNodes('//m:build/m:item', $xns))
$objById = @{}; foreach ($o in $xml.SelectNodes('//m:resources/m:object', $xns)) { $objById[$o.GetAttribute('id')] = $o }

$hasSettings = ($null -ne $settingsPath) -and (Test-Path $settingsPath)
[xml]$settings = $null; $settObjById = @{}
if ($hasSettings) {
    $settings = [xml][System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)
    foreach ($node in $settings.config.ChildNodes) { if ($node.LocalName -eq 'object') { $settObjById[$node.GetAttribute('id')] = $node } }
}

$report = New-Object System.Collections.Generic.List[string]

# ── Setup New Geometry ────────────────────────────────────────────────────────
$maxExisting = ($allModelFiles | ForEach-Object { [int]($_.BaseName -replace 'object_','') } | Measure-Object -Maximum).Maximum
$newModelNum = $maxExisting + 57
$newModelName = "object_$newModelNum.model"
$newModelPath = "/3D/Objects/$newModelName"
$newPrefix = ([int]($newModelNum -band 0xFFFF)).ToString('x4')

$geoContent = [System.IO.File]::ReadAllText($allModelFiles[0].FullName, [System.Text.Encoding]::UTF8)
if ($geoContent -match 'p:UUID="([0-9a-fA-F]{4})') { $geoContent = $geoContent -replace $Matches[1], $newPrefix }
[System.IO.File]::WriteAllText((Join-Path $objectsDir $newModelName), $geoContent, (New-Object System.Text.UTF8Encoding($false)))

# ── Dynamic Merge Loop ────────────────────────────────────────────────────────
for ($i = 0; $i -lt ($buildItems.Count - 1); $i += 2) {
    $itemA = $buildItems[$i]; $itemB = $buildItems[$i+1]
    $idA = $itemA.GetAttribute('objectid'); $idB = $itemB.GetAttribute('objectid')
    $objA = $objById[$idA]; $objB = $objById[$idB]

    [double[]]$txA = Parse-Tx ($itemA.GetAttribute('transform'))
    [double[]]$txB = Parse-Tx ($itemB.GetAttribute('transform'))

    $midX = ($txA[9] + $txB[9])/2.0; $midY = ($txA[10] + $txB[10])/2.0; $midZ = ($txA[11] + $txB[11])/2.0
    [double[]]$txNew = (1,0,0, 0,1,0, 0,0,1, $midX, $midY, $midZ)
    $invTxNew = Inv-Tx $txNew
    $itemA.SetAttribute('transform', (Fmt-Tx $txNew))

    # Merge Components
    $newCompsEl = $xml.CreateElement('components', $nsCore)
    
    # Process ObjA components
    foreach ($c in $objA.SelectNodes('m:components/m:component', $xns)) {
        [double[]]$compTx = Parse-Tx ($c.GetAttribute('transform'))
        [double[]]$bakedTx = Mul-Tx $invTxNew (Mul-Tx $txA $compTx)
        
        $newComp = $xml.CreateElement('component', $nsCore)
        $newComp.SetAttribute('path', $nsProd, $newModelPath)
        $newComp.SetAttribute('objectid', $c.GetAttribute('objectid'))
        $newComp.SetAttribute('UUID', $nsProd, [guid]::NewGuid().ToString())
        $newComp.SetAttribute('transform', (Fmt-Tx $bakedTx))
        $newCompsEl.AppendChild($newComp) | Out-Null
    }

    # Process ObjB components
    foreach ($c in $objB.SelectNodes('m:components/m:component', $xns)) {
        [double[]]$compTx = Parse-Tx ($c.GetAttribute('transform'))
        [double[]]$bakedTx = Mul-Tx $invTxNew (Mul-Tx $txB $compTx)
        
        $newComp = $xml.CreateElement('component', $nsCore)
        $newComp.SetAttribute('path', $nsProd, $newModelPath)
        $newComp.SetAttribute('objectid', $c.GetAttribute('objectid'))
        $newComp.SetAttribute('UUID', $nsProd, [guid]::NewGuid().ToString())
        $newComp.SetAttribute('transform', (Fmt-Tx $bakedTx))
        $newCompsEl.AppendChild($newComp) | Out-Null
    }
    
    if ($null -ne ($oldC = $objA.SelectSingleNode('m:components', $xns))) { $objA.ReplaceChild($newCompsEl, $oldC) } else { $objA.AppendChild($newCompsEl) }
    $objA.SetAttribute('UUID', $nsProd, (([int]$newModelNum).ToString('x8') + '-71cb-4c03-9d28-80fed5dfa1dc'))

    # Update model_settings.config
    if ($hasSettings -and $settObjById[$idA] -and $settObjById[$idB]) {
        $sA = $settObjById[$idA]; $sB = $settObjById[$idB]
        if ($n = $sA.SelectSingleNode('metadata[@key="name"]')) { $n.SetAttribute('value', 'Assembly') }
        foreach ($p in @($sB.SelectNodes('part'))) { $sA.AppendChild($settings.ImportNode($p, $true)) | Out-Null }
        $sB.ParentNode.RemoveChild($sB) | Out-Null
        
        $assemble = $settings.SelectSingleNode('//assemble')
        if ($asmA = $assemble.SelectSingleNode("assemble_item[@object_id='$idA']")) { $asmA.SetAttribute('transform', "1 0 0 0 1 0 0 0 1 $($txNew[9]) $($txNew[10]) $($txNew[11])") }
        if ($asmB = $assemble.SelectSingleNode("assemble_item[@object_id='$idB']")) { $assemble.RemoveChild($asmB) | Out-Null }
    }

    $itemB.ParentNode.RemoveChild($itemB) | Out-Null
    $objB.ParentNode.RemoveChild($objB) | Out-Null
}

# ── Finalize ──────────────────────────────────────────────────────────────────
if ($hasSettings -and ($plate = $settings.SelectSingleNode('//plate'))) {
    $inst = @($plate.SelectNodes('model_instance')); for ($j=1; $j -lt $inst.Count; $j++) { $plate.RemoveChild($inst[$j]) | Out-Null }
}
Save-Xml $xml $modelFile
if ($hasSettings) { Save-Xml $settings $settingsPath }
foreach ($f in $allModelFiles) { Remove-Item $f.FullName -Force }
[System.IO.File]::WriteAllText($relsPath, "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Target='/3D/Objects/$newModelName' Id='rel-1' Type='http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel'/></Relationships>")

# Repack
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
Get-ChildItem $WorkDir -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($WorkDir.Length).TrimStart('\','/').Replace('\','/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel)
}
$zip.Dispose()
Write-Host "Success: Managed dynamic merge of $($buildItems.Count) objects."