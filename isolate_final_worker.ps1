param(
    [string]$WorkDir,
    [string]$OutputPath
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  isolate_final_worker.ps1 - ISOLATE CENTER OBJECT FOR FINAL.3MF (V2)
# ════════════════════════════════════════════════════════════════════════════════

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

function Parse-Tx([string]$s) { if ([string]::IsNullOrWhiteSpace($s)) { return [double[]](1,0,0, 0,1,0, 0,0,1, 0,0,0) }; return [double[]]($s.Trim() -split '\s+') }
function Save-Xml([xml]$doc, [string]$path) {
    $settings = New-Object System.Xml.XmlWriterSettings; $settings.Encoding = New-Object System.Text.UTF8Encoding($false); $settings.Indent = $true
    $w = [System.Xml.XmlWriter]::Create($path, $settings); $doc.Save($w); $w.Close()
}
function Find-File([string]$base, [string]$rel) {
    $p = Join-Path $base $rel; if (Test-Path $p) { return $p }; $p2 = Join-Path $base ($rel -replace '\\','/'); if (Test-Path $p2) { return $p2 }; return $null
}

$modelFile = (Get-ChildItem -Path $WorkDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
$objectsDir = Join-Path $WorkDir '3D/Objects'
$relsPath = Find-File $WorkDir '3D/_rels/3dmodel.model.rels'
$settingsPath = Find-File $WorkDir 'Metadata/model_settings.config'
$cutInfoPath = Find-File $WorkDir 'Metadata/cut_information.xml'

[xml]$xml = [System.IO.File]::ReadAllText($modelFile, [System.Text.Encoding]::UTF8)
$xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

$buildItems = @($xml.SelectNodes('//m:build/m:item', $xns))
$objById = @{}
foreach ($o in $xml.SelectNodes('//m:resources/m:object', $xns)) { $objById[$o.GetAttribute('id')] = $o }

$hasSettings = ($null -ne $settingsPath) -and (Test-Path $settingsPath)
[xml]$settings = $null; $settObjById = @{}
if ($hasSettings) {
    $settings = [xml][System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)
    foreach ($node in $settings.config.ChildNodes) { if ($node.LocalName -eq 'object') { $settObjById[$node.GetAttribute('id')] = $node } }
}

# ── 1. Find Majority Geometry (Ignore text/outliers) ──
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

$targetItems = @()
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
    if ($isTarget) { $targetItems += $item }
}

if ($targetItems.Count -eq 0) { exit 1 }

# ── 2. Find Closest Target Object to Physical Center (128, 128) ──
$closestItem = $null
$minDist = [double]::MaxValue
foreach ($item in $targetItems) {
    $tx = Parse-Tx ($item.GetAttribute('transform'))
    $dist = [math]::Pow($tx[9] - 128, 2) + [math]::Pow($tx[10] - 128, 2)
    if ($dist -lt $minDist) {
        $minDist = $dist
        $closestItem = $item
    }
}

# ── 3. Eradicate Everything Else on the Plate ──
$survivorObjectId = $closestItem.GetAttribute('objectid')
$killedIds = New-Object System.Collections.Generic.HashSet[string]

foreach ($item in $buildItems) {
    $id = $item.GetAttribute('objectid')
    if ($item -ne $closestItem) {
        $killedIds.Add($id) | Out-Null
        
        # Remove from build plate
        $item.ParentNode.RemoveChild($item) | Out-Null
        
        # Remove core object definition
        $obj = $objById[$id]
        if ($null -ne $obj) { $obj.ParentNode.RemoveChild($obj) | Out-Null }
    }
}

# ── 4. Target-Locked Metadata Eradication & Shifting ──
if ($hasSettings) {
    # Clean up <plate> instances
    if ($plate = $settings.SelectSingleNode('//plate')) {
        $foundSurvivor = $false
        foreach ($inst in @($plate.SelectNodes('model_instance'))) {
            $metaId = $inst.SelectSingleNode('metadata[@key="object_id"]')
            if ($null -ne $metaId) {
                if ($metaId.GetAttribute('value') -eq $survivorObjectId -and -not $foundSurvivor) {
                    $foundSurvivor = $true
                } else {
                    $inst.ParentNode.RemoveChild($inst) | Out-Null
                }
            } else {
                $inst.ParentNode.RemoveChild($inst) | Out-Null
            }
        }
    }
    
    # Remove orphaned <object> metadata blocks entirely
    foreach ($kId in $killedIds) {
        $sMember = $settObjById[$kId]
        if ($null -ne $sMember -and $null -ne $sMember.ParentNode) {
            $sMember.ParentNode.RemoveChild($sMember) | Out-Null
        }
    }

    # Clean up <assemble> trackers and handle transform shifts
    if ($assemble = $settings.SelectSingleNode('//assemble')) {
        foreach ($kId in $killedIds) {
            $asmMember = $assemble.SelectSingleNode("assemble_item[@object_id='$kId']")
            if ($null -ne $asmMember) { $assemble.RemoveChild($asmMember) | Out-Null }
        }
        # Force the survivor's assemble transform to perfectly match its actual plate coordinates
        $asmSurv = $assemble.SelectSingleNode("assemble_item[@object_id='$survivorObjectId']")
        if ($null -ne $asmSurv) {
            $tx = Parse-Tx ($closestItem.GetAttribute('transform'))
            $asmSurv.SetAttribute('transform', "1 0 0 0 1 0 0 0 1 $($tx[9]) $($tx[10]) $($tx[11])")
        }
    }
}

# Clean cut definitions
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

# ── 5. Clean Unused .model files & rebuild .rels ──
$preservedModelPaths = New-Object System.Collections.Generic.HashSet[string]
$survivorObj = $objById[$survivorObjectId]
if ($null -ne $survivorObj) {
    foreach ($c in $survivorObj.SelectNodes('m:components/m:component', $xns)) {
        $path = $c.GetAttribute('path', $nsProd)
        if ([string]::IsNullOrEmpty($path)) { $path = $c.GetAttribute('p:path') }
        if (-not [string]::IsNullOrEmpty($path)) { $preservedModelPaths.Add($path) | Out-Null }
    }
}

# Delete physical files that no longer exist on the plate
$allModelFiles = Get-ChildItem -Path $objectsDir -Filter '*.model'
foreach ($f in $allModelFiles) { 
    $checkPath = "/3D/Objects/" + $f.Name
    if (-not $preservedModelPaths.Contains($checkPath)) {
        Remove-Item $f.FullName -Force 
    }
}

# Rebuild pointers to prevent loading errors
if ($null -ne $relsPath) {
    $relsXml = "<?xml version='1.0' encoding='UTF-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
    $relIdx = 1
    foreach ($path in $preservedModelPaths) {
        $relsXml += "<Relationship Target='$path' Id='rel-ign-$relIdx' Type='http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel'/>"
        $relIdx++
    }
    $relsXml += "</Relationships>"
    [System.IO.File]::WriteAllText($relsPath, $relsXml, (New-Object System.Text.UTF8Encoding($false)))
}

Save-Xml $xml $modelFile
if ($hasSettings) { Save-Xml $settings $settingsPath }

# ── 6. Repack ──
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Create')
Get-ChildItem $WorkDir -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($WorkDir.Length).TrimStart('\','/').Replace('\','/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel)
}
$zip.Dispose()