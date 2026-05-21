param(
    [string]$InputPath,                                                                    # Full.3mf to read from
    [string]$OutputPath,                                                                   # BOD.3mf destination
    [string]$BambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe",
    [int]$PairCount    = 5                                                                 # How many pairs to keep
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
