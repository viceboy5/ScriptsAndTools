param(
    [string]$WorkDir,
    [string]$InputPath,
    [string]$OutputPath,
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  merge_3mf_worker.ps1 - DYNAMIC N-WAY MULTI-OBJECT VERSION
#
#  MASTER_MERGE_COUNT controls how many consecutive build items are collapsed
#  into a single Assembly object per group.
#
#  Examples:
#    MASTER_MERGE_COUNT = 2  → pairs   (original behaviour)
#    MASTER_MERGE_COUNT = 3  → triples
#    MASTER_MERGE_COUNT = 4  → quads
#
#  If the file contains 6 build items and MASTER_MERGE_COUNT = 3:
#    Group 1: items 0,1,2 → Assembly 1
#    Group 2: items 3,4,5 → Assembly 2
#
#  Any leftover items at the end (if Count % MASTER_MERGE_COUNT != 0) are
#  left untouched.
#
#  Per-group calculations that scale with N:
#    - New build transform  = centroid of all N original build translations
#    - Component baking     = inv(centroid_tx) · orig_build_tx[i] · orig_comp_tx
#    - Component ordering   = [obj0 normal], [obj1 normal], ...,
#                             [obj0 special], [obj1 special], ...
#    - face_count           = sum of all N face_counts
#    - model_instances      = keep first, delete rest
#    - assemble_items       = update first to centroid, delete rest
# ════════════════════════════════════════════════════════════════════════════════

# ── MASTER CONTROL VARIABLE ───────────────────────────────────────────────────
$MASTER_MERGE_COUNT = 3   # <-- Change this to merge N objects at a time
# ─────────────────────────────────────────────────────────────────────────────

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

# ── Dynamic N-Way Merge Loop ──────────────────────────────────────────────────
# Process build items in groups of $MASTER_MERGE_COUNT.
# Each group collapses into one Assembly, surviving on buildItems[groupStart].

$groupCount = 0
for ($groupStart = 0; $groupStart + $MASTER_MERGE_COUNT -le $buildItems.Count; $groupStart += $MASTER_MERGE_COUNT) {

    # Collect the K items and objects for this group
    $groupItems = @()
    $groupObjs  = @()
    $groupTxs   = @()   # build-plate transforms for each member

    for ($k = 0; $k -lt $MASTER_MERGE_COUNT; $k++) {
        $item = $buildItems[$groupStart + $k]
        $id   = $item.GetAttribute('objectid')
        $groupItems += $item
        $groupObjs  += $objById[$id]
        $groupTxs   += ,( Parse-Tx ($item.GetAttribute('transform')) )
    }

    $idSurvivor = $groupItems[0].GetAttribute('objectid')

    # ── Centroid of all N build translations ──────────────────────────────────
    [double]$sumX = 0; [double]$sumY = 0; [double]$sumZ = 0
    foreach ($tx in $groupTxs) { $sumX += $tx[9]; $sumY += $tx[10]; $sumZ += $tx[11] }
    [double]$cX = $sumX / $MASTER_MERGE_COUNT
    [double]$cY = $sumY / $MASTER_MERGE_COUNT
    [double]$cZ = $sumZ / $MASTER_MERGE_COUNT

    # Write centroid to build item and read back exact stored value
    [double[]]$txNew = (1,0,0, 0,1,0, 0,0,1, $cX, $cY, $cZ)
    $txNewStr = Fmt-Tx $txNew
    $groupItems[0].SetAttribute('transform', $txNewStr)
    [double[]]$txNew    = Parse-Tx $txNewStr   # exact stored value
    [double[]]$invTxNew = Inv-Tx $txNew

    # ── Build merged <components> in correct N-way order ──────────────────────
    # Ordering: [normal comps for each obj in order], [special comps for each obj in order]
    # "special" = objectid == '40' (the negative/support cube, always floats last)
    $newCompsEl = $xml.CreateElement('components', $nsCore)

    foreach ($phase in @('normal', 'special')) {
        for ($k = 0; $k -lt $MASTER_MERGE_COUNT; $k++) {
            $obj  = $groupObjs[$k]
            [double[]]$origBuildTx = $groupTxs[$k]

            foreach ($c in $obj.SelectNodes('m:components/m:component', $xns)) {
                $isSpecial = ($c.GetAttribute('objectid') -eq '40')
                if (($phase -eq 'normal'  -and $isSpecial) -or
                    ($phase -eq 'special' -and -not $isSpecial)) { continue }

                [double[]]$compTx  = Parse-Tx ($c.GetAttribute('transform'))
                [double[]]$bakedTx = Mul-Tx $invTxNew (Mul-Tx $origBuildTx $compTx)

                $newComp = $xml.CreateElement('component', $nsCore)
                $newComp.SetAttribute('path',      $nsProd, $newModelPath)
                $newComp.SetAttribute('objectid',  $c.GetAttribute('objectid'))
                $newComp.SetAttribute('UUID',      $nsProd, [guid]::NewGuid().ToString())
                $newComp.SetAttribute('transform', (Fmt-Tx $bakedTx))
                $newCompsEl.AppendChild($newComp) | Out-Null
            }
        }
    }

    # Replace survivor's components
    if ($null -ne ($oldC = $groupObjs[0].SelectSingleNode('m:components', $xns))) {
        $groupObjs[0].ReplaceChild($newCompsEl, $oldC) | Out-Null
    } else {
        $groupObjs[0].AppendChild($newCompsEl) | Out-Null
    }

    # Update survivor UUID to reflect new model file number
    $groupObjs[0].SetAttribute('UUID', $nsProd, (([int]$newModelNum).ToString('x8') + '-71cb-4c03-9d28-80fed5dfa1dc'))

    # ── model_settings.config ─────────────────────────────────────────────────
    if ($hasSettings) {
        $sSurvivor = $settObjById[$idSurvivor]

        if ($null -ne $sSurvivor) {
            # Rename Assembly
            if ($n = $sSurvivor.SelectSingleNode('metadata[@key="name"]')) { $n.SetAttribute('value', 'Assembly') }

            # Sum face_counts and append parts from all non-survivor members
            [int]$totalFaces = 0
            $fcNode = $sSurvivor.SelectSingleNode('metadata[@key="face_count"]')
            if ($null -ne $fcNode) { [int]::TryParse($fcNode.GetAttribute('value'), [ref]$totalFaces) | Out-Null }

            for ($k = 1; $k -lt $MASTER_MERGE_COUNT; $k++) {
                $memberId  = $groupItems[$k].GetAttribute('objectid')
                $sMember   = $settObjById[$memberId]
                if ($null -eq $sMember) { continue }

                # Add face_count
                $memberFcNode = $sMember.SelectSingleNode('metadata[@key="face_count"]')
                if ($null -ne $memberFcNode) {
                    [int]$mfc = 0
                    [int]::TryParse($memberFcNode.GetAttribute('value'), [ref]$mfc) | Out-Null
                    $totalFaces += $mfc
                }

                # Append parts
                foreach ($p in @($sMember.SelectNodes('part'))) {
                    $sSurvivor.AppendChild($settings.ImportNode($p, $true)) | Out-Null
                }

                # Remove member settings block
                $sMember.ParentNode.RemoveChild($sMember) | Out-Null
            }

            if ($null -ne $fcNode) { $fcNode.SetAttribute('value', $totalFaces.ToString()) }

            # assemble_items: update survivor, delete members
            $assemble = $settings.SelectSingleNode('//assemble')
            if ($null -ne $assemble) {
                $asmSurvivor = $assemble.SelectSingleNode("assemble_item[@object_id='$idSurvivor']")
                if ($null -ne $asmSurvivor) {
                    $asmSurvivor.SetAttribute('transform', "1 0 0 0 1 0 0 0 1 $($txNew[9]) $($txNew[10]) $($txNew[11])")
                }
                for ($k = 1; $k -lt $MASTER_MERGE_COUNT; $k++) {
                    $memberId = $groupItems[$k].GetAttribute('objectid')
                    $asmMember = $assemble.SelectSingleNode("assemble_item[@object_id='$memberId']")
                    if ($null -ne $asmMember) { $assemble.RemoveChild($asmMember) | Out-Null }
                }
            }
        }
    }

    # ── Remove non-survivor build items and resource objects ──────────────────
    for ($k = 1; $k -lt $MASTER_MERGE_COUNT; $k++) {
        $groupItems[$k].ParentNode.RemoveChild($groupItems[$k]) | Out-Null
        $groupObjs[$k].ParentNode.RemoveChild($groupObjs[$k])   | Out-Null
    }

    $groupCount++
}

$leftover = $buildItems.Count % $MASTER_MERGE_COUNT
if ($leftover -ne 0) { Write-Host "Note: $leftover build item(s) left ungrouped (not a multiple of $MASTER_MERGE_COUNT)." }

# ── Finalize ──────────────────────────────────────────────────────────────────
# Collapse plate model_instances (keep only one per group — already done per
# group above for assemble_items; the plate instances mirror that collapse)
if ($hasSettings -and ($plate = $settings.SelectSingleNode('//plate'))) {
    $inst = @($plate.SelectNodes('model_instance'))
    # Keep only enough instances for surviving assemblies (one per group)
    $survivingCount = [math]::Floor($buildItems.Count / $MASTER_MERGE_COUNT)
    for ($j = $survivingCount; $j -lt $inst.Count; $j++) { $plate.RemoveChild($inst[$j]) | Out-Null }
}

# cut_information.xml: keep one entry per surviving assembly, delete rest
if (($null -ne $cutInfoPath) -and (Test-Path $cutInfoPath)) {
    [xml]$cutXml = [System.IO.File]::ReadAllText($cutInfoPath, [System.Text.Encoding]::UTF8)
    $cutObjs = @($cutXml.SelectNodes('//*[local-name()="object"]'))
    $survivingCount = [math]::Floor($buildItems.Count / $MASTER_MERGE_COUNT)
    for ($j = $survivingCount; $j -lt $cutObjs.Count; $j++) { $cutObjs[$j].ParentNode.RemoveChild($cutObjs[$j]) | Out-Null }
    $ws = New-Object System.Xml.XmlWriterSettings; $ws.Encoding = New-Object System.Text.UTF8Encoding($false); $ws.Indent = $true
    $w = [System.Xml.XmlWriter]::Create($cutInfoPath, $ws); $cutXml.Save($w); $w.Close()
}

# layer_heights_profile.txt: delete entirely
if (($null -ne $layerProfPath) -and (Test-Path $layerProfPath)) { Remove-Item $layerProfPath -Force }

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
Write-Host "Success: $groupCount group(s) merged at N=$MASTER_MERGE_COUNT. Total input items: $($buildItems.Count)."