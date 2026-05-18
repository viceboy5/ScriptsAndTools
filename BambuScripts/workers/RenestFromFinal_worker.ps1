param(
    [string]$FinalPath,                # The edited Final.3mf (single master object)
    [string]$TransformSourcePath = "", # Nest.3mf or Full.3mf to read plate transforms from
    [string]$OutputPath          = "", # Defaults to <stem>_Renest.3mf next to Final
    [string]$BambuPath           = "C:\Program Files\Bambu Studio\bambu-studio.exe"
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

# ════════════════════════════════════════════════════════════════════════════════
#  RenestFromFinal_worker.ps1
#
#  Takes an edited Final.3mf (single object, any changes applied) and a
#  Nest.3mf or Full.3mf that carries the original plate layout.
#
#  Outputs a new full-plate 3MF where every instance is a copy of the edited
#  Final object placed at the transforms from the source plate.
#
#  Workflow:
#    1. Edit your Final.3mf however you like in Bambu Studio
#    2. Drop the Final.3mf onto RenestFromFinal.bat
#    3. The script finds the sibling Nest/Full automatically for transforms
#    4. Output is a ready-to-slice plate file with all your edits applied
# ════════════════════════════════════════════════════════════════════════════════

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

# ── Helpers ───────────────────────────────────────────────────────────────────
function Parse-Tx([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return [double[]](1,0,0, 0,1,0, 0,0,1, 0,0,0) }
    return [double[]]($s.Trim() -split '\s+')
}
# 3x3 matrix helpers (row-major, indices 0-8)
function Mul-3x3([double[]]$a, [double[]]$b) {
    $r = New-Object double[] 9
    $r[0] = $a[0]*$b[0] + $a[1]*$b[3] + $a[2]*$b[6]
    $r[1] = $a[0]*$b[1] + $a[1]*$b[4] + $a[2]*$b[7]
    $r[2] = $a[0]*$b[2] + $a[1]*$b[5] + $a[2]*$b[8]
    $r[3] = $a[3]*$b[0] + $a[4]*$b[3] + $a[5]*$b[6]
    $r[4] = $a[3]*$b[1] + $a[4]*$b[4] + $a[5]*$b[7]
    $r[5] = $a[3]*$b[2] + $a[4]*$b[5] + $a[5]*$b[8]
    $r[6] = $a[6]*$b[0] + $a[7]*$b[3] + $a[8]*$b[6]
    $r[7] = $a[6]*$b[1] + $a[7]*$b[4] + $a[8]*$b[7]
    $r[8] = $a[6]*$b[2] + $a[7]*$b[5] + $a[8]*$b[8]
    return $r
}
function Transpose-3x3([double[]]$m) {
    return [double[]]($m[0],$m[3],$m[6], $m[1],$m[4],$m[7], $m[2],$m[5],$m[8])
}
function Get-TxRot([double[]]$tx) { return [double[]]($tx[0],$tx[1],$tx[2], $tx[3],$tx[4],$tx[5], $tx[6],$tx[7],$tx[8]) }
function Apply-RotCorrection([double[]]$tx, [double[]]$corr) {
    $r = Mul-3x3 $corr (Get-TxRot $tx)
    return [double[]]($r[0],$r[1],$r[2], $r[3],$r[4],$r[5], $r[6],$r[7],$r[8], $tx[9],$tx[10],$tx[11])
}
function Is-IdentityRot([double[]]$r) {
    $eps = 1e-6
    return ([math]::Abs($r[0]-1) -lt $eps -and [math]::Abs($r[4]-1) -lt $eps -and [math]::Abs($r[8]-1) -lt $eps -and
            [math]::Abs($r[1]) -lt $eps -and [math]::Abs($r[2]) -lt $eps -and [math]::Abs($r[3]) -lt $eps -and
            [math]::Abs($r[5]) -lt $eps -and [math]::Abs($r[6]) -lt $eps -and [math]::Abs($r[7]) -lt $eps)
}
function Save-Xml([xml]$doc, [string]$path) {
    $ws = New-Object System.Xml.XmlWriterSettings
    $ws.Encoding = New-Object System.Text.UTF8Encoding($false); $ws.Indent = $true
    $w = [System.Xml.XmlWriter]::Create($path, $ws); $doc.Save($w); $w.Close()
}
function Read-ZipEntry([string]$zipPath, [string]$entryName) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entry = $zip.GetEntry($entryName)
        if ($null -eq $entry) { return $null }
        $sr = New-Object System.IO.StreamReader($entry.Open())
        $content = $sr.ReadToEnd(); $sr.Close(); return $content
    } finally { $zip.Dispose() }
}
function Find-File([string]$base, [string]$rel) {
    $p = Join-Path $base ($rel -replace '/', [System.IO.Path]::DirectorySeparatorChar)
    if (Test-Path $p) { return $p }
    return $null
}

# ── Resolve input paths ───────────────────────────────────────────────────────
$FinalPath = $FinalPath.Trim('"')
if (-not (Test-Path $FinalPath)) { Write-Error "Final file not found: $FinalPath"; exit 1 }

$finalDir  = Split-Path $FinalPath -Parent
$finalStem = [System.IO.Path]::GetFileNameWithoutExtension($FinalPath) -replace '(?i)_Final$', ''

# Auto-detect transform source:
#   If both Nest and Full exist alongside Final → use Nest (file has been merged)
#   If only Full exists → use Full
if ([string]::IsNullOrWhiteSpace($TransformSourcePath)) {
    $nestCandidate = Join-Path $finalDir "${finalStem}_Nest.3mf"
    $fullCandidate = Join-Path $finalDir "${finalStem}_Full.3mf"
    $nestExists    = Test-Path $nestCandidate
    $fullExists    = Test-Path $fullCandidate

    if ($nestExists -and $fullExists) {
        $TransformSourcePath = $nestCandidate
        $autoDetectReason    = "Both Nest and Full exist - using Nest (merged plate)"
    } elseif ($nestExists) {
        $TransformSourcePath = $nestCandidate
        $autoDetectReason    = "Only Nest found - using Nest"
    } elseif ($fullExists) {
        $TransformSourcePath = $fullCandidate
        $autoDetectReason    = "Only Full found - using Full"
    }
}
if ([string]::IsNullOrWhiteSpace($TransformSourcePath)) {
    Write-Error "Could not find a Nest or Full file alongside the Final. Pass -TransformSourcePath explicitly."
    exit 1
}
$TransformSourcePath = $TransformSourcePath.Trim('"')

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $finalDir "${finalStem}_Renest.3mf"
}

# ── Confirm before proceeding ─────────────────────────────────────────────────
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  RenestFromFinal - Planned Operation"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host ""
Write-Host "  Master (edited object)  : $(Split-Path $FinalPath -Leaf)"
if ($autoDetectReason) {
Write-Host "  Transform source        : $(Split-Path $TransformSourcePath -Leaf)  [$autoDetectReason]"
} else {
Write-Host "  Transform source        : $(Split-Path $TransformSourcePath -Leaf)"
}
Write-Host "  Output                  : $(Split-Path $OutputPath -Leaf)"
Write-Host ""
Write-Host "  The edited master object from Final will be cloned once per"
Write-Host "  transform found in the source plate, then saved to Output."
Write-Host ""
$confirm = (Read-Host "  Proceed? [Y/N]").Trim()
if ($confirm -notmatch '^[Yy]') {
    Write-Host "Cancelled."
    exit 0
}
Write-Host ""

# ════════════════════════════════════════════════════════════════════════════════
#  STEP 1 — Read transforms from the source plate file
# ════════════════════════════════════════════════════════════════════════════════
$srcModelText    = Read-ZipEntry $TransformSourcePath '3D/3dmodel.model'
$srcSettingsText = Read-ZipEntry $TransformSourcePath 'Metadata/model_settings.config'
if ($null -eq $srcModelText) { Write-Error "Cannot read 3D/3dmodel.model from transform source."; exit 1 }

[xml]$srcModel = $srcModelText
$srcXns = New-Object System.Xml.XmlNamespaceManager($srcModel.NameTable)
$srcXns.AddNamespace('m', $nsCore); $srcXns.AddNamespace('p', $nsProd)

$srcBuildItems = @($srcModel.SelectNodes('//m:build/m:item', $srcXns))

# Build lookup: objectid -> settings object node (for outlier filtering)
$srcSettObjById = @{}
$srcIdentifyById = @{}
if ($null -ne $srcSettingsText) {
    [xml]$srcSettings = $srcSettingsText
    foreach ($obj in $srcSettings.SelectNodes('//*[local-name()="object"]')) {
        $srcSettObjById[$obj.GetAttribute('id')] = $obj
    }
    foreach ($inst in $srcSettings.SelectNodes('//plate/model_instance')) {
        $oidMeta = $inst.SelectSingleNode('metadata[@key="object_id"]')
        $iidMeta = $inst.SelectSingleNode('metadata[@key="identify_id"]')
        if ($null -ne $oidMeta -and $null -ne $iidMeta) {
            $srcIdentifyById[$oidMeta.GetAttribute('value')] = $iidMeta.GetAttribute('value')
        }
    }
}

# Face-count majority filter - excludes text labels, version stamps, etc.
$fcMap = @{}
foreach ($item in $srcBuildItems) {
    $id = $item.GetAttribute('objectid'); $fc = 'unknown'
    if ($null -ne $srcSettObjById[$id]) {
        $fcNode = $srcSettObjById[$id].SelectSingleNode('metadata[@face_count]')
        if ($null -ne $fcNode) { $fc = $fcNode.GetAttribute('face_count') }
    }
    if (-not $fcMap.Contains($fc)) { $fcMap[$fc] = 0 }
    $fcMap[$fc]++
}
$majorityFc = ($fcMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Name

$sourceTransforms  = [System.Collections.Generic.List[string]]::new()
$sourceIdentifyIds = [System.Collections.Generic.List[string]]::new()

foreach ($item in $srcBuildItems) {
    $id = $item.GetAttribute('objectid'); $keep = $true
    if ($null -ne $srcSettObjById[$id]) {
        $fcNode   = $srcSettObjById[$id].SelectSingleNode('metadata[@face_count]')
        $fc       = if ($null -ne $fcNode) { $fcNode.GetAttribute('face_count') } else { 'unknown' }
        if ($fc -ne $majorityFc) { $keep = $false }
        $nameNode = $srcSettObjById[$id].SelectSingleNode('metadata[@key="name"]')
        if ($null -ne $nameNode) {
            $n = $nameNode.GetAttribute('value')
            if ($n -match '(?i)text|version' -and $n -notmatch '(?i)\.(stl|3mf|obj|step|stp)$') { $keep = $false }
        }
    }
    if ($keep) {
        $sourceTransforms.Add($item.GetAttribute('transform'))
        $iid = if ($srcIdentifyById.Contains($id)) { $srcIdentifyById[$id] } else { ($sourceTransforms.Count * 442).ToString() }
        $sourceIdentifyIds.Add($iid)
    }
}

$n = $sourceTransforms.Count
Write-Host "Found $n instance transform(s) to replicate."
if ($n -eq 0) { Write-Error "No valid transforms found in source file."; exit 1 }

# Grab the component[0] rotation from the source template assembly.
# Used later to detect if the Final master was re-oriented by Bambu Studio.
$srcTemplateCompRot = $null
$srcTemplateObj = $srcModel.SelectSingleNode('//m:resources/m:object[m:components/m:component]', $srcXns)
if ($null -ne $srcTemplateObj) {
    $srcComp0 = $srcTemplateObj.SelectSingleNode('m:components/m:component', $srcXns)
    if ($null -ne $srcComp0) {
        $srcTemplateCompRot = Get-TxRot (Parse-Tx $srcComp0.GetAttribute('transform'))
    }
}

# ════════════════════════════════════════════════════════════════════════════════
#  STEP 2 — Extract Final.3mf to working directory
# ════════════════════════════════════════════════════════════════════════════════
$workDir = Join-Path $env:TEMP ("Renest_" + [guid]::NewGuid().ToString().Substring(0,8))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($FinalPath, $workDir)

    $modelFilePath = (Get-ChildItem -Path $workDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
    $settingsPath  = Find-File $workDir 'Metadata/model_settings.config'
    $cutInfoPath   = Find-File $workDir 'Metadata/cut_information.xml'
    $vlhPath       = Join-Path $workDir "Metadata\layer_heights_profile.txt"
    $relsPath      = Find-File $workDir '3D/_rels/3dmodel.model.rels'

    [xml]$xml = [System.IO.File]::ReadAllText($modelFilePath, [System.Text.Encoding]::UTF8)
    $xns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $xns.AddNamespace('m', $nsCore); $xns.AddNamespace('p', $nsProd)

    $buildNode = $xml.SelectSingleNode('//m:build', $xns)
    $buildItems = @($xml.SelectNodes('//m:build/m:item', $xns))
    if ($buildItems.Count -eq 0) { Write-Error "No build items in Final.3mf"; exit 1 }

    # ── Identify the master (printable) assembly object ───────────────────────
    # If Final has more than one (e.g. after re-editing in Bambu), pick centre-most
    $masterItem = $buildItems[0]
    if ($buildItems.Count -gt 1) {
        $minDist = [double]::MaxValue
        foreach ($bi in $buildItems) {
            $tx = Parse-Tx ($bi.GetAttribute('transform'))
            $d  = [math]::Pow($tx[9]-128,2) + [math]::Pow($tx[10]-128,2)
            if ($d -lt $minDist) { $minDist = $d; $masterItem = $bi }
        }
    }
    $masterId = $masterItem.GetAttribute('objectid')

    $objById = @{}
    foreach ($o in $xml.SelectNodes('//m:resources/m:object', $xns)) { $objById[$o.GetAttribute('id')] = $o }
    $masterObj = $objById[$masterId]

    # Collect which paths are needed (keep .rels references intact)
    $usedPaths = New-Object System.Collections.Generic.HashSet[string]
    foreach ($node in $xml.SelectNodes('//*[@*[local-name()="path"]]')) {
        $p = $node.GetAttribute('path', $nsProd)
        if ([string]::IsNullOrWhiteSpace($p)) { $p = $node.GetAttribute('p:path') }
        if (-not [string]::IsNullOrWhiteSpace($p)) {
            if (-not $p.StartsWith('/')) { $p = '/' + $p }
            $usedPaths.Add($p) | Out-Null
        }
    }

    # ── Detect orientation mismatch between Final master and source template ──────
    # When Bambu Studio re-orients an object (e.g. after adding a modifier) it bakes
    # a rotation into ALL existing component transforms.  The source plate transforms
    # were designed for the ORIGINAL orientation, so we must undo the difference.
    $rotCorrection = [double[]](1,0,0, 0,1,0, 0,0,1)   # identity = no correction
    $masterComp0 = $masterObj.SelectSingleNode('m:components/m:component', $xns)
    if ($null -ne $masterComp0 -and $null -ne $srcTemplateCompRot) {
        $finalCompRot = Get-TxRot (Parse-Tx $masterComp0.GetAttribute('transform'))
        # R_correction_inv = Transpose(R_final) * R_src_template
        # Applying this to each plate transform un-bakes the Final rotation
        # and re-expresses the transform in the source template's frame.
        $candidate = Mul-3x3 (Transpose-3x3 $finalCompRot) $srcTemplateCompRot
        if (-not (Is-IdentityRot $candidate)) {
            $rotCorrection = $candidate
            Write-Host "Orientation correction applied (Final master was re-oriented vs source template)."
        }
    }

    # ── Snapshot component UUIDs base from the master ──────────────────────────
    # We'll regenerate UUIDs per-clone to avoid duplicates
    $masterComps = @($masterObj.SelectNodes('m:components/m:component', $xns))

    # ── Remove all existing build items ───────────────────────────────────────
    foreach ($bi in $buildItems) { $bi.ParentNode.RemoveChild($bi) | Out-Null }

    # Remove all existing printable (assembly) objects, keep internal meshes
    $hasSettings = ($null -ne $settingsPath) -and (Test-Path $settingsPath)
    $printableIdsInFinal = New-Object System.Collections.Generic.HashSet[string]
    if ($hasSettings) {
        [xml]$settingsTmp = [System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)
        foreach ($obj in $settingsTmp.SelectNodes('//*[local-name()="object"]')) {
            $printableIdsInFinal.Add($obj.GetAttribute('id')) | Out-Null
        }
    } else {
        $printableIdsInFinal.Add($masterId) | Out-Null
    }

    foreach ($id in @($printableIdsInFinal)) {
        $obj = $objById[$id]
        if ($null -ne $obj -and $null -ne $obj.ParentNode) { $obj.ParentNode.RemoveChild($obj) | Out-Null }
    }

    $resourcesNode = $xml.SelectSingleNode('//m:resources', $xns)

    # ── Find the highest ID in any external object file (e.g. object_1.model) ──
    # These objects live in 3D/Objects/ and are referenced via p:path components.
    # Their IDs share a global namespace with the main model in Bambu Studio,
    # so our assembly IDs must never collide with them.
    $maxExternalId = 0
    $objFilesDir = Join-Path (Split-Path $modelFilePath -Parent) 'Objects'
    if (Test-Path $objFilesDir) {
        foreach ($extFile in (Get-ChildItem -LiteralPath $objFilesDir -Filter '*.model')) {
            [xml]$extXml = [System.IO.File]::ReadAllText($extFile.FullName, [System.Text.Encoding]::UTF8)
            foreach ($extObj in $extXml.SelectNodes('//*[local-name()="object"]')) {
                $v = 0; if ([int]::TryParse($extObj.GetAttribute('id'), [ref]$v) -and $v -gt $maxExternalId) { $maxExternalId = $v }
            }
        }
        if ($maxExternalId -gt 0) { Write-Host "Max ID in external object files: $maxExternalId" }
    }

    # ── Find the highest internal-mesh ID to start our new IDs above it ────────
    $nextId = $maxExternalId + 1
    foreach ($id in $objById.Keys) {
        if ($printableIdsInFinal.Contains($id)) { continue }
        $v = 0; if ([int]::TryParse($id, [ref]$v) -and $v -ge $nextId) { $nextId = $v + 1 }
    }

    # ════════════════════════════════════════════════════════════════════════════
    #  STEP 3 — Clone master object N times, one per source transform
    # ════════════════════════════════════════════════════════════════════════════
    $newObjIds = [System.Collections.Generic.List[string]]::new()
    $uuidObjCounter = 1

    for ($i = 0; $i -lt $n; $i++) {
        $newId = ($nextId + $i).ToString()
        $newObjIds.Add($newId)

        # Deep clone the master assembly object
        $clone = $masterObj.CloneNode($true)
        $clone.SetAttribute('id', $newId)

        # Fresh object UUID
        $clone.SetAttribute('UUID', $nsProd, ($uuidObjCounter.ToString("x8") + "-71cb-4c03-9d28-80fed5dfa1dc")) | Out-Null
        $uuidObjCounter++

        # Fresh component UUIDs within this clone
        $cloneComps = @($clone.SelectNodes('m:components/m:component', $xns))
        $uuidCompBase = ($uuidObjCounter * 65536)
        foreach ($cc in $cloneComps) {
            $cc.SetAttribute('UUID', $nsProd, ($uuidCompBase.ToString("x8") + "-comp-4c03-9d28-80fed5dfa1dc")) | Out-Null
            $uuidCompBase++
        }

        $resourcesNode.AppendChild($clone) | Out-Null

        # Build item with this transform (apply orientation correction if needed)
        $newItem = $xml.CreateElement('item', $nsCore)
        $newItem.SetAttribute('objectid', $newId)
        $tx = $sourceTransforms[$i]
        if (-not [string]::IsNullOrWhiteSpace($tx)) {
            if (-not (Is-IdentityRot $rotCorrection)) {
                $corrected = Apply-RotCorrection (Parse-Tx $tx) $rotCorrection
                $tx = ($corrected | ForEach-Object { ([double]$_).ToString('G9') }) -join ' '
            }
            $newItem.SetAttribute('transform', $tx)
        }
        $newItem.SetAttribute('UUID', $nsProd, ($i.ToString("x8") + "-b1ec-4553-aec9-835e5b724bb4")) | Out-Null
        $newItem.SetAttribute('printable', '1')
        $buildNode.AppendChild($newItem) | Out-Null
    }

    Write-Host "Cloned master object x$n in 3dmodel.model."

    # ════════════════════════════════════════════════════════════════════════════
    #  STEP 4 — Rebuild model_settings.config
    # ════════════════════════════════════════════════════════════════════════════
    if ($hasSettings) {
        [xml]$settings = [System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)

        # Grab the master settings object before we nuke it
        $masterSettObj = $null
        foreach ($obj in $settings.SelectNodes('//*[local-name()="object"]')) {
            if ($printableIdsInFinal.Contains($obj.GetAttribute('id'))) { $masterSettObj = $obj; break }
        }

        # Remove all printable object entries
        foreach ($obj in @($settings.SelectNodes('//*[local-name()="object"]'))) {
            if ($printableIdsInFinal.Contains($obj.GetAttribute('id'))) {
                $obj.ParentNode.RemoveChild($obj) | Out-Null
            }
        }

        # Remove and recreate assemble + plate sections
        $configNode = $settings.DocumentElement
        foreach ($tag in @('assemble','plate')) {
            $node = $settings.SelectSingleNode("//$tag")
            if ($null -ne $node) { $node.ParentNode.RemoveChild($node) | Out-Null }
        }

        $newAssemble = $settings.CreateElement('assemble')
        $configNode.AppendChild($newAssemble) | Out-Null

        # Copy plate metadata from the source settings (preserves filament_map_mode,
        # locked, thumbnail paths, etc.).  Fall back to hardcoded defaults if missing.
        $newPlate = $settings.CreateElement('plate')
        $srcPlateNode = $settings.SelectSingleNode('//plate')   # original (now removed) — use Final's
        # Re-read original Final plate metadata from the pre-edit XML snapshot
        $stdPlateMeta = [ordered]@{
            'plater_id'              = '1'
            'plater_name'            = ''
            'locked'                 = 'false'
            'filament_map_mode'      = 'Auto For Flush'
            'filament_maps'          = '1 1 1 1'
            'thumbnail_file'         = 'Metadata/plate_1.png'
            'thumbnail_no_light_file'= 'Metadata/plate_no_light_1.png'
            'top_file'               = 'Metadata/top_1.png'
            'pick_file'              = 'Metadata/pick_1.png'
        }
        foreach ($kv in $stdPlateMeta.GetEnumerator()) {
            $m = $settings.CreateElement('metadata')
            $m.SetAttribute('key',   $kv.Key)
            $m.SetAttribute('value', $kv.Value)
            $newPlate.AppendChild($m) | Out-Null
        }
        $configNode.AppendChild($newPlate) | Out-Null

        for ($i = 0; $i -lt $n; $i++) {
            $newId = $newObjIds[$i]
            $tx    = $sourceTransforms[$i]
            # Apply same orientation correction as the build item
            if (-not (Is-IdentityRot $rotCorrection)) {
                $corrected = Apply-RotCorrection (Parse-Tx $tx) $rotCorrection
                $tx = ($corrected | ForEach-Object { ([double]$_).ToString('G9') }) -join ' '
            }

            # Clone master settings entry with new ID
            if ($null -ne $masterSettObj) {
                $clonedSett = $settings.ImportNode($masterSettObj, $true)
                $clonedSett.SetAttribute('id', $newId)
                $configNode.InsertBefore($clonedSett, $newAssemble) | Out-Null
            }

            # assemble_item — match Bambu's format (offset not modelmesh_id)
            $asmItem = $settings.CreateElement('assemble_item')
            $asmItem.SetAttribute('object_id',   $newId)
            $asmItem.SetAttribute('instance_id', '0')
            $asmItem.SetAttribute('transform',   $tx)
            $asmItem.SetAttribute('offset',      '0 0 0')
            $newAssemble.AppendChild($asmItem) | Out-Null

            # model_instance in plate — object_id, instance_id, identify_id (no 'centered')
            $inst = $settings.CreateElement('model_instance')
            foreach ($kv in @( @('object_id',$newId), @('instance_id','0'), @('identify_id',$sourceIdentifyIds[$i]) )) {
                $m = $settings.CreateElement('metadata')
                $m.SetAttribute('key',   $kv[0])
                $m.SetAttribute('value', $kv[1])
                $inst.AppendChild($m) | Out-Null
            }
            $newPlate.AppendChild($inst) | Out-Null
        }

        Save-Xml $settings $settingsPath
        Write-Host "Rebuilt model_settings.config ($n object entries)."
    }

    # ════════════════════════════════════════════════════════════════════════════
    #  STEP 5 — Global ID renumbering (internal meshes first, assemblies after)
    # ════════════════════════════════════════════════════════════════════════════
    $printableSet = New-Object System.Collections.Generic.HashSet[string]
    $newObjIds | ForEach-Object { $printableSet.Add($_) | Out-Null }

    $allObjs = @($xml.SelectNodes('//m:resources/m:object', $xns))
    $sorted  = $allObjs | Sort-Object `
        @{ Expression = { if ($printableSet.Contains($_.GetAttribute('id'))) { 1 } else { 0 } }; Ascending = $true },
        @{ Expression = { [int]$_.GetAttribute('id') }; Ascending = $true }

    # Mesh objects (non-assembly) get sequential IDs starting at 1.
    # Assembly objects must start AFTER both the mesh objects AND all external IDs.
    $internalMeshCount  = ($allObjs | Where-Object { -not $printableSet.Contains($_.GetAttribute('id')) }).Count
    $firstAssemblyId    = [Math]::Max($internalMeshCount + 1, $maxExternalId + 1)
    $meshCounter        = 1
    $assemblyCounter    = $firstAssemblyId

    $idMap = @{}
    foreach ($obj in $sorted) {
        $old = $obj.GetAttribute('id')
        if ($printableSet.Contains($old)) {
            $newId = ($assemblyCounter++).ToString()
        } else {
            $newId = ($meshCounter++).ToString()
        }
        $idMap[$old] = $newId
        $obj.SetAttribute('id', $newId)
        $resourcesNode.AppendChild($obj) | Out-Null
    }

    # Remap build items
    foreach ($bi in @($xml.SelectNodes('//m:build/m:item', $xns))) {
        $old = $bi.GetAttribute('objectid'); if ($idMap.Contains($old)) { $bi.SetAttribute('objectid', $idMap[$old]) }
    }
    # Remap component objectids
    foreach ($comp in @($xml.SelectNodes('//m:components/m:component', $xns))) {
        $old = $comp.GetAttribute('objectid'); if ($idMap.Contains($old)) { $comp.SetAttribute('objectid', $idMap[$old]) }
    }

    Save-Xml $xml $modelFilePath

    # Remap in model_settings.config
    if ($hasSettings) {
        [xml]$s2 = [System.IO.File]::ReadAllText($settingsPath, [System.Text.Encoding]::UTF8)
        foreach ($obj  in @($s2.SelectNodes('//*[local-name()="object"]')))                             { $old = $obj.GetAttribute('id');           if ($idMap[$old]) { $obj.SetAttribute('id',          $idMap[$old]) } }
        foreach ($part in @($s2.SelectNodes('//*[local-name()="part"]')))                              { $old = $part.GetAttribute('id');          if ($idMap[$old]) { $part.SetAttribute('id',         $idMap[$old]) } }
        foreach ($asm  in @($s2.SelectNodes('//assemble/assemble_item')))                              { $old = $asm.GetAttribute('object_id');    if ($idMap[$old]) { $asm.SetAttribute('object_id',   $idMap[$old]) } }
        foreach ($meta in @($s2.SelectNodes('//plate/model_instance/metadata[@key="object_id"]')))    { $old = $meta.GetAttribute('value');        if ($idMap[$old]) { $meta.SetAttribute('value',      $idMap[$old]) } }
        Save-Xml $s2 $settingsPath
    }

    # Remap cut_information.xml
    if ($null -ne $cutInfoPath -and (Test-Path $cutInfoPath)) {
        [xml]$cutXml = [System.IO.File]::ReadAllText($cutInfoPath, [System.Text.Encoding]::UTF8)
        $cutMod = $false
        foreach ($co in @($cutXml.SelectNodes('//*[local-name()="object"]'))) {
            $old = $co.GetAttribute('id')
            if ($idMap.Contains($old)) { $co.SetAttribute('id', $idMap[$old]); $cutMod = $true }
            else                       { $co.ParentNode.RemoveChild($co) | Out-Null; $cutMod = $true }
        }
        if ($cutMod) { Save-Xml $cutXml $cutInfoPath }
    }

    Write-Host "ID renumbering complete ($($sorted.Count) total objects)."

    # ════════════════════════════════════════════════════════════════════════════
    #  STEP 6 — Rebuild VLH, purge stale thumbnails
    # ════════════════════════════════════════════════════════════════════════════
    if (Test-Path $vlhPath) {
        $vlhLines = @(Get-Content $vlhPath)
        $vlhData  = $null
        foreach ($line in $vlhLines) { if ($line -match '\|(.+)$') { $vlhData = $matches[1]; break } }
        if ($null -ne $vlhData) {
            # VLH uses sequential plate indices 1..N, NOT model_settings object IDs.
            # (e.g. a Nest with model_settings IDs 46-141 still has VLH IDs 1-96)
            $newVlh = 1..$n | ForEach-Object { "object_id=$_|$vlhData" }
            [System.IO.File]::WriteAllText($vlhPath, ($newVlh -join "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
        }
    }

    $metaDir = Join-Path $workDir "Metadata"
    foreach ($pat in @("plate_*.png","pick_*.png","plate_*.json")) {
        Get-ChildItem $metaDir -Filter $pat -ErrorAction SilentlyContinue | Remove-Item -Force
    }

    # ════════════════════════════════════════════════════════════════════════════
    #  STEP 7 — Repack
    # ════════════════════════════════════════════════════════════════════════════
    $tempOut      = $OutputPath + ".renest.tmp.3mf"
    if (Test-Path $tempOut) { Remove-Item $tempOut -Force }

    $resolvedWork = (Get-Item -LiteralPath $workDir).FullName.TrimEnd('\','/') + '\'
    $zip = [System.IO.Compression.ZipFile]::Open($tempOut, 'Create')
    try {
        Get-ChildItem -LiteralPath $resolvedWork -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($resolvedWork.Length).Replace('\','/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
        }
    } finally { $zip.Dispose() }

    Move-Item -LiteralPath $tempOut -Destination $OutputPath -Force
    Write-Host ""
    Write-Host "Output: $OutputPath"

    # ════════════════════════════════════════════════════════════════════════════
    #  STEP 8 — Optional Bambu resave (cleans stale metadata, generates thumbnails)
    # ════════════════════════════════════════════════════════════════════════════
    if (Test-Path $BambuPath) {
        Write-Host "Running Bambu Studio resave to clean metadata..." -NoNewline
        $tempResave = $OutputPath + ".resave.tmp.3mf"
        $logOut     = Join-Path $env:TEMP "bambu_renest_out.txt"
        $logErr     = Join-Path $env:TEMP "bambu_renest_err.txt"
        $procArgs   = "--debug 3 --no-check --uptodate --allow-newer-file --export-3mf `"$tempResave`" `"$OutputPath`""
        $proc       = Start-Process -FilePath $BambuPath -ArgumentList $procArgs `
                                    -RedirectStandardOutput $logOut -RedirectStandardError $logErr `
                                    -WindowStyle Hidden -PassThru
        $proc.WaitForExit()
        foreach ($log in @($logOut, $logErr)) { if (Test-Path $log) { Remove-Item $log -Force -ErrorAction SilentlyContinue } }
        if (Test-Path $tempResave) { Move-Item $tempResave $OutputPath -Force; Write-Host " done." }
        else { Write-Host " skipped (export produced no output)." }
    }

    Write-Host "Done! $n instances written to: $OutputPath"

} finally {
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
