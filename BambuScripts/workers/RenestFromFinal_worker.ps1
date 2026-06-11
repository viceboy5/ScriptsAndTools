param(
    [string]$FinalPath,                # The edited Final.3mf (single master object)
    [string]$TransformSourcePath = "", # Nest.3mf or Full.3mf to read plate transforms from
    [string]$OutputPath          = "", # Defaults to <stem>_Renest.3mf next to Final
    [string]$BambuPath           = "C:\Program Files\Bambu Studio\bambu-studio.exe",
    [string]$RotCorrectionPath   = "", # Explicit rotation-correction JSON (e.g. a temp file from
                                       # the CardQueueEditor). Defaults to <stem>_RotCorrection.json
                                       # next to Final when omitted.
    [switch]$NoConfirm                 # Skip the Y/N prompt (used when caller batches multiple files)
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

$nsCore = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
$nsProd = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

# â”€â”€ Helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
# Rotate a 3-vector by a 3x3 rotation matrix (row-vector convention: v' = v * R)
function Rotate-Vec([double[]]$v, [double[]]$r) {
    $x = $v[0]*$r[0]+$v[1]*$r[3]+$v[2]*$r[6]
    $y = $v[0]*$r[1]+$v[1]*$r[4]+$v[2]*$r[7]
    $z = $v[0]*$r[2]+$v[1]*$r[5]+$v[2]*$r[8]
    return [double[]]($x, $y, $z)
}
# Local-frame offset between the source template object and the Final object,
# found by MATCHING shared components: pieces the user added, removed, moved,
# or duplicated cannot influence it, while the unmoved shared components all
# differ by exactly the same vector - the true frame offset. Used as a SANITY
# CHECK only (clone translations are taken verbatim from the source plate; a
# large offset here means the Final's frame genuinely moved and placement may
# be off). Takes every (source comp - final comp) translation difference,
# buckets them on a 0.5 mm grid, then refines around the densest bucket.
# Returns (dx, dy, dz, matchedCount), or $null when fewer than max(3, 20% of
# the smaller component set) agree.
function Get-MatchedFrameOffset($srcTxList, $finTxList, [double[]]$rotCorr) {
    $srcPts = [System.Collections.Generic.List[double[]]]::new()
    foreach ($s in $srcTxList) {
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        $tx = Parse-Tx $s
        if (Is-IdentityRot (Get-TxRot $tx)) { $srcPts.Add([double[]]($tx[9],$tx[10],$tx[11])) | Out-Null }
    }
    $finPts = [System.Collections.Generic.List[double[]]]::new()
    foreach ($s in $finTxList) {
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        $tx = Parse-Tx $s
        if (Is-IdentityRot (Get-TxRot $tx)) {
            # Same frame convention as the centroid path: compare in the
            # rotation-corrected final frame.
            $finPts.Add((Rotate-Vec ([double[]]($tx[9],$tx[10],$tx[11])) $rotCorr)) | Out-Null
        }
    }
    if ($srcPts.Count -eq 0 -or $finPts.Count -eq 0) { return $null }
    $grid = 0.5
    $buckets = @{}
    foreach ($sp in $srcPts) {
        foreach ($fp in $finPts) {
            $dx = $sp[0]-$fp[0]; $dy = $sp[1]-$fp[1]; $dz = $sp[2]-$fp[2]
            $key = "$([Math]::Round($dx/$grid))|$([Math]::Round($dy/$grid))|$([Math]::Round($dz/$grid))"
            if (-not $buckets.ContainsKey($key)) { $buckets[$key] = @{ n = 0; sx = 0.0; sy = 0.0; sz = 0.0 } }
            $b = $buckets[$key]
            $b.n++; $b.sx += $dx; $b.sy += $dy; $b.sz += $dz
        }
    }
    $best = $null
    foreach ($b in $buckets.Values) { if ($null -eq $best -or $b.n -gt $best.n) { $best = $b } }
    if ($null -eq $best -or $best.n -lt 3) { return $null }
    # Refine: a true cluster can straddle grid-bucket boundaries, splitting its
    # count. Re-scan all pairs against the best bucket's mean with a 1 mm
    # radius and average the matches - this recovers the full cluster and the
    # exact offset. Each final component is counted at most once (its best
    # source partner) so duplicated source geometry can't inflate the count.
    $bx = $best.sx/$best.n; $by = $best.sy/$best.n; $bz = $best.sz/$best.n
    $rn = 0; $rsx = 0.0; $rsy = 0.0; $rsz = 0.0
    foreach ($fp in $finPts) {
        $bestD = [double]::MaxValue; $bdx = 0.0; $bdy = 0.0; $bdz = 0.0
        foreach ($sp in $srcPts) {
            $dx = $sp[0]-$fp[0]; $dy = $sp[1]-$fp[1]; $dz = $sp[2]-$fp[2]
            $d = [Math]::Sqrt(($dx-$bx)*($dx-$bx) + ($dy-$by)*($dy-$by) + ($dz-$bz)*($dz-$bz))
            if ($d -lt $bestD) { $bestD = $d; $bdx = $dx; $bdy = $dy; $bdz = $dz }
        }
        if ($bestD -le 1.0) { $rn++; $rsx += $bdx; $rsy += $bdy; $rsz += $bdz }
    }
    # Accept when at least 3 components and a fifth of the smaller set agree -
    # the agreeing components are the unmoved ones that define the frame.
    $minCount = [Math]::Max(3, [int][Math]::Ceiling([Math]::Min($srcPts.Count, $finPts.Count) * 0.2))
    if ($rn -lt $minCount) { return $null }
    return [double[]](($rsx/$rn), ($rsy/$rn), ($rsz/$rn), $rn)
}

# Extract the uniform scale factor embedded in a 3x3 rotation+scale matrix.
# Scale = norm of the first row (all rows should be equal for uniform scale).
function Get-RowScale([double[]]$r) {
    $s = [Math]::Sqrt($r[0]*$r[0] + $r[1]*$r[1] + $r[2]*$r[2])
    if ($s -lt 1e-9) { return 1.0 }
    return $s
}
# Strip any uniform scale baked into a row-vector rotation matrix
function Normalize-Rot3x3([double[]]$r) {
    $s = Get-RowScale $r
    if ([Math]::Abs($s - 1.0) -lt 1e-9) { return [double[]]$r }
    return [double[]]($r | ForEach-Object { $_ / $s })
}
# Yaw angle (degrees) of the local +X axis (row 0 of R), projected onto the
# world XY plane, measured from world +X.
function Get-YawAngleDeg([double[]]$r) {
    $rn = Normalize-Rot3x3 $r
    return [Math]::Atan2($rn[1], $rn[0]) * 180.0 / [Math]::PI
}
# Relative yaw (degrees) of $rSrc with respect to $rRefNorm, defined so that
# $rSrc = $rRefNorm * ZRotation(result). Found by computing
# M = Transpose($rRefNorm) * Normalize($rSrc), which collapses to a pure
# Z-rotation matrix when $rSrc and $rRefNorm share the same tilt (row2),
# and reading the angle off M's top-left 2x2 block. This works regardless
# of which local axis the object's tilt happens to point along.
function Get-RelativeYawDeg([double[]]$rSrc, [double[]]$rRefNorm) {
    $rs = Normalize-Rot3x3 $rSrc
    $m0 = $rRefNorm[0]*$rs[0] + $rRefNorm[3]*$rs[3] + $rRefNorm[6]*$rs[6]
    $m1 = $rRefNorm[0]*$rs[1] + $rRefNorm[3]*$rs[4] + $rRefNorm[6]*$rs[7]
    return [Math]::Atan2($m1, $m0) * 180.0 / [Math]::PI
}
# Row-vector rotation matrix for a rotation by $angleDeg about the world "up"
# (Z) axis only.
function Get-ZRotationMatrix([double]$angleDeg) {
    $rad = $angleDeg * [Math]::PI / 180.0
    $c = [Math]::Cos($rad); $s = [Math]::Sin($rad)
    return [double[]]($c, $s, 0, (-$s), $c, 0, 0, 0, 1)
}
# Snap a rotation matrix to the nearest axis-aligned rotation: each row is
# replaced by the signed world axis it points closest to, claiming the most
# decisive (row, axis) pairs first so every axis is used exactly once. A nest
# plate rotation is typically built from 90-degree flips (which a Bambu
# trim/cut bakes into the mesh) plus an extra non-90 in-plane "nest angle"
# (which the bake normalizes away). This recovers the 90-degree part; the
# leftover, Get-RelativeYawDeg(original, snapped), is the nest angle about
# the world up axis - the part that must be re-applied to the baked Final.
function Get-AxisSnappedRot([double[]]$r) {
    $rn = Normalize-Rot3x3 $r
    $snap = New-Object double[] 9
    $usedRow = @{}; $usedCol = @{}
    for ($k = 0; $k -lt 3; $k++) {
        $bestRow = -1; $bestCol = -1; $bestAbs = -1.0
        for ($row = 0; $row -lt 3; $row++) {
            if ($usedRow.ContainsKey($row)) { continue }
            for ($c = 0; $c -lt 3; $c++) {
                if ($usedCol.ContainsKey($c)) { continue }
                $a = [Math]::Abs($rn[$row*3+$c])
                if ($a -gt $bestAbs) { $bestAbs = $a; $bestRow = $row; $bestCol = $c }
            }
        }
        $usedRow[$bestRow] = $true; $usedCol[$bestCol] = $true
        $snap[$bestRow*3+$bestCol] = if ($rn[$bestRow*3+$bestCol] -ge 0) { 1.0 } else { -1.0 }
    }
    return $snap
}
# Yaw-only manual correction: every clone keeps $rRefNorm (normally the
# Final's own current tilt/orientation) and is additionally spun about the
# world "up" (Z) axis by $deltaYawDeg (the per-instance yaw difference between
# a Nest Source instance and the Nest Source's reference instance, plus any
# user-added "Rotate 90" steps). The spin is applied via right-multiplication
# (R_new = $rRefNorm * ZRotation(deltaYaw)), matching how the Nest Source's
# own per-instance rotations relate to each other - so the output's tilt is
# unchanged from $rRefNorm and only its world-Z orientation varies per clone.
# Translation handling is unchanged from Apply-TxCorrection.
function Apply-TxCorrectionYawOnly([double[]]$tx, [double[]]$rRefNorm, [double]$deltaYawDeg, [double[]]$tDelta, [double]$sRatio = 1.0) {
    $rSrc = Get-TxRot $tx
    $yawRot = Get-ZRotationMatrix $deltaYawDeg
    $rNew = Mul-3x3 $rRefNorm $yawRot
    if ([Math]::Abs($sRatio - 1.0) -gt 1e-6) {
        for ($ri = 0; $ri -lt 9; $ri++) { $rNew[$ri] *= $sRatio }
    }
    $tDeltaXY = [double[]]($tDelta[0], $tDelta[1], 0)
    $dRot = Rotate-Vec $tDeltaXY $rSrc
    $t0 = $tx[9]  + $dRot[0]
    $t1 = $tx[10] + $dRot[1]
    $t2 = $tx[11] + $tDelta[2]
    return [double[]](
        $rNew[0],$rNew[1],$rNew[2], $rNew[3],$rNew[4],$rNew[5], $rNew[6],$rNew[7],$rNew[8],
        $t0, $t1, $t2)
}

# Apply rotation correction (left-multiply), scale correction, and translation delta.
# R_out = sRatio * (rCorr * R_src)   where sRatio = scaleFinal / scaleSrc
# t_out = t_src + tDelta * R_src  (tDelta rotated by original source rotation)
function Apply-TxCorrection([double[]]$tx, [double[]]$rCorr, [double[]]$tDelta, [double]$sRatio = 1.0) {
    $rSrc = Get-TxRot $tx
    $rNew = Mul-3x3 $rCorr $rSrc
    if ([Math]::Abs($sRatio - 1.0) -gt 1e-6) {
        for ($ri = 0; $ri -lt 9; $ri++) { $rNew[$ri] *= $sRatio }
    }
    # Rotate the XY centroid correction through rSrc for X/Y output, but discard dRot[2].
    # When the plate rotation maps a local axis to world Z (e.g. flat colorcut 90-deg tilt),
    # any XY centroid offset bleeds into dRot[2] and corrupts world Z.  Z is corrected
    # separately and directly via tDelta[2] = finalBuildZ - srcMeanZ.
    $tDeltaXY = [double[]]($tDelta[0], $tDelta[1], 0)
    $dRot = Rotate-Vec $tDeltaXY $rSrc
    $t0 = $tx[9]  + $dRot[0]
    $t1 = $tx[10] + $dRot[1]
    $t2 = $tx[11] + $tDelta[2]
    return [double[]](
        $rNew[0],$rNew[1],$rNew[2], $rNew[3],$rNew[4],$rNew[5], $rNew[6],$rNew[7],$rNew[8],
        $t0, $t1, $t2)
}
# Extract the tilt-only component of a rotation matrix, stripping any Z-axis (yaw) rotation.
# Uses Rodrigues formula: finds shortest rotation that maps [0,0,1] to R*[0,0,1].
# Row-vector convention: [0,0,1] maps to R's third row = [R[6],R[7],R[8]].
# Because Z rotation leaves the Z direction unchanged, this always gives the tilt
# regardless of how much yaw the user may have added on top.
function Get-TiltOnlyCorrection([double[]]$r) {
    $vx = $r[6]; $vy = $r[7]; $vz = $r[8]
    $vlen = [Math]::Sqrt($vx*$vx + $vy*$vy + $vz*$vz)
    if ($vlen -gt 1e-9) { $vx /= $vlen; $vy /= $vlen; $vz /= $vlen }

    # Z already points straight up - no tilt, return identity
    if ([Math]::Abs($vx) -lt 1e-6 -and [Math]::Abs($vy) -lt 1e-6 -and $vz -gt 0.9999) {
        return [double[]](1,0,0, 0,1,0, 0,0,1)
    }
    # Z points straight down - 180 deg flip around X
    if ([Math]::Abs($vx) -lt 1e-6 -and [Math]::Abs($vy) -lt 1e-6 -and $vz -lt -0.9999) {
        return [double[]](1,0,0, 0,-1,0, 0,0,-1)
    }

    # Rodrigues axis = cross([0,0,1], v) = [-vy, vx, 0]
    $n0 = -$vy; $n1 = $vx
    $nlen = [Math]::Sqrt($n0*$n0 + $n1*$n1)
    $n0 /= $nlen; $n1 /= $nlen

    $cosA = [Math]::Max(-1.0, [Math]::Min(1.0, $vz))
    $sinA = [Math]::Sqrt(1.0 - $cosA*$cosA)
    $t    = 1.0 - $cosA

    # Row-vector Rodrigues matrix = R_col^T
    $m00 = $cosA + $t*$n0*$n0;   $m01 = $t*$n0*$n1;          $m02 = -$sinA*$n1
    $m10 = $t*$n0*$n1;            $m11 = $cosA + $t*$n1*$n1;  $m12 =  $sinA*$n0
    $m20 = $sinA*$n1;             $m21 = -$sinA*$n0;           $m22 =  $cosA
    return [double[]]($m00,$m01,$m02, $m10,$m11,$m12, $m20,$m21,$m22)
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

# â”€â”€ Resolve input paths â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
$FinalPath = $FinalPath.Trim('"')
if (-not (Test-Path $FinalPath)) { Write-Error "Final file not found: $FinalPath"; exit 1 }

$finalDir  = Split-Path $FinalPath -Parent
$finalStem = [System.IO.Path]::GetFileNameWithoutExtension($FinalPath) -replace '(?i)_Final$', ''

# Auto-detect transform source:
#   If both Nest and Full exist alongside Final â†’ use Nest (file has been merged)
#   If only Full exists â†’ use Full
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

# â”€â”€ Confirm before proceeding â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
Write-Host ""
Write-Host "â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”"
Write-Host "  RenestFromFinal - Planned Operation"
Write-Host "â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”"
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
if (-not $NoConfirm) {
    $confirm = (Read-Host "  Proceed? [Y/N]").Trim()
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Cancelled."
        exit 0
    }
} else {
    Write-Host "  (NoConfirm - proceeding automatically)"
}
Write-Host ""

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  STEP 1 â€” Read transforms from the source plate file
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

# Grab the component transforms from the source template assembly.
$srcTemplateCompTransforms = [System.Collections.Generic.List[string]]::new()
$srcTemplateObj = $srcModel.SelectSingleNode('//m:resources/m:object[m:components/m:component]', $srcXns)
if ($null -ne $srcTemplateObj) {
    foreach ($srcComp in $srcTemplateObj.SelectNodes('m:components/m:component', $srcXns)) {
        $srcTemplateCompTransforms.Add($srcComp.GetAttribute('transform')) | Out-Null
    }
}
# Mean Z translation across all source build items (used to compute Z correction).
$srcMeanZ = 0.0
foreach ($st in $sourceTransforms) { $srcMeanZ += (Parse-Tx $st)[11] }
if ($sourceTransforms.Count -gt 0) { $srcMeanZ /= $sourceTransforms.Count }

# Extract source plate metadata (preserves filament_maps, filament_volume_maps, etc.)
$srcPlateMeta = [ordered]@{}
if ($null -ne $srcSettingsText) {
    [xml]$srcSettingsForMeta = $srcSettingsText
    $srcPlateNode = $srcSettingsForMeta.SelectSingleNode('//plate')
    if ($null -ne $srcPlateNode) {
        foreach ($m in $srcPlateNode.SelectNodes('metadata')) {
            $srcPlateMeta[$m.GetAttribute('key')] = $m.GetAttribute('value')
        }
    }
}

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  STEP 2 â€” Extract Final.3mf to working directory
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

    # â”€â”€ Identify the master (printable) assembly object â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€ Collect Final master component transforms (for debug output) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $masterCompTransforms = [System.Collections.Generic.List[string]]::new()
    foreach ($mc in $masterObj.SelectNodes('m:components/m:component', $xns)) {
        $masterCompTransforms.Add($mc.GetAttribute('transform')) | Out-Null
    }

    # â”€â”€ Derive rotation correction from the Final master's build item transform â”€â”€â”€â”€â”€
    # The Final's build item rotation captures how Bambu oriented the assembly on
    # the editing plate relative to the "natural" local frame of the geometry.
    # Source plate transforms assume the original (Nest/Full) orientation.
    # Left-multiplying each source rotation by R_final_build re-aligns the clones
    # so they sit on the plate the same way the source instances did.
    # If the Final's build item is identity (no reorientation), no correction is needed.
    $finalBuildRot = Get-TxRot (Parse-Tx $masterItem.GetAttribute('transform'))
    $compsMatch    = $false   # initialised here; set inside else block if applicable
    # Variables initialised here so the debug section can always reference them
    $rotMatchesSource = $false
    $sameTilt         = $false
    $refR8norm        = 1.0
    $finalR8norm      = 1.0
    $trimBakedDetected = $false   # set true when trim-bake path is taken

    # â”€â”€ Manual orientation correction override â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # If the user has run the CardQueueEditor "Fix Orientation" picker and saved
    # a <stem>_RotCorrection.json, that selection is the user's confirmed ground
    # truth for how the Final relates to the source orientation. Skip ALL of the
    # automatic detection heuristics below (trim-bake detection, tilt-depth
    # check, etc.) - they are best-effort guesses and can misfire (e.g. detecting
    # a "trim bake" and cancelling out a legitimate diagonal nest rotation) when
    # the user has already verified the correct orientation by hand. Start from
    # identity; the manual correction is layered on afterwards.
    $rotCorrJsonPath = if (-not [string]::IsNullOrWhiteSpace($RotCorrectionPath)) { $RotCorrectionPath } else { Join-Path $finalDir "${finalStem}_RotCorrection.json" }
    if (Test-Path $rotCorrJsonPath) {
        $rotCorrection = [double[]](1,0,0, 0,1,0, 0,0,1)
        Write-Host "Manual orientation correction present - skipping automatic rotation-correction detection."
    } elseif (Is-IdentityRot $finalBuildRot) {
        # â”€â”€ Trim-bake detection â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # When Bambu Studio trims/cuts an object it bakes the current plate rotation
        # into the mesh vertex positions and resets the build item transform to
        # identity.  The resulting Final looks identical but all rotation metadata
        # is gone.  If we then apply the source transforms (which still carry the
        # original plate rotation R) to this pre-rotated geometry, every clone
        # double-rotates by R.
        #
        # Detection heuristic: cut_information.xml is present (Bambu writes it for
        # every trim/cut operation) AND at least one source build item has a
        # non-identity rotation.
        #
        # Correction: R_correction = R_src^T = R_src^-1 (for the normalised, pure
        # rotation part).  Apply-TxCorrection then computes:
        #   R_new = R_src^T * R_src = I  â†’  pure translation, no rotation added.
        # The pre-baked mesh therefore lands at each plate position correctly.
        $firstSrcRotNorm  = $null
        $srcHasRotation   = $false
        foreach ($txStr in $sourceTransforms) {
            if ([string]::IsNullOrWhiteSpace($txStr)) { continue }
            $r = Get-TxRot (Parse-Tx $txStr)
            if (-not (Is-IdentityRot $r)) {
                # Strip uniform scale so we get a pure rotation matrix for transposition
                $rsc = Get-RowScale $r
                $firstSrcRotNorm = if ($rsc -gt 1e-9) { [double[]]($r | ForEach-Object { $_ / $rsc }) } else { [double[]]$r }
                $srcHasRotation = $true
                break
            }
        }
        $hasCutInfo = ($null -ne $cutInfoPath) -and (Test-Path $cutInfoPath)

        if ($hasCutInfo -and $srcHasRotation) {
            # The trim baked the source's plate rotation into the mesh vertices -
            # but only its 90-degree-aligned part: the bake normalizes away any
            # extra non-90 in-plane "nest angle" the plate rotation carried.
            # Inverting the FULL source rotation (R_src^T * R_src = I) would
            # therefore also cancel that nest angle, leaving every clone
            # axis-aligned instead of angled like the nest. Invert only the
            # axis-snapped (90-degree) part:
            #   R_new_i = R_snap^T * R_src_i = ZRotation(nestAngle + deltaYaw_i)
            # so each clone keeps its source instance's full in-plane angle
            # (a pure world-Z rotation on the baked geometry).
            $srcSnapRot        = Get-AxisSnappedRot $firstSrcRotNorm
            $srcNestAngleDeg   = Get-RelativeYawDeg $firstSrcRotNorm $srcSnapRot
            $rotCorrection     = Transpose-3x3 $srcSnapRot
            $trimBakedDetected = $true
            Write-Host ("Trim-baked geometry detected (cut_information.xml present, source has rotation, Final is identity).")
            Write-Host ("Correction: inverse of axis-snapped source rotation applied (preserves nest angle {0:F2} deg about world Z)." -f $srcNestAngleDeg)
        } else {
            $rotCorrection = [double[]](1,0,0, 0,1,0, 0,0,1)
            Write-Host "Final master build item is identity - no rotation correction needed."
        }
    } else {
        # Check if the Final's build rotation matches any source plate item rotation exactly.
        # If the user exported the Final directly from the nest its rotation will appear in
        # the source build items; geometry is already in the expected frame so no correction
        # is needed (applying one would double-rotate every clone).
        $rotEps = 1e-3
        foreach ($txStr in $sourceTransforms) {
            if ([string]::IsNullOrWhiteSpace($txStr)) { continue }
            $srcItemRot = Get-TxRot (Parse-Tx $txStr)
            $match = $true
            for ($ri = 0; $ri -lt 9 -and $match; $ri++) {
                if ([Math]::Abs($finalBuildRot[$ri] - $srcItemRot[$ri]) -gt $rotEps) { $match = $false }
            }
            if ($match) { $rotMatchesSource = $true; break }
        }

        if ($rotMatchesSource) {
            $rotCorrection = [double[]](1,0,0, 0,1,0, 0,0,1)
            Write-Host "Final rotation matches a source plate item - same orientation, no correction needed."
        } else {
            # â”€â”€ Tilt-depth check â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            # r[8] / |row| = cos(tilt_angle) and is INVARIANT under any additional
            # Z-axis (yaw) rotation.  If the user opened the Final (already at the
            # source's tilt) and simply spun it around Z for editing convenience,
            # the tilt depth will be identical to the source reference even though
            # the full rotation matrix no longer matches.  Applying Get-TiltOnlyCorrection
            # in that case would extract a DIFFERENT tilt direction (the source tilt
            # rotated by the extra yaw) and double-tilt every clone on the plate.
            # Solution: skip the correction whenever the tilt depth agrees with the
            # source.  Only apply a tilt correction when the user genuinely changed
            # the tilt (e.g. stood the object upright when the source has it flat).
            $sumR8 = 0.0; $cntR8 = 0
            foreach ($txStr in $sourceTransforms) {
                if ([string]::IsNullOrWhiteSpace($txStr)) { continue }
                $rt = Get-TxRot (Parse-Tx $txStr)
                $rLen = [Math]::Sqrt($rt[6]*$rt[6]+$rt[7]*$rt[7]+$rt[8]*$rt[8])
                if ($rLen -gt 1e-9) { $sumR8 += $rt[8]/$rLen; $cntR8++ }
            }
            $refR8norm   = if ($cntR8 -gt 0) { $sumR8/$cntR8 } else { 1.0 }
            $finalRowLen = [Math]::Sqrt($finalBuildRot[6]*$finalBuildRot[6]+$finalBuildRot[7]*$finalBuildRot[7]+$finalBuildRot[8]*$finalBuildRot[8])
            $finalR8norm = if ($finalRowLen -gt 1e-9) { $finalBuildRot[8]/$finalRowLen } else { 1.0 }
            $sameTilt    = ([Math]::Abs($finalR8norm - $refR8norm) -lt 0.02)

            if ($sameTilt) {
                $rotCorrection = [double[]](1,0,0, 0,1,0, 0,0,1)
                Write-Host ("Tilt depth matches source (final={0:F4} src_ref={1:F4}) - Z-yaw only, no tilt correction." -f $finalR8norm, $refR8norm)
            } else {
                # Tilt genuinely differs - extract tilt-only correction, strip Z yaw so
                # the nest layout positions don't spin.
                $rotCorrection = Get-TiltOnlyCorrection $finalBuildRot
                if (Is-IdentityRot $rotCorrection) {
                    Write-Host ("Tilt differs (final={0:F4} src_ref={1:F4}) but correction is identity (upright model)." -f $finalR8norm, $refR8norm)
                } else {
                    Write-Host ("Tilt differs (final={0:F4} src_ref={1:F4}) - tilt-only correction applied (Z yaw stripped)." -f $finalR8norm, $refR8norm)
                }
            }
        }
    }

    # â”€â”€ Apply manual orientation correction from the CardQueueEditor "Fix Orientation"
    # picker, if present. This is layered on top of whatever automatic correction was
    # derived above (identity if none).
    $manualCorrApplied  = $false
    $manualCorrDesc     = ""
    $manualYawMode      = $false
    $manualSrcRefRotNorm = $null
    $manualFinalRotNorm  = $null
    $manualSrcRefYawDeg = 0.0
    $manualExtraYawDeg  = 0.0
    $manualRefSpinDeg   = 0.0
    if (Test-Path $rotCorrJsonPath) {
        try {
            $rotCorrJson = Get-Content -LiteralPath $rotCorrJsonPath -Raw | ConvertFrom-Json
            if ($null -ne $rotCorrJson.srcRefRot) {
                # Yaw-only manual correction: every clone KEEPS the Final's own
                # current tilt/orientation (R_finalNorm) - it is rotated about
                # the world "up" (Z) axis only. The per-instance Z-rotation is:
                #   refSpin    - the user-confirmed total spin for the
                #                reference position (the non-90 "nest angle"
                #                a Bambu trim/cut bake normalized away, plus
                #                the user's "Rotate 90" steps), exactly as
                #                previewed in the Fix Orientation picker;
                #   deltaYaw_i - how much each Nest Source instance is rotated
                #                (about world Z) relative to the reference.
                # srcRefRot is used ONLY to measure deltaYaw_i - its tilt is
                # never adopted by the output.
                $manualYawMode       = $true
                $manualSrcRefRotNorm = Normalize-Rot3x3 ([double[]]($rotCorrJson.srcRefRot))
                $manualFinalRotNorm  = Normalize-Rot3x3 $finalBuildRot
                $manualSrcRefYawDeg  = [double]$rotCorrJson.srcRefYawDeg
                $manualExtraYawDeg   = [double]$rotCorrJson.extraYawDeg
                if ($null -ne $rotCorrJson.refSpinDeg) {
                    # refSpinDeg is the user-CONFIRMED total reference spin,
                    # exactly as previewed in the picker. Apply it verbatim -
                    # do NOT re-derive the nest angle by axis-snapping here,
                    # or the applied spin could silently differ from what the
                    # user saw and confirmed.
                    $manualRefSpinDeg = [double]$rotCorrJson.refSpinDeg
                } else {
                    # Older JSON without refSpinDeg: reconstruct it the same
                    # way the picker did (snap-derived nest angle + extra yaw).
                    $manualRefSpinDeg = (Get-RelativeYawDeg $manualSrcRefRotNorm (Get-AxisSnappedRot $manualSrcRefRotNorm)) + $manualExtraYawDeg
                }
                $rotCorrection       = $manualFinalRotNorm
                $manualCorrApplied   = $true
                $manualCorrDesc = ("yaw-only about world Z, Final's own tilt kept: refSpin={0:F2} (srcRefYaw={1:F1} extraYaw={2:F1})" -f $manualRefSpinDeg, $manualSrcRefYawDeg, $manualExtraYawDeg)
                Write-Host "Manual orientation correction loaded from $(Split-Path $rotCorrJsonPath -Leaf): $manualCorrDesc"
            } elseif ($null -ne $rotCorrJson.rotCorrection) {
                $manualCorr = [double[]]($rotCorrJson.rotCorrection)
                if ($manualCorr.Count -eq 9) {
                    $rotCorrection = Mul-3x3 $manualCorr $rotCorrection
                    $manualCorrApplied = $true
                    $manualCorrDesc = ("axis {0} -> {1} ({2:F1} deg)" -f $rotCorrJson.selection.axis, $rotCorrJson.selection.userDir, [double]$rotCorrJson.angleDegrees)
                    Write-Host "Manual orientation correction applied from $(Split-Path $rotCorrJsonPath -Leaf): $manualCorrDesc"
                }
            }
        } catch {
            Write-Host "WARNING: failed to read/apply $(Split-Path $rotCorrJsonPath -Leaf): $($_.Exception.Message)"
        }
    }

    # â”€â”€ Extract scale from Final build item â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Bambu embeds the model's print scale in the build item rotation matrix.
    # If the user scaled the model in the Final, scaleFinal != scaleSrc, and each
    # clone's rotation matrix must be rescaled so geometry and Z position match.
    $scaleFinal = Get-RowScale $finalBuildRot
    Write-Host ("Scale from Final build item: {0:F4}" -f $scaleFinal)

    # â”€â”€ Compute centroid-based translation delta â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # When Bambu re-saves an edited Final it re-centers the geometry (centroid moves
    # to near the local origin), but the source plate translations assume the original
    # (un-centered) local frame.  delta = centroid_src - centroid_final * rotCorrection
    # is added to each clone translation (rotated by R_src) to compensate.
    #
    # IMPORTANT: this only works when BOTH the source template AND the Final have
    # identity-rotation components so their centroids are in the same coordinate frame.
    # For colorcut designs the Final's components all carry a ~90-degree rotation, making
    # Get-CompCentroid return (0,0,0).  Comparing that against the Nest centroid (which IS
    # in identity-rotation frame) produces a bogus delta that bleeds into world Z via the
    # plate rotation.  When the Final has no identity-rotation components, skip the
    # centroid correction entirely (zero both sides).
    # The clone XY translations are taken VERBATIM from the source plate build
    # items - no computed offset. The Final's local frame is the same frame
    # the source template used (Bambu keeps the object origin "locator"
    # stable through edits), so the source translations are already correct;
    # any content-derived correction (centroid means etc.) only lets pieces
    # the user added, moved, or duplicated drag the whole nest around.
    #
    # As a SANITY CHECK we still measure the frame offset by matching the
    # components shared by both objects (in the source template's local frame
    # - in manual yaw mode $rotCorrection is not a frame mapping, so use the
    # transpose of the axis-snapped source reference rotation, i.e. the bake
    # rotation, instead). If the shared components disagree with the source
    # by more than 1 mm the Final's frame genuinely moved - unusual - and we
    # warn, but still trust the Bambu transforms.
    $matchFrameRot = if ($manualYawMode) { Transpose-3x3 (Get-AxisSnappedRot $manualSrcRefRotNorm) } else { $rotCorrection }
    $matchedOffset = Get-MatchedFrameOffset $srcTemplateCompTransforms $masterCompTransforms $matchFrameRot
    if ($null -ne $matchedOffset) {
        $frameCheckDesc = ("frame check: offset=({0:F2},{1:F2},{2:F2}) from {3} matched components" -f $matchedOffset[0], $matchedOffset[1], $matchedOffset[2], [int]$matchedOffset[3])
        if ([Math]::Abs($matchedOffset[0]) -gt 1.0 -or [Math]::Abs($matchedOffset[1]) -gt 1.0 -or [Math]::Abs($matchedOffset[2]) -gt 1.0) {
            Write-Host ("WARNING: the Final's local frame appears shifted vs the source template ({0}). Clones may land off-position." -f $frameCheckDesc)
        }
    } else {
        $frameCheckDesc = "frame check: no dominant component match (heavily edited Final) - trusting Bambu transforms"
    }
    # Z: a cut/edit can change the bottom face, and Bambu re-seats the Final
    # on the plate. Clones must use the Final's own build-item Z (Bambu's
    # actual seating for the edited geometry), not the source's.
    $finalBuildZ = (Parse-Tx $masterItem.GetAttribute('transform'))[11]
    $td2 = $finalBuildZ - $srcMeanZ
    $tDelta = [double[]](0.0, 0.0, $td2)
    Write-Host ("Translations taken verbatim from source plate ({0})." -f $frameCheckDesc)
    Write-Host ("  Z seating: finalBuildZ={0:F4} srcMeanZ={1:F4} td2={2:F4}" -f $finalBuildZ,$srcMeanZ,$td2)

    # â”€â”€ Snapshot component UUIDs base from the master â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # We'll regenerate UUIDs per-clone to avoid duplicates
    $masterComps = @($masterObj.SelectNodes('m:components/m:component', $xns))

    # â”€â”€ Remove all existing build items â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€ Find the highest ID in any external object file (e.g. object_1.model) â”€â”€
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

    # â”€â”€ Find the highest internal-mesh ID to start our new IDs above it â”€â”€â”€â”€â”€â”€â”€â”€
    $nextId = $maxExternalId + 1
    foreach ($id in $objById.Keys) {
        if ($printableIdsInFinal.Contains($id)) { continue }
        $v = 0; if ([int]::TryParse($id, [ref]$v) -and $v -ge $nextId) { $nextId = $v + 1 }
    }

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  STEP 3 â€” Clone master object N times, one per source transform
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

        # Build item with corrected transform (rotation + centroid offset)
        $newItem = $xml.CreateElement('item', $nsCore)
        $newItem.SetAttribute('objectid', $newId)
        $tx = $sourceTransforms[$i]
        if (-not [string]::IsNullOrWhiteSpace($tx)) {
            $srcScale = Get-RowScale (Get-TxRot (Parse-Tx $tx))
            $sRatio   = if ($srcScale -gt 1e-9) { $scaleFinal / $srcScale } else { 1.0 }
            if ($manualYawMode) {
                $deltaYaw  = (Get-RelativeYawDeg (Get-TxRot (Parse-Tx $tx)) $manualSrcRefRotNorm) + $manualRefSpinDeg
                $corrected = Apply-TxCorrectionYawOnly (Parse-Tx $tx) $manualFinalRotNorm $deltaYaw $tDelta $sRatio
            } else {
                $corrected = Apply-TxCorrection (Parse-Tx $tx) $rotCorrection $tDelta $sRatio
            }
            $tx = ($corrected | ForEach-Object { ([double]$_).ToString('G9') }) -join ' '
            $newItem.SetAttribute('transform', $tx)
        }
        $newItem.SetAttribute('UUID', $nsProd, ($i.ToString("x8") + "-b1ec-4553-aec9-835e5b724bb4")) | Out-Null
        $newItem.SetAttribute('printable', '1')
        $buildNode.AppendChild($newItem) | Out-Null
    }

    Write-Host "Cloned master object x$n in 3dmodel.model."

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  STEP 4 â€” Rebuild model_settings.config
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

        # Copy plate metadata from the source Nest/Full settings (preserves filament_maps,
        # filament_volume_maps, filament_map_mode, locked, etc.).
        # Override thumbnail/id keys with standard plate-1 names.
        $newPlate = $settings.CreateElement('plate')
        $overridePlateMeta = [ordered]@{
            'plater_id'              = '1'
            'thumbnail_file'         = 'Metadata/plate_1.png'
            'thumbnail_no_light_file'= 'Metadata/plate_no_light_1.png'
            'top_file'               = 'Metadata/top_1.png'
            'pick_file'              = 'Metadata/pick_1.png'
        }
        # Merge: start from source, then apply overrides
        $mergedPlateMeta = [ordered]@{}
        foreach ($kv in $srcPlateMeta.GetEnumerator()) { $mergedPlateMeta[$kv.Key] = $kv.Value }
        foreach ($kv in $overridePlateMeta.GetEnumerator()) { $mergedPlateMeta[$kv.Key] = $kv.Value }
        # Ensure plater_id is first
        $orderedPlateMeta = [ordered]@{ 'plater_id' = '1' }
        foreach ($kv in $mergedPlateMeta.GetEnumerator()) {
            if ($kv.Key -ne 'plater_id') { $orderedPlateMeta[$kv.Key] = $kv.Value }
        }
        foreach ($kv in $orderedPlateMeta.GetEnumerator()) {
            $m = $settings.CreateElement('metadata')
            $m.SetAttribute('key',   $kv.Key)
            $m.SetAttribute('value', $kv.Value)
            $newPlate.AppendChild($m) | Out-Null
        }
        $configNode.AppendChild($newPlate) | Out-Null

        for ($i = 0; $i -lt $n; $i++) {
            $newId = $newObjIds[$i]
            $srcScale = Get-RowScale (Get-TxRot (Parse-Tx $sourceTransforms[$i]))
            $sRatio   = if ($srcScale -gt 1e-9) { $scaleFinal / $srcScale } else { 1.0 }
            if ($manualYawMode) {
                $deltaYaw  = (Get-RelativeYawDeg (Get-TxRot (Parse-Tx $sourceTransforms[$i])) $manualSrcRefRotNorm) + $manualRefSpinDeg
                $corrected = Apply-TxCorrectionYawOnly (Parse-Tx $sourceTransforms[$i]) $manualFinalRotNorm $deltaYaw $tDelta $sRatio
            } else {
                $corrected = Apply-TxCorrection (Parse-Tx $sourceTransforms[$i]) $rotCorrection $tDelta $sRatio
            }
            $tx = ($corrected | ForEach-Object { ([double]$_).ToString('G9') }) -join ' '

            # Clone master settings entry with new ID
            if ($null -ne $masterSettObj) {
                $clonedSett = $settings.ImportNode($masterSettObj, $true)
                $clonedSett.SetAttribute('id', $newId)
                $configNode.InsertBefore($clonedSett, $newAssemble) | Out-Null
            }

            # assemble_item â€” match Bambu's format (offset not modelmesh_id)
            $asmItem = $settings.CreateElement('assemble_item')
            $asmItem.SetAttribute('object_id',   $newId)
            $asmItem.SetAttribute('instance_id', '0')
            $asmItem.SetAttribute('transform',   $tx)
            $asmItem.SetAttribute('offset',      '0 0 0')
            $newAssemble.AppendChild($asmItem) | Out-Null

            # model_instance in plate â€” object_id, instance_id, identify_id (no 'centered')
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

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  STEP 5 â€” Global ID renumbering (internal meshes first, assemblies after)
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  STEP 6 â€” Rebuild VLH, purge stale thumbnails
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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
    foreach ($pat in @("plate_*.png","pick_*.png")) {
        Get-ChildItem $metaDir -Filter $pat -ErrorAction SilentlyContinue | Remove-Item -Force
    }

    # Read plate_1.json and wipe tower settings from the source (Nest/Full) zip into memory.
    # All are re-injected after the Bambu resave step, which strips plate_1.json and
    # overwrites project_settings.config with the Final's (wrong) wipe tower settings.
    $plate1JsonBytes      = $null
    $srcWipeTowerX        = $null
    $srcWipeTowerY        = $null
    $srcPrimeTowerWidth   = $null
    $srcZipForPlate = [System.IO.Compression.ZipFile]::OpenRead($TransformSourcePath)
    try {
        # plate_1.json
        $plate1Entry = $srcZipForPlate.Entries | Where-Object { $_.FullName -eq 'Metadata/plate_1.json' } | Select-Object -First 1
        if ($null -ne $plate1Entry) {
            $ms = New-Object System.IO.MemoryStream
            $plate1Stream = $plate1Entry.Open()
            try { $plate1Stream.CopyTo($ms) } finally { $plate1Stream.Dispose() }
            $plate1JsonBytes = $ms.ToArray()
            $ms.Dispose()
            Write-Host "plate_1.json read from source ($($plate1JsonBytes.Length) bytes)."
        } else {
            Write-Host "WARNING: plate_1.json not found in source - wipe tower position may be wrong."
        }

        # wipe_tower_x / wipe_tower_y from project_settings.config
        $projEntry = $srcZipForPlate.Entries | Where-Object { $_.FullName -eq 'Metadata/project_settings.config' } | Select-Object -First 1
        if ($null -ne $projEntry) {
            $srProj = New-Object System.IO.StreamReader($projEntry.Open())
            $projText = $srProj.ReadToEnd(); $srProj.Dispose()
            $xm  = [regex]::Match($projText, '"wipe_tower_x"\s*:\s*\[([^\]]*)\]')
            $ym  = [regex]::Match($projText, '"wipe_tower_y"\s*:\s*\[([^\]]*)\]')
            $ptw = [regex]::Match($projText, '"prime_tower_width"\s*:\s*"([^"]*)"')
            if ($xm.Success)  { $srcWipeTowerX      = $xm.Groups[1].Value.Trim() }
            if ($ym.Success)  { $srcWipeTowerY      = $ym.Groups[1].Value.Trim() }
            if ($ptw.Success) { $srcPrimeTowerWidth = $ptw.Groups[1].Value.Trim() }
            if ($null -ne $srcWipeTowerX) {
                Write-Host "Wipe tower from source: x=[$srcWipeTowerX] y=[$srcWipeTowerY] prime_tower_width=[$srcPrimeTowerWidth]"
            }
        }
    } finally { $srcZipForPlate.Dispose() }

    # Also write to workdir so it lands in the initial pack (Step 7).
    # Step 8 may strip it; the re-inject below handles that case.
    if ($null -ne $plate1JsonBytes) {
        [System.IO.File]::WriteAllBytes((Join-Path $metaDir "plate_1.json"), $plate1JsonBytes)
    }

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  STEP 7 â€” Repack
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  DEBUG - Write transform debug file alongside output
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    $debugPath = $OutputPath -replace '\.3mf$', '_debug.txt'
    $dbLines   = [System.Collections.Generic.List[string]]::new()
    $dbLines.Add("RenestFromFinal Debug Output")
    $dbLines.Add("Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $dbLines.Add("Final     : $FinalPath")
    $dbLines.Add("Source    : $TransformSourcePath")
    $dbLines.Add("Output    : $OutputPath")
    $dbLines.Add("")

    $dbLines.Add("======================================")
    $dbLines.Add("FINAL.3MF - Master Build Item")
    $dbLines.Add("======================================")
    $dbLines.Add("  Object ID   : $masterId")
    $masterBuildTx = $masterItem.GetAttribute('transform')
    $dbLines.Add("  Build TX    : $(if ([string]::IsNullOrWhiteSpace($masterBuildTx)) { '(identity/none)' } else { $masterBuildTx })")
    $dbLines.Add("  Build ROT   : $($finalBuildRot -join ' ')")
    $dbLines.Add("  Rot match   : $(if ($rotMatchesSource) { 'YES - exact match to a source item, no correction' } else { 'NO - no exact match' })")
    $dbLines.Add("  Tilt check  : final r8/|r|=$($finalR8norm.ToString('F4'))  src_ref=$($refR8norm.ToString('F4'))  sameTilt=$sameTilt")
    $dbLines.Add("  Trim baked  : $trimBakedDetected")
    $corrDesc = if ($trimBakedDetected) { 'inverse source rotation (trim-bake cancellation)' } `
                elseif (Is-IdentityRot $rotCorrection) { 'none (identity)' } `
                else { 'applied (tilt-only, Z yaw stripped)' }
    $dbLines.Add("  Correction  : $corrDesc")
    $dbLines.Add("  Manual fix  : $(if ($manualCorrApplied) { "applied ($manualCorrDesc) from $(Split-Path $rotCorrJsonPath -Leaf)" } else { 'none' })")
    $dbLines.Add("")

    $dbLines.Add("FINAL.3MF - Master Object Component Transforms")
    $dbLines.Add("  ($($masterCompTransforms.Count) components)")
    for ($ci = 0; $ci -lt $masterCompTransforms.Count; $ci++) {
        $v = $masterCompTransforms[$ci]
        $dbLines.Add("  comp[$ci] TX : $(if ([string]::IsNullOrWhiteSpace($v)) { '(identity/none)' } else { $v })")
    }
    $dbLines.Add("")

    $dbLines.Add("SOURCE - Template Object Component Transforms")
    $dbLines.Add("  ($($srcTemplateCompTransforms.Count) components)")
    for ($ci = 0; $ci -lt $srcTemplateCompTransforms.Count; $ci++) {
        $v = $srcTemplateCompTransforms[$ci]
        $dbLines.Add("  comp[$ci] TX : $(if ([string]::IsNullOrWhiteSpace($v)) { '(identity/none)' } else { $v })")
    }
    $dbLines.Add("")

    $dbLines.Add("======================================")
    $dbLines.Add("TRANSLATION HANDLING")
    $dbLines.Add("======================================")
    $dbLines.Add("  XY             : verbatim from source plate build items (no computed offset)")
    $dbLines.Add("  $frameCheckDesc")
    $dbLines.Add("  tDelta         : $($tDelta[0].ToString('F3')) $($tDelta[1].ToString('F3')) $($tDelta[2].ToString('F3'))  (Z seating only)")
    $dbLines.Add("")
    $dbLines.Add("======================================")
    $dbLines.Add("SOURCE vs APPLIED TRANSFORMS  ($n total, first 5 shown)")
    $dbLines.Add("======================================")
    $maxShow = [Math]::Min($n, 5)
    for ($i = 0; $i -lt $maxShow; $i++) {
        $v = $sourceTransforms[$i]
        $dbLines.Add("  [$i] src : $(if ([string]::IsNullOrWhiteSpace($v)) { '(identity/none)' } else { $v })")
        if (-not [string]::IsNullOrWhiteSpace($v)) {
            $dbSrcScale = Get-RowScale (Get-TxRot (Parse-Tx $v))
            $dbSRatio   = if ($dbSrcScale -gt 1e-9) { $scaleFinal / $dbSrcScale } else { 1.0 }
            if ($manualYawMode) {
                $dbDeltaYaw = (Get-RelativeYawDeg (Get-TxRot (Parse-Tx $v)) $manualSrcRefRotNorm) + $manualRefSpinDeg
                $cv = Apply-TxCorrectionYawOnly (Parse-Tx $v) $manualFinalRotNorm $dbDeltaYaw $tDelta $dbSRatio
            } else {
                $cv = Apply-TxCorrection (Parse-Tx $v) $rotCorrection $tDelta $dbSRatio
            }
            $cvStr = ($cv | ForEach-Object { ([double]$_).ToString('G9') }) -join ' '
            $dbLines.Add("  [$i] out : $cvStr")
        }
    }
    if ($n -gt $maxShow) { $dbLines.Add("  ... ($($n - $maxShow) more not shown)") }

    [System.IO.File]::WriteAllLines($debugPath, $dbLines, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Debug file : $debugPath"

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  STEP 8 â€” Optional Bambu resave (cleans stale metadata, generates thumbnails)
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

    # Re-inject plate_1.json and fix wipe tower settings after pack/resave.
    # Bambu resave strips plate_1.json and overwrites project_settings.config with
    # the Final's wipe tower settings; we restore all of them from the source here.
    if ($null -ne $plate1JsonBytes -or $null -ne $srcWipeTowerX) {
        $outZip = [System.IO.Compression.ZipFile]::Open($OutputPath, 'Update')
        try {
            # Inject plate_1.json
            if ($null -ne $plate1JsonBytes) {
                $existing = $outZip.Entries | Where-Object { $_.FullName -eq 'Metadata/plate_1.json' } | Select-Object -First 1
                if ($null -ne $existing) { $existing.Delete() }
                $newEntry    = $outZip.CreateEntry('Metadata/plate_1.json')
                $entryStream = $newEntry.Open()
                try { $entryStream.Write($plate1JsonBytes, 0, $plate1JsonBytes.Length) }
                finally { $entryStream.Dispose() }
                Write-Host "plate_1.json injected into output (wipe tower bbox preserved)."
            }

            # Patch wipe tower position + prime tower width in project_settings.config
            if ($null -ne $srcWipeTowerX) {
                $projE = $outZip.Entries | Where-Object { $_.FullName -eq 'Metadata/project_settings.config' } | Select-Object -First 1
                if ($null -ne $projE) {
                    $srOut = New-Object System.IO.StreamReader($projE.Open())
                    $projOut = $srOut.ReadToEnd(); $srOut.Dispose()
                    $projOut = [regex]::Replace($projOut, '("wipe_tower_x"\s*:\s*\[)[^\]]*(\])', "`$1$srcWipeTowerX`$2")
                    $projOut = [regex]::Replace($projOut, '("wipe_tower_y"\s*:\s*\[)[^\]]*(\])', "`$1$srcWipeTowerY`$2")
                    if ($null -ne $srcPrimeTowerWidth) {
                        # prime_tower_width is a plain string value: "prime_tower_width": "18"
                        # Use lookbehind so no backreference digits can collide with the numeric value
                        $projOut = [regex]::Replace($projOut, '(?<="prime_tower_width"\s*:\s*")[^"]*', $srcPrimeTowerWidth)
                    }
                    $projE.Delete()
                    $newProj     = $outZip.CreateEntry('Metadata/project_settings.config')
                    $projStream  = $newProj.Open()
                    $projBytes   = [System.Text.Encoding]::UTF8.GetBytes($projOut)
                    try { $projStream.Write($projBytes, 0, $projBytes.Length) }
                    finally { $projStream.Dispose() }
                    Write-Host "project_settings.config patched: wipe_tower x=[$srcWipeTowerX] y=[$srcWipeTowerY] prime_tower_width=[$srcPrimeTowerWidth]"
                }
            }
        } finally { $outZip.Dispose() }
    }

    Write-Host "Done! $n instances written to: $OutputPath"

} finally {
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
