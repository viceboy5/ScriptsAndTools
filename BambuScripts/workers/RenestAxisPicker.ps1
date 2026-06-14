param(
    [string]$FinalPath,
    [string]$NestPath = ""
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ════════════════════════════════════════════════════════════════════════════════
#  RenestAxisPicker.ps1
#
#  Helper tool for RenestFromFinal when a Final.3mf has had its transforms
#  "reset" (merged/re-simplified) and the renest no longer orients correctly.
#
#  Shows the Nest source's centermost instance with its local X/Y/Z axes
#  (red/green/blue, matching Bambu's Object-coordinates gizmo) overlaid on the
#  plate preview, alongside the Final's isolated top-down view rotated by a
#  user-chosen 0/90/180/270 degree increment about the vertical (plate Z) axis.
#
#  Once the user picks the rotation that makes Final's orientation match the
#  reference axes, writes "<stem>_RotCorrection.json" next to the Final.
#  RenestFromFinal_worker.ps1 picks this file up automatically and uses it in
#  place of its auto-detected rotation correction.
# ════════════════════════════════════════════════════════════════════════════════

# ── Resolve input paths ───────────────────────────────────────────────────────
$FinalPath = $FinalPath.Trim('"')
if (-not (Test-Path $FinalPath)) { Write-Error "Final file not found: $FinalPath"; exit 1 }

$finalDir  = Split-Path $FinalPath -Parent
$finalStem = [System.IO.Path]::GetFileNameWithoutExtension($FinalPath) -replace '(?i)_Final$', ''

if ([string]::IsNullOrWhiteSpace($NestPath)) {
    $nestCandidate = Join-Path $finalDir "${finalStem}_Nest.3mf"
    $fullCandidate = Join-Path $finalDir "${finalStem}_Full.3mf"
    if (Test-Path $nestCandidate) { $NestPath = $nestCandidate }
    elseif (Test-Path $fullCandidate) { $NestPath = $fullCandidate }
    else { Write-Error "Could not find a Nest or Full file alongside the Final. Pass -NestPath explicitly."; exit 1 }
}
$NestPath = $NestPath.Trim('"')

$workDir = Join-Path $env:TEMP ("RenestAxisPicker_" + [System.Guid]::NewGuid().ToString("N").Substring(0,8))
New-Item -ItemType Directory -Path $workDir | Out-Null

# ── Helpers ───────────────────────────────────────────────────────────────────
function Read-ZipEntryBytes([string]$zipPath, [string]$entryName) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entry = $zip.GetEntry($entryName)
        if ($null -eq $entry) { return $null }
        $ms = New-Object System.IO.MemoryStream
        $entry.Open().CopyTo($ms)
        return $ms.ToArray()
    } finally { $zip.Dispose() }
}
function Read-ZipEntryText([string]$zipPath, [string]$entryName) {
    $bytes = Read-ZipEntryBytes $zipPath $entryName
    if ($null -eq $bytes) { return $null }
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}
function Save-ZipEntry([string]$zipPath, [string]$entryName, [string]$outPath) {
    $bytes = Read-ZipEntryBytes $zipPath $entryName
    if ($null -eq $bytes) { return $false }
    [System.IO.File]::WriteAllBytes($outPath, $bytes)
    return $true
}

# ── Locate centermost source instance & its rotation matrix ───────────────────
# Bed/preview assumptions for X1C: 256mm bed, 512x512px top_1.png -> 2 px/mm
$bedMM   = 256.0
$imgPx   = 512
$scale   = $imgPx / $bedMM

$plateJsonText = Read-ZipEntryText $NestPath "Metadata/plate_1.json"
if ($null -eq $plateJsonText) { Write-Error "Nest file has no Metadata/plate_1.json: $NestPath"; exit 1 }
$plateJson = $plateJsonText | ConvertFrom-Json

$bedCenter = $bedMM / 2.0
$best = $null; $bestDist = [double]::MaxValue
foreach ($o in $plateJson.bbox_objects) {
    if ($o.id -eq 1000) { continue }   # skip wipe tower
    $cx = ($o.bbox[0] + $o.bbox[2]) / 2
    $cy = ($o.bbox[1] + $o.bbox[3]) / 2
    $d = [Math]::Sqrt(($cx-$bedCenter)*($cx-$bedCenter) + ($cy-$bedCenter)*($cy-$bedCenter))
    if ($d -lt $bestDist) { $bestDist = $d; $best = $o; $bestCx = $cx; $bestCy = $cy }
}
if ($null -eq $best) { Write-Error "No object instances found in $NestPath plate_1.json"; exit 1 }

$modelText = Read-ZipEntryText $NestPath "3D/3dmodel.model"
if ($null -eq $modelText) { Write-Error "Nest file has no 3D/3dmodel.model: $NestPath"; exit 1 }
$itemMatches = [regex]::Matches($modelText, '<item objectid="(\d+)"[^>]*transform="([^"]+)"[^>]*/>')
$bestItem = $null; $bestItemDist = [double]::MaxValue
foreach ($m in $itemMatches) {
    $vals = $m.Groups[2].Value -split '\s+' | ForEach-Object {[double]$_}
    if ($vals.Count -lt 12) { continue }
    $tx = $vals[9]; $ty = $vals[10]
    $d = [Math]::Sqrt(($tx-$bestCx)*($tx-$bestCx) + ($ty-$bestCy)*($ty-$bestCy))
    if ($d -lt $bestItemDist) { $bestItemDist = $d; $bestItem = $vals }
}
if ($null -eq $bestItem) { Write-Error "Could not match a build item transform to the centermost instance."; exit 1 }

# Row-vector rotation matrix R: rows = images of local +X, +Y, +Z (v' = v * R)
$R = [double[]]($bestItem[0..8])
$srcTx = $bestItem[9]; $srcTy = $bestItem[10]

Write-Host "Centermost source instance: plate pos=($([Math]::Round($srcTx,2)), $([Math]::Round($srcTy,2)))mm"
Write-Host ("  R = [{0}]" -f ($R -join ', '))

# ── Draw X(red)/Y(green)/Z(blue) axis arrows at a pixel position ───────────────
function Draw-Axes([System.Drawing.Graphics]$g, [double]$cxPx, [double]$cyPx, [double[]]$R, [double]$scale, [int]$arrowLen) {
    $axes = @(
        @{ name="X"; color=[System.Drawing.Color]::Red;  px=$R[0]; py=$R[1] },
        @{ name="Y"; color=[System.Drawing.Color]::Lime; px=$R[3]; py=$R[4] },
        @{ name="Z"; color=[System.Drawing.Color]::Blue; px=$R[6]; py=$R[7] }
    )
    foreach ($axis in $axes) {
        $dx = $axis.px * $scale
        $dy = -1 * $axis.py * $scale   # plate Y -> image Y is flipped
        $mag = [Math]::Sqrt($dx*$dx + $dy*$dy)
        if ($mag -lt 1e-6) { continue }
        $dx = $dx/$mag*$arrowLen; $dy = $dy/$mag*$arrowLen
        $pen = New-Object System.Drawing.Pen ($axis.color), 4
        $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::ArrowAnchor
        $ptC = New-Object System.Drawing.PointF($cxPx, $cyPx)
        $ptE = New-Object System.Drawing.PointF(($cxPx + $dx), ($cyPx + $dy))
        $g.DrawLine($pen, $ptC, $ptE)
        $pen.Dispose()
    }
    $g.FillEllipse([System.Drawing.Brushes]::Yellow, ($cxPx-4), ($cyPx-4), 8, 8)
}

# ── Build the Nest reference overlay (left panel) ──────────────────────────────
$nestTopPath = Join-Path $workDir "nest_top_1.png"
if (-not (Save-ZipEntry $NestPath "Metadata/top_1.png" $nestTopPath)) {
    Write-Error "Nest file has no Metadata/top_1.png: $NestPath"; exit 1
}
$nestImg = [System.Drawing.Bitmap]::FromFile($nestTopPath)
$g = [System.Drawing.Graphics]::FromImage($nestImg)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$cxPx = $srcTx * $scale
$cyPx = $imgPx - ($srcTy * $scale)
Draw-Axes $g $cxPx $cyPx $R $scale 40
$g.Dispose()
$nestOverlayPath = Join-Path $workDir "nest_overlay.png"
$nestImg.Save($nestOverlayPath, [System.Drawing.Imaging.ImageFormat]::Png)
$nestImg.Dispose()

# ── Final's isolated top-down view (right panel, shown as-is, no overlay) ──────
$finalTopPath = Join-Path $workDir "final_top_1.png"
if (-not (Save-ZipEntry $FinalPath "Metadata/top_1.png" $finalTopPath)) {
    Write-Error "Final file has no Metadata/top_1.png: $FinalPath"; exit 1
}

# ── Reference axis directions (cardinal) for X/Y/Z, from the Nest centermost R ──
# Returns "Left"/"Right"/"Up"/"Down" in image space, or $null if the axis is
# perpendicular to the plate (near-zero in-plane projection).
function Get-CardinalDir([double]$plateX, [double]$plateY, [double]$scale) {
    $dx = $plateX * $scale
    $dy = -1 * $plateY * $scale   # plate Y -> image Y is flipped
    if ([Math]::Sqrt($dx*$dx + $dy*$dy) -lt 1e-3) { return $null }
    if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) {
        if ($dx -gt 0) { return "Right" } else { return "Left" }
    } else {
        if ($dy -gt 0) { return "Down" } else { return "Up" }
    }
}
$refDirs = [ordered]@{
    X = Get-CardinalDir $R[0] $R[1] $scale
    Y = Get-CardinalDir $R[3] $R[4] $scale
    Z = Get-CardinalDir $R[6] $R[7] $scale
}
$refSummary = ($refDirs.Keys | ForEach-Object {
    $v = $refDirs[$_]; if ($null -eq $v) { $v = "(perpendicular - not usable)" }
    "$_ -> $v"
}) -join "   "
Write-Host "Reference directions (Nest): $refSummary"

# ── WPF picker window ───────────────────────────────────────────────────────────
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Renest Axis Picker" Height="700" Width="1000"
        WindowStartupLocation="CenterScreen" Background="#FF2B2D33">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Foreground="White" FontSize="14" Margin="0,0,0,10" TextWrapping="Wrap"
            Text="Left: Nest source centermost instance with reference axes (Red=local X, Green=local Y, Blue=local Z). Right: Final's isolated view (unrotated). Pick an axis below, then click the arrow showing which direction that axis currently points on the Final, relative to the object."/>
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" BorderBrush="Gray" BorderThickness="1" Margin="0,0,5,0">
                <Image Name="imgNest" Stretch="Uniform"/>
            </Border>
            <Grid Grid.Column="1" Margin="5,0,0,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnUp"    Grid.Row="0" Grid.Column="1" Content="Up"    Width="60" Height="30" HorizontalAlignment="Center" Margin="0,0,0,4"/>
                <Button Name="btnLeft"  Grid.Row="1" Grid.Column="0" Content="Left"  Width="50" Height="40" VerticalAlignment="Center" Margin="0,0,4,0"/>
                <Border Grid.Row="1" Grid.Column="1" BorderBrush="Gray" BorderThickness="1">
                    <Image Name="imgFinal" Stretch="Uniform"/>
                </Border>
                <Button Name="btnRight" Grid.Row="1" Grid.Column="2" Content="Right" Width="50" Height="40" VerticalAlignment="Center" Margin="4,0,0,0"/>
                <Button Name="btnDown"  Grid.Row="2" Grid.Column="1" Content="Down"  Width="60" Height="30" HorizontalAlignment="Center" Margin="0,4,0,0"/>
            </Grid>
        </Grid>
        <TextBlock Name="txtRef" Grid.Row="2" Foreground="#FFCCCCCC" FontSize="13" Margin="0,8,0,0" HorizontalAlignment="Center"/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,8,0,0">
            <TextBlock Text="Axis: " Foreground="White" VerticalAlignment="Center" FontSize="14" Margin="0,0,10,0"/>
            <RadioButton Name="rbX" Content="X" GroupName="axis" IsChecked="True" Foreground="Red"        FontWeight="Bold" Margin="10,0" FontSize="16"/>
            <RadioButton Name="rbY" Content="Y" GroupName="axis" Foreground="LimeGreen"  FontWeight="Bold" Margin="10,0" FontSize="16"/>
            <RadioButton Name="rbZ" Content="Z" GroupName="axis" Foreground="DodgerBlue" FontWeight="Bold" Margin="10,0" FontSize="16"/>
        </StackPanel>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
            <Button Name="btnConfirm" Content="Confirm" Width="120" Height="32" Margin="10,0"/>
            <Button Name="btnCancel" Content="Cancel" Width="120" Height="32" Margin="10,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$imgNest  = $window.FindName("imgNest")
$imgFinal = $window.FindName("imgFinal")
$txtRef   = $window.FindName("txtRef")
$dirButtons = @{
    Up    = $window.FindName("btnUp")
    Down  = $window.FindName("btnDown")
    Left  = $window.FindName("btnLeft")
    Right = $window.FindName("btnRight")
}
$axisButtons = @{
    X = $window.FindName("rbX")
    Y = $window.FindName("rbY")
    Z = $window.FindName("rbZ")
}
$btnConfirm = $window.FindName("btnConfirm")
$btnCancel  = $window.FindName("btnCancel")

function Load-BitmapImage([string]$path) {
    $bi = New-Object System.Windows.Media.Imaging.BitmapImage
    $bi.BeginInit()
    $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bi.UriSource = New-Object System.Uri($path)
    $bi.EndInit()
    $bi.Freeze()
    return $bi
}

$imgNest.Source  = Load-BitmapImage $nestOverlayPath
$imgFinal.Source = Load-BitmapImage $finalTopPath

# Image-space cardinal directions, clockwise-visual angle convention (y grows downward):
#   Right=0, Down=90, Left=180, Up=270
$angleOf = @{ Right = 0; Down = 90; Left = 180; Up = 270 }

# Mutable shared state - using a hashtable so event-handler closures (which take
# their own copies of scalar script-scope variables) all mutate the same object.
$script:state = @{ Axis = "X"; Direction = $null; Result = $null }

function Update-RefText {
    $ref = $refDirs[$script:state.Axis]
    if ($null -eq $ref) {
        $txtRef.Text = "Local $($script:state.Axis) axis is perpendicular to the plate in the reference - pick a different axis."
    } else {
        $txtRef.Text = "Reference: local $($script:state.Axis) axis points '$ref' on the Nest. Click the arrow showing where it points on the Final."
    }
}

function Bitmap-To-BitmapImage([System.Drawing.Bitmap]$bmp) {
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $ms.Position = 0
    $bi = New-Object System.Windows.Media.Imaging.BitmapImage
    $bi.BeginInit()
    $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bi.StreamSource = $ms
    $bi.EndInit()
    $bi.Freeze()
    $ms.Dispose()
    return $bi
}

# Rotate a pixel-space direction by thetaDeg (clockwise-visual, y-down convention)
function Rotate-PixelDir([double]$dx, [double]$dy, [double]$thetaDeg) {
    $th = $thetaDeg * [Math]::PI / 180.0
    $c = [Math]::Cos($th); $s = [Math]::Sin($th)
    return @(($dx*$c - $dy*$s), ($dx*$s + $dy*$c))
}

# Find the pixel-space center of the object's silhouette (bounding box of pixels
# that differ from the corner/background color), so the axis overlay is centered
# on the object itself rather than on the image canvas.
function Get-ContentCenter([System.Drawing.Bitmap]$bmp) {
    $bg = $bmp.GetPixel(0, 0)
    $minX = $bmp.Width; $maxX = -1; $minY = $bmp.Height; $maxY = -1
    $step = 2
    for ($y = 0; $y -lt $bmp.Height; $y += $step) {
        for ($x = 0; $x -lt $bmp.Width; $x += $step) {
            $p = $bmp.GetPixel($x, $y)
            $diff = [Math]::Abs($p.R-$bg.R) + [Math]::Abs($p.G-$bg.G) + [Math]::Abs($p.B-$bg.B) + [Math]::Abs($p.A-$bg.A)
            if ($diff -gt 24) {
                if ($x -lt $minX) { $minX = $x }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    if ($maxX -lt $minX -or $maxY -lt $minY) {
        return @(($bmp.Width/2.0), ($bmp.Height/2.0))
    }
    return @(((($minX+$maxX)/2.0)), ((($minY+$maxY)/2.0)))
}

# Redraw the Final preview. If both an axis and a direction have been picked,
# overlay all three reference axes (X/Y/Z) rotated so the selected axis points
# in the chosen direction - lets the user sanity-check the other two axes too.
function Update-FinalPreview {
    $bmp = [System.Drawing.Bitmap]::FromFile($finalTopPath)
    $dir = $script:state.Direction
    $ref = $refDirs[$script:state.Axis]
    if ($null -ne $dir -and $null -ne $ref) {
        $thetaPreview = ($angleOf[$dir] - $angleOf[$ref] + 360) % 360
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $center = Get-ContentCenter $bmp
        $cx = $center[0]; $cy = $center[1]
        $axesDef = @(
            @{ name="X"; color=[System.Drawing.Color]::Red;  px=$R[0]; py=$R[1] },
            @{ name="Y"; color=[System.Drawing.Color]::Lime; px=$R[3]; py=$R[4] },
            @{ name="Z"; color=[System.Drawing.Color]::Blue; px=$R[6]; py=$R[7] }
        )
        $arrowLen = [Math]::Min($bmp.Width, $bmp.Height) / 6.0
        foreach ($axis in $axesDef) {
            $dx0 = $axis.px * $scale
            $dy0 = -1 * $axis.py * $scale
            $mag = [Math]::Sqrt($dx0*$dx0 + $dy0*$dy0)
            if ($mag -lt 1e-6) { continue }
            $rot = Rotate-PixelDir $dx0 $dy0 $thetaPreview
            $rdx = $rot[0]/$mag*$arrowLen; $rdy = $rot[1]/$mag*$arrowLen
            $pen = New-Object System.Drawing.Pen ($axis.color), 4
            $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::ArrowAnchor
            $ptC = New-Object System.Drawing.PointF($cx, $cy)
            $ptE = New-Object System.Drawing.PointF(($cx + $rdx), ($cy + $rdy))
            $g.DrawLine($pen, $ptC, $ptE)
            $pen.Dispose()
        }
        $g.FillEllipse([System.Drawing.Brushes]::Yellow, ($cx-4), ($cy-4), 8, 8)
        $g.Dispose()
    }
    $imgFinal.Source = Bitmap-To-BitmapImage $bmp
    $bmp.Dispose()
}

Update-RefText
Update-FinalPreview

$normalBrush = [System.Windows.Media.Brushes]::LightGray
$selectedBrush = [System.Windows.Media.Brushes]::LightSkyBlue

foreach ($axisName in @("X","Y","Z")) {
    $axisButtons[$axisName].Add_Checked({
        param($s,$e)
        $script:state.Axis = $s.Tag
        Update-RefText
        Update-FinalPreview
    }.GetNewClosure())
    $axisButtons[$axisName].Tag = $axisName
}

foreach ($dirName in @("Up","Down","Left","Right")) {
    $dirButtons[$dirName].Add_Click({
        param($s,$e)
        $script:state.Direction = $s.Tag
        foreach ($b in $dirButtons.Values) { $b.Background = $normalBrush }
        $s.Background = $selectedBrush
        Update-FinalPreview
    }.GetNewClosure())
    $dirButtons[$dirName].Tag = $dirName
    $dirButtons[$dirName].Background = $normalBrush
}

$btnConfirm.Add_Click({
    $ref = $refDirs[$script:state.Axis]
    if ($null -eq $ref) {
        [System.Windows.MessageBox]::Show("The local $($script:state.Axis) axis is perpendicular to the plate in the Nest reference and can't be used. Pick X, Y, or Z (whichever shows a direction).", "Renest Axis Picker") | Out-Null
        return
    }
    if ($null -eq $script:state.Direction) {
        [System.Windows.MessageBox]::Show("Click an arrow to indicate which direction the selected axis currently points on the Final.", "Renest Axis Picker") | Out-Null
        return
    }
    $script:state.Result = @{ axis = $script:state.Axis; refDir = $ref; userDir = $script:state.Direction }
    $window.Close()
})
$btnCancel.Add_Click({
    $script:state.Result = $null
    $window.Close()
})

$window.ShowDialog() | Out-Null

if ($null -eq $script:state.Result) {
    Write-Host "Cancelled - no correction file written."
    Remove-Item -Recurse -Force $workDir -ErrorAction SilentlyContinue
    exit 0
}
$script:result = $script:state.Result

# ── Compute rotCorrection from the (axis, refDir, userDir) selection ──────────
# thetaImage = angle to rotate userDir (CW, visually) so it lands on refDir.
# Plate-space rotation (row-vector R_z(phi), CCW) relates to image rotation by phi = -thetaImage.
$thetaImage = ($angleOf[$script:result.refDir] - $angleOf[$script:result.userDir] + 360) % 360
$phiDeg = (360 - $thetaImage) % 360
$phi = $phiDeg * [Math]::PI / 180.0

# R_final_reset = rotation about the vertical (plate Z) axis by phi
# (row-vector convention, v' = v * R).  rotCorrection = transpose(R_final_reset).
$cosT = [Math]::Cos($phi); $sinT = [Math]::Sin($phi)
$Rfinal = [double[]](
    $cosT,  $sinT, 0,
   -$sinT,  $cosT, 0,
    0,      0,     1
)
$rotCorrection = [double[]]($Rfinal[0],$Rfinal[3],$Rfinal[6], $Rfinal[1],$Rfinal[4],$Rfinal[7], $Rfinal[2],$Rfinal[5],$Rfinal[8])

$jsonOut = [ordered]@{
    selection        = $script:result
    angleDegrees     = $phiDeg
    rotCorrection    = $rotCorrection
    sourceCentermost = @{ tx = $srcTx; ty = $srcTy; R = $R }
}
$outJsonPath = Join-Path $finalDir "${finalStem}_RotCorrection.json"
$jsonOut | ConvertTo-Json -Depth 5 | Set-Content -Path $outJsonPath -Encoding UTF8

Write-Host ("Selection: axis=$($script:result.axis) refDir=$($script:result.refDir) userDir=$($script:result.userDir)")
Write-Host "Plate-Z rotation: $phiDeg degrees"
Write-Host "rotCorrection = [$($rotCorrection -join ', ')]"
Write-Host "Wrote: $outJsonPath"

Remove-Item -Recurse -Force $workDir -ErrorAction SilentlyContinue
