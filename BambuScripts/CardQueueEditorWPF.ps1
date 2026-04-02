# ════════════════════════════════════════════════════════════════════════════════
# FULL WPF BATCH PRE-FLIGHT EDITOR ENGINE
# ════════════════════════════════════════════════════════════════════════════════

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'

# --- 1. ENVIRONMENT & LIBRARY ---
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = [System.Environment]::CurrentDirectory }
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

$LibraryColors = [ordered]@{}
$HexToName = @{}
if (Test-Path $colorCsvPath) {
    $csvLines = Get-Content -Path $colorCsvPath
    foreach ($line in $csvLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split ','
        if ($parts.Count -ge 4) {
            $name = $parts[0].Replace('"','').Trim()
            if ($name -match '(?i)^name$' -or $name -eq "N/A" -or $name -eq "") { continue }
            try {
                $r = [int]$parts[1].Replace('"','').Trim()
                $g = [int]$parts[2].Replace('"','').Trim()
                $b = [int]$parts[3].Replace('"','').Trim()
                $hex = "#{0:X2}{1:X2}{2:X2}FF" -f $r, $g, $b
                $LibraryColors[$name] = $hex
                $HexToName[$hex] = $name
                $HexToName[$hex.Substring(0,7)] = $name
            } catch { continue }
        }
    }
}

# --- 2. C# CLASSES ---
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ModernFolderPicker {
    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }
    [ComImport, Guid("D57C7288-D4AD-4768-BE02-9D969532D960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
        [PreserveSig] int Show(IntPtr hwndOwner);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void GetResult(out IShellItem ppsi);
    }
    static readonly Guid CLSID_FileOpenDialog = new Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7");
    const uint FOS_PICKFOLDERS = 0x00000020;
    const uint FOS_FORCEFILESYSTEM = 0x00000040;
    const uint SIGDN_FILESYSPATH = 0x80058000;
    public static string Pick(IntPtr owner, string title) {
        try {
            Type t = Type.GetTypeFromCLSID(CLSID_FileOpenDialog);
            object inst = Activator.CreateInstance(t);
            var dialog = (IFileOpenDialog)inst;
            try {
                uint opts; dialog.GetOptions(out opts);
                dialog.SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                if (!string.IsNullOrEmpty(title)) dialog.SetTitle(title);
                if (dialog.Show(owner) != 0) return null;
                IShellItem item; dialog.GetResult(out item);
                string path; item.GetDisplayName(SIGDN_FILESYSPATH, out path);
                Marshal.ReleaseComObject(item);
                return path;
            } finally { Marshal.ReleaseComObject(dialog); }
        } catch { return null; }
    }
}
'@

try {
    Add-Type -TypeDefinition @'
    using System;
    using System.Drawing;
    using System.Collections.Generic;
    public class MapLine { public Point Start; public Point End; public Color LineColor; }
    public class FastMergeMap {
        public static List<MapLine> GetMergeLines(Bitmap pre, Bitmap post) {
            var preAnchors = new Dictionary<int, Point>();
            var postAnchors = new Dictionary<int, Point>();
            int w = pre.Width; int h = pre.Height;
            for (int y = 0; y < h; y++) {
                for (int x = 0; x < w; x++) {
                    Color cPre = pre.GetPixel(x, y);
                    if (cPre.A == 255 && (cPre.R > 10 || cPre.G > 10 || cPre.B > 10)) {
                        int argb = cPre.ToArgb();
                        if (!preAnchors.ContainsKey(argb)) preAnchors[argb] = new Point(x, y);
                    }
                    if (x < post.Width && y < post.Height) {
                        Color cPost = post.GetPixel(x, y);
                        if (cPost.A == 255 && (cPost.R > 10 || cPost.G > 10 || cPost.B > 10)) {
                            int argb = cPost.ToArgb();
                            if (!postAnchors.ContainsKey(argb)) postAnchors[argb] = new Point(x, y);
                        }
                    }
                }
            }
            var lines = new List<MapLine>();
            foreach (var oldPt in preAnchors.Values) {
                if (oldPt.X < post.Width && oldPt.Y < post.Height) {
                    Color currentColor = post.GetPixel(oldPt.X, oldPt.Y);
                    int currentArgb = currentColor.ToArgb();
                    if (postAnchors.ContainsKey(currentArgb))
                        lines.Add(new MapLine { Start = oldPt, End = postAnchors[currentArgb], LineColor = currentColor });
                }
            }
            return lines;
        }
    }
'@ -ReferencedAssemblies "System.Drawing"
} catch {}

# --- 3. HELPER FUNCTIONS ---
function Get-WpfColor([string]$hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { $hex = "#808080" }
    if ($hex.Length -eq 9) { $hex = "#" + $hex.Substring(1,6) }
    try { return [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($hex)) }
    catch { return [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Colors]::Gray) }
}

function Get-WpfColorMedia([string]$hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { return [System.Windows.Media.Colors]::Gray }
    if ($hex.Length -eq 9) { $hex = "#" + $hex.Substring(1,6) }
    try { return [System.Windows.Media.ColorConverter]::ConvertFromString($hex) }
    catch { return [System.Windows.Media.Colors]::Gray }
}

function Create-TextBlock([string]$text, [string]$hexColor, [int]$fontSize, [string]$weight) {
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $text; $tb.Foreground = Get-WpfColor $hexColor; $tb.FontSize = $fontSize
    if ($weight -eq "Bold") { $tb.FontWeight = [System.Windows.FontWeights]::Bold }
    $tb.VerticalAlignment = "Center"
    return $tb
}

function Load-WpfImage([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit()
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.StreamSource = $ms
        $bmp.EndInit()
        $ms.Close(); $ms.Dispose()
        return $bmp
    } catch { return $null }
}

function ParseFile([string]$filename) {
    $compoundExts = @('.gcode.3mf', '.gcode.stl', '.gcode.step', '.f3d.3mf')
    $ext = $null
    foreach ($ce in $compoundExts) { if ($filename.ToLower().EndsWith($ce.ToLower())) { $ext = $filename.Substring($filename.Length - $ce.Length); break } }
    if (-not $ext) { $ext = [System.IO.Path]::GetExtension($filename) }
    if (-not $ext) { $ext = '' }
    $stem  = $filename.Substring(0, $filename.Length - $ext.Length)
    $parts = [string[]]( ($stem -split '[\s._-]+') | Where-Object { $_ -ne '' } )
    if ($ext -ieq '.png') { return @{ Suffix = ''; Extension = $ext; Stem = $stem; Parts = $parts } }
    $suffix = if ($parts.Count -gt 0) { $parts[-1] } else { $stem }
    return @{ Suffix = $suffix; Extension = $ext; Stem = $stem; Parts = $parts }
}

function SmartFill([string]$anchorName, [string]$gpName) {
    $stem = $anchorName -replace '(?i)\.gcode\.3mf$|\.3mf$', ''
    $stem = $stem -replace '(?i)_Full$', ''
    $escaped = [regex]::Escape($gpName)
    if ($gpName -ne '' -and $stem -imatch "^(.+)_${escaped}$") { $prefix = $Matches[1] } else { $prefix = $stem }
    $parts = $prefix -split '_'
    # BUG FIX: was returning $parts (full array) instead of $parts[0] for Char
    if ($parts.Count -ge 2) { return @{ Char = $parts[0]; Adj = ($parts[1..($parts.Count-1)] -join '') } }
    return @{ Char = $prefix; Adj = '' }
}

function Invoke-RandomizePickColors($sourcePath, $destPath) {
    $rng = New-Object System.Random
    try {
        $bmp = New-Object System.Drawing.Bitmap($sourcePath)
        $colorMap = @{}
        for ($y = 0; $y -lt $bmp.Height; $y++) {
            for ($x = 0; $x -lt $bmp.Width; $x++) {
                $px = $bmp.GetPixel($x, $y)
                if ($px.A -lt 10) { continue }
                $key = "$($px.R),$($px.G),$($px.B)"
                if (-not $colorMap.ContainsKey($key)) { $colorMap[$key] = [System.Drawing.Color]::FromArgb($px.A, $rng.Next(20,256), $rng.Next(20,256), $rng.Next(20,256)) }
                $bmp.SetPixel($x, $y, $colorMap[$key])
            }
        }
        $bmp.Save($destPath) | Out-Null; $bmp.Dispose()
        return $true
    } catch { return $false }
}

# --- 4. BACKEND QUEUE DATA ---
$script:jobs = New-Object System.Collections.ArrayList
$script:processQueue = New-Object System.Collections.Queue
$script:activeProcess = $null
$script:activeProcessJob = $null
$gpQueue = [ordered]@{}

# Check for files passed by the VBScript, otherwise launch empty
$foundFiles = @()
if ($args.Count -gt 0) {
    foreach ($p in $args) {
        if (Test-Path $p -PathType Container) { $foundFiles += Get-ChildItem -Path $p -Filter "*Full.3mf" -Recurse -File }
        elseif ($p -match '(?i)Full\.3mf$') { $foundFiles += Get-Item $p }
    }
}

foreach ($f in $foundFiles) {
    $parentPath = $f.DirectoryName
    $gp = $f.Directory.Parent
    $gpPath = if ($gp) { $gp.FullName } else { "ROOT_" + $parentPath }
    if (-not $gpQueue.Contains($gpPath)) { $gpQueue[$gpPath] = [ordered]@{} }
    if (-not $gpQueue[$gpPath].Contains($parentPath)) { $gpQueue[$gpPath][$parentPath] = $f }
}

# --- 5. THE XAML LAYOUT ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Batch Pre-Flight Editor - WPF Engine"
        Width="1550" Height="850" MinWidth="1200" MinHeight="600"
        Background="#16171B" WindowStartupLocation="CenterScreen" AllowDrop="True">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Background="#1C1D23" Grid.Row="0">
            <Grid>
                <TextBlock Name="LblGlobalTitle" Text="Loading files into queue..." Foreground="White" FontSize="18" FontWeight="Bold" VerticalAlignment="Center" Margin="15,0,0,0"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,15,0">
                    <Button Name="BtnCombineData" Content="Combine TSV Data" Background="#E8A135" Foreground="White" FontWeight="Bold" Width="180" Height="40" Margin="0,0,15,0" BorderThickness="0" Cursor="Hand"/>
                    <Button Name="BtnProcessAll" Content="Add All To Queue" Background="#4CAF72" Foreground="White" FontWeight="Bold" Width="220" Height="40" IsEnabled="False" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>
            </Grid>
        </Border>
        <ScrollViewer Grid.Row="1" Background="#0D0E10" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel Name="MainStack" Orientation="Vertical" Margin="15"/>
        </ScrollViewer>
    </Grid>
</Window>
"@
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$lblGlobalTitle = $window.FindName("LblGlobalTitle")
$btnProcessAll  = $window.FindName("BtnProcessAll")
$btnCombineData = $window.FindName("BtnCombineData")
$mainStack      = $window.FindName("MainStack")

function Update-GlobalProcessAllStatus {
    $hasAnyCollision = $false
    foreach ($gp in $script:jobs) {
        foreach ($p in $gp.Parents) {
            if ($p.HasCollision) { $hasAnyCollision = $true; break }
        }
        if ($hasAnyCollision) { break }
    }
    if ($hasAnyCollision) {
        $btnProcessAll.IsEnabled = $false
        $btnProcessAll.Background = Get-WpfColor "#555555"
    } else {
        $btnProcessAll.IsEnabled = $true
        $btnProcessAll.Background = Get-WpfColor "#4CAF72"
    }
}

# --- 6. CORE LOGIC FUNCTIONS ---
function Validate-PJob($pJob) {
    if ($pJob.IsQueued -or $pJob.IsDone) { return }
    $colorsSafe = $true
    foreach ($slot in $pJob.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $colorsSafe = $false } }
    if ($pJob.HasCollision) {
        $pJob.BtnApply.Content = "Name Collision!"; $pJob.BtnApply.Background = Get-WpfColor "#D95F5F"; $pJob.BtnApply.IsEnabled = $false
    } elseif (-not $colorsSafe) {
        $pJob.BtnApply.Content = "Unmatched Colors"; $pJob.BtnApply.Background = Get-WpfColor "#E8A135"; $pJob.BtnApply.IsEnabled = $false
    } else {
        $pJob.BtnApply.Content = "Add to Queue"; $pJob.BtnApply.Background = Get-WpfColor "#4CAF72"; $pJob.BtnApply.IsEnabled = $true
    }
}

function Update-ParentPreview($pJob, $gpJob) {
    $ch = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
    $ad = $pJob.TBAdj.Text -replace '[^a-zA-Z0-9]', ''
    $th = $gpJob.TBTheme.Text -replace '[^a-zA-Z0-9]', ''

    # 1. Split CamelCase into spaces (e.g., "BabyDragon" -> "Baby Dragon")
    $chSpaced = $ch -creplace '([a-z])([A-Z])', '$1 $2'
    $adSpaced = $ad -creplace '([a-z])([A-Z])', '$1 $2'

    # 2. Format with Adjective in parentheses
    $displayTitle = $chSpaced
    if (-not [string]::IsNullOrWhiteSpace($adSpaced)) {
        $displayTitle += " ($adSpaced)"
    }

    # 3. Apply to the UI card in ALL CAPS
    if ($null -ne $pJob.LblCharCard) { $pJob.LblCharCard.Text = $displayTitle.ToUpper() }

    $nameCounts = @{}
    foreach ($r in $pJob.FileRows) {
        $sf = $r.SuffixBox.Text -replace '[^a-zA-Z0-9]', ''
        $parts = New-Object System.Collections.ArrayList
        if ($ch) { [void]$parts.Add($ch) }; if ($ad) { [void]$parts.Add($ad) }
        if ($th) { [void]$parts.Add($th) }; if ($sf) { [void]$parts.Add($sf) }
        $r.TargetName = ($parts.ToArray() -join '_') + $r.Ext
        if (-not $nameCounts.ContainsKey($r.TargetName)) { $nameCounts[$r.TargetName] = 0 }
        $nameCounts[$r.TargetName]++
    }

    $hasCollision = $false
    foreach ($r in $pJob.FileRows) {
        $r.NewLbl.Text = $r.TargetName
        if ($nameCounts[$r.TargetName] -gt 1) { $r.NewLbl.Foreground = Get-WpfColor "#D95F5F"; $hasCollision = $true }
        else { $r.NewLbl.Foreground = Get-WpfColor "#4CAF72" }
    }
    $pJob.HasCollision = $hasCollision
    Validate-PJob $pJob
    Update-GlobalProcessAllStatus
}

function Add-FileRow($pJob, $gpJob, $fi) {
    $parsed = ParseFile $fi.Name

    $fRow = New-Object System.Windows.Controls.Border
    $fRow.Background = Get-WpfColor "#16171B"; $fRow.BorderBrush = Get-WpfColor "#2A2C35"
    $fRow.BorderThickness = New-Object System.Windows.Thickness(0,0,0,1); $fRow.Height = 40

    $fGrid = New-Object System.Windows.Controls.Grid
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(70))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(30))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(40))}))
    $fRow.Child = $fGrid

    $sBadge = New-Object System.Windows.Controls.TextBox
    $sBadge.Text = $parsed.Suffix; $sBadge.Background = Get-WpfColor "#1E2028"; $sBadge.Foreground = Get-WpfColor "#E8A135"
    $sBadge.VerticalAlignment = "Center"; $sBadge.Margin = New-Object System.Windows.Thickness(10,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($sBadge, 0); $fGrid.Children.Add($sBadge) | Out-Null

    $lOld = Create-TextBlock $fi.Name "#6B6E7A" 11 "Normal"
    $lOld.Margin = New-Object System.Windows.Thickness(10,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($lOld, 1); $fGrid.Children.Add($lOld) | Out-Null

    $lArr = Create-TextBlock "->" "#A0A0A0" 12 "Normal"
    [System.Windows.Controls.Grid]::SetColumn($lArr, 2); $fGrid.Children.Add($lArr) | Out-Null

    $lNew = Create-TextBlock "" "#4CAF72" 11 "Bold"
    [System.Windows.Controls.Grid]::SetColumn($lNew, 3); $fGrid.Children.Add($lNew) | Out-Null

    $btnDel = New-Object System.Windows.Controls.Button
    $btnDel.Content = "X"; $btnDel.Background = Get-WpfColor "#D95F5F"; $btnDel.Foreground = Get-WpfColor "#FFFFFF"
    $btnDel.BorderThickness = 0; $btnDel.Width = 20; $btnDel.Height = 20; $btnDel.Cursor = [System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnDel, 4); $fGrid.Children.Add($btnDel) | Out-Null

    $frObj = [PSCustomObject]@{ OldPath = $fi.FullName; SuffixBox = $sBadge; NewLbl = $lNew; Ext = $parsed.Extension; TargetName = "" }
    $pJob.FileRows.Add($frObj) | Out-Null

    $sBadge.Tag = @{ P = $pJob; G = $gpJob }
    $sBadge.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })

    $btnDel.Tag = @{ P = $pJob; G = $gpJob; Row = $fRow; FileRow = $frObj; Path = $fi.FullName }
    $btnDel.Add_Click({
        $t = $this.Tag
        $res = [System.Windows.MessageBox]::Show("Permanently delete:`n$($t.Path | Split-Path -Leaf)?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($res -eq 'Yes') {
            Remove-Item $t.Path -Force -ErrorAction SilentlyContinue
            $t.P.FileRows.Remove($t.FileRow) | Out-Null
            $t.P.PnlFiles.Children.Remove($t.Row) | Out-Null
            Update-ParentPreview $t.P $t.G
        }
    })
    $pJob.PnlFiles.Children.Add($fRow) | Out-Null
}

function Refresh-PJob($pJob, $gpJob) {
    # Reset state
    $pJob.ProcessedAnchorPath = ""
    $newAnchor = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)Full\.3mf$' -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Select-Object -First 1
    if ($newAnchor) { $pJob.AnchorFile = $newAnchor }

    # Re-extract temp work
    if (Test-Path $pJob.TempWork) { Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue }
    try { [System.IO.Compression.ZipFile]::ExtractToDirectory($pJob.AnchorFile.FullName, $pJob.TempWork) } catch {}

    # Update folder label
    if ($pJob.LblFolder) { $pJob.LblFolder.Text = "Folder: $(Split-Path $pJob.FolderPath -Leaf)" }

    $gcodeFile = Get-ChildItem -Path $pJob.FolderPath -Filter "*Full.gcode.3mf" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    # Reload plate image
    if ($gcodeFile) {
        $extractedGcodePlate = Join-Path $pJob.TempWork "plate_1_gcode.png"
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $entry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
            if ($entry) {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $extractedGcodePlate, $true)
                $pJob.PbPlate.Source = Load-WpfImage $extractedGcodePlate
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    } else {
        $baseImgPath = Join-Path $pJob.TempWork "Metadata\plate_1.png"
        $pJob.PbPlate.Source = Load-WpfImage $baseImgPath
    }

    # Reload custom PNG
    $customPng = Get-ChildItem -Path $pJob.FolderPath -Filter "*.png" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" } | Select-Object -First 1
    if ($customPng) {
        $pJob.CustomImagePath = $customPng.FullName
        $pJob.PbPlate.Source = Load-WpfImage $customPng.FullName
    } else { $pJob.CustomImagePath = $null }

    # Reload pick image
    $pickPath = $null
    if ($gcodeFile) {
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)(^|/)pick_1\.png$" } | Select-Object -First 1
            if ($pickEntry) {
                $rawPickPath = Join-Path $pJob.TempWork "pick_1_raw.png"
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $rawPickPath, $true)
                $pickPath = Join-Path $pJob.TempWork "pick_1.png"
                Invoke-RandomizePickColors $rawPickPath $pickPath | Out-Null
                if (-not (Test-Path $pickPath)) { $pickPath = $rawPickPath }
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    }
    if ($pickPath) { $pJob.PbPick.Source = Load-WpfImage $pickPath }

    # Reset overlays
    $pJob.ProcessingOverlay.Visibility = "Collapsed"
    $pJob.PickProcessingOverlay.Visibility = "Collapsed"
    if ($pJob.PbPlateFinished) { $pJob.PbPlateFinished.Visibility = "Collapsed" }

    # Reload file rows
    $pJob.PnlFiles.Children.Clear(); $pJob.FileRows.Clear()
    $files = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($fi in $files) { Add-FileRow $pJob $gpJob $fi }
    Update-ParentPreview $pJob $gpJob

    # Reset UI state
    $pJob.BtnApply.Content = "Add to Queue"; $pJob.BtnApply.Background = Get-WpfColor "#4CAF72"; $pJob.BtnApply.IsEnabled = $true; $pJob.BtnApply.Width = 150
    if ($pJob.BtnRevertDone) { $pJob.BtnRevertDone.Visibility = "Collapsed" }
    $pJob.RowPanel.IsEnabled = $true
    $pJob.IsDone = $false; $pJob.IsQueued = $false

    $pJob.ChkMerge.IsEnabled = $true; $pJob.ChkSlice.IsEnabled = $true
    $pJob.ChkExtract.IsEnabled = $true; $pJob.ChkImage.IsEnabled = $true
    $pJob.ChkMerge.IsChecked = $true; $pJob.ChkSlice.IsChecked = $true
    $pJob.ChkExtract.IsChecked = $true; $pJob.ChkImage.IsChecked = $true
}

function Enqueue-PJob($pJob, $gpJob) {
    if ($pJob.IsQueued -or $pJob.IsDone -or $pJob.HasCollision) { return }
    foreach ($slot in $pJob.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { return } }

    $pJob.IsQueued = $true
    $pJob.BtnApply.Content = "Queued..."; $pJob.BtnApply.Background = Get-WpfColor "#E8A135"
    $pJob.RowPanel.IsEnabled = $false
    $pJob.ProcessingOverlay.Text = "[ PREPARING ]"; $pJob.ProcessingOverlay.Visibility = "Visible"
    $pJob.PickProcessingOverlay.Text = "[ PREPARING ]"; $pJob.PickProcessingOverlay.Visibility = "Visible"

    $script:processQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob })
}

function Start-NextProcess {
    if ($script:activeProcess -ne $null -or $script:processQueue.Count -eq 0) { return }

    $jobWrapper = $script:processQueue.Dequeue()
    $pJob = $jobWrapper.PJob; $gpJob = $jobWrapper.GpJob
    $script:activeProcessJob = $jobWrapper
    $pJob.BtnApply.Content = "Processing..."; $pJob.IsQueued = $false

    $th = $gpJob.TBTheme.Text -replace '[^a-zA-Z0-9]', ''
    $oldGrand = if ($gpJob.DiGrand) { $gpJob.DiGrand.FullName } else { "" }

    # Uncheck merge/image if colors unresolved
    $colorsSafe = $true
    foreach ($slot in $pJob.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $colorsSafe = $false } }
    if (-not $colorsSafe) { $pJob.ChkMerge.IsChecked = $false; $pJob.ChkImage.IsChecked = $false }

    # Apply color substitutions to extracted temp files
    $allTextFiles = Get-ChildItem -Path $pJob.TempWork -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
    $modifiedFiles = New-Object System.Collections.ArrayList

    foreach ($file in $allTextFiles) {
        $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
        $modified = $false
        foreach ($slot in $pJob.UISlots) {
            $selName = $slot.Combo.Text
            if ($LibraryColors.Contains($selName)) {
                $newHex = $LibraryColors[$selName].ToUpper(); $oldHex = $slot.OldHex.ToUpper()
                $oldHex9 = if ($oldHex.Length -eq 7) { $oldHex + "FF" } else { $oldHex }
                $oldHex7 = $oldHex.Substring(0,7)
                $newHex9 = if ($newHex.Length -eq 7) { $newHex + "FF" } else { $newHex }
                $newHex7 = $newHex.Substring(0,7)
                if ($content -match "(?i)$oldHex9") { $content = $content -ireplace [regex]::Escape($oldHex9), $newHex9; $modified = $true }
                if ($content -match "(?i)$oldHex7") { $content = $content -ireplace [regex]::Escape($oldHex7), $newHex7; $modified = $true }
            }
        }
        if ($modified) {
            [System.IO.File]::WriteAllText($file.FullName, $content, (New-Object System.Text.UTF8Encoding($false)))
            $modifiedFiles.Add($file) | Out-Null
        }
    }

    # Rename anchor file
    $anchorTargetName = ""; $anchorFileRow = $null
    $currentAnchorLocation = $pJob.AnchorFile.FullName
    foreach ($r in $pJob.FileRows) {
        if ((Split-Path $r.OldPath -Leaf) -eq $pJob.AnchorFile.Name) {
            $anchorTargetName = $r.NewLbl.Text; $anchorFileRow = $r; $currentAnchorLocation = $r.OldPath; break
        }
    }
    if ($anchorTargetName -eq "") { $anchorTargetName = $pJob.AnchorFile.Name }
    $newFilePath = Join-Path $pJob.FolderPath $anchorTargetName

    if ($currentAnchorLocation -ne $newFilePath) {
        if (Test-Path $newFilePath) { Remove-Item $newFilePath -Force -ErrorAction SilentlyContinue }
        try { Rename-Item $currentAnchorLocation $anchorTargetName -Force } catch {}
    }
    if ($null -ne $anchorFileRow) { $anchorFileRow.OldPath = $newFilePath }

    # Inject color changes back into the zip
    if ($modifiedFiles.Count -gt 0) {
        $zip = [System.IO.Compression.ZipFile]::Open($newFilePath, 'Update')
        foreach ($file in $modifiedFiles) {
            $rel = $file.FullName.Substring($pJob.TempWork.Length).TrimStart('\','/').Replace('\','/')
            $entry = $zip.GetEntry($rel)
            if ($entry) { $entry.Delete() }
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $rel) | Out-Null
        }
        $zip.Dispose()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    }

    $pJob.ProcessedAnchorPath = $newFilePath
    [System.GC]::Collect()
    Start-Sleep -Milliseconds 100

    # Rename all other files
    foreach ($r in $pJob.FileRows) {
        if ($r.OldPath -eq $pJob.ProcessedAnchorPath) { continue }
        $targetName = $r.NewLbl.Text
        $newPath = Join-Path $pJob.FolderPath $targetName
        if ($r.OldPath -ne $newPath -and (Test-Path $r.OldPath)) {
            Rename-Item $r.OldPath $targetName -Force; $r.OldPath = $newPath
        }
    }

    # Rename parent folder
    $cleanChar = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
    $cleanAdj  = $pJob.TBAdj.Text  -replace '[^a-zA-Z0-9]', ''
    $pParts = New-Object System.Collections.ArrayList
    if ($cleanChar) { $pParts.Add($cleanChar) | Out-Null }
    if ($cleanAdj)  { $pParts.Add($cleanAdj)  | Out-Null }
    if ($th)        { $pParts.Add($th)         | Out-Null }
    $newParentName = $pParts.ToArray() -join '_'

    $oldFolder = $pJob.FolderPath
    if ($newParentName -ne '' -and $newParentName -ne (Split-Path $oldFolder -Leaf)) {
        $newFolder = Join-Path (Split-Path $oldFolder -Parent) $newParentName
        try {
            Rename-Item $oldFolder $newParentName -Force -ErrorAction Stop
            $pJob.FolderPath = $newFolder
            $pJob.ProcessedAnchorPath = $pJob.ProcessedAnchorPath.Replace($oldFolder, $newFolder)
            if ($pJob.CustomImagePath) { $pJob.CustomImagePath = $pJob.CustomImagePath.Replace($oldFolder, $newFolder) }
            foreach ($r in $pJob.FileRows) { $r.OldPath = $r.OldPath.Replace($oldFolder, $newFolder) }
        } catch {}
    }

    # Rename grandparent folder and propagate to ALL jobs
    if (-not $gpJob.ChkSkip.IsChecked -and $th -ne '' -and $oldGrand -ne '' -and $th -ne (Split-Path $oldGrand -Leaf)) {
        $newGrand = Join-Path (Split-Path $oldGrand -Parent) $th
        try {
            Rename-Item $oldGrand $th -Force -ErrorAction Stop
            $gpJob.GpPath = $newGrand; $gpJob.DiGrand = [System.IO.DirectoryInfo]::new($newGrand)
            foreach ($p in $gpJob.Parents) {
                $p.FolderPath = $p.FolderPath.Replace($oldGrand, $newGrand)
                if ($p.ProcessedAnchorPath) { $p.ProcessedAnchorPath = $p.ProcessedAnchorPath.Replace($oldGrand, $newGrand) }
                if ($p.CustomImagePath) { $p.CustomImagePath = $p.CustomImagePath.Replace($oldGrand, $newGrand) }
                foreach ($fr in $p.FileRows) { $fr.OldPath = $fr.OldPath.Replace($oldGrand, $newGrand) }
            }
            # Propagate to all OTHER grandparent jobs that share this root path
            foreach ($otherGp in $script:jobs) {
                if ($otherGp -ne $gpJob -and $otherGp.GpPath.StartsWith($oldGrand)) {
                    $otherGp.GpPath = $otherGp.GpPath.Replace($oldGrand, $newGrand)
                    $otherGp.DiGrand = [System.IO.DirectoryInfo]::new($otherGp.GpPath)
                    foreach ($otherP in $otherGp.Parents) {
                        $otherP.FolderPath = $otherP.FolderPath.Replace($oldGrand, $newGrand)
                        if ($otherP.ProcessedAnchorPath) { $otherP.ProcessedAnchorPath = $otherP.ProcessedAnchorPath.Replace($oldGrand, $newGrand) }
                        foreach ($fr in $otherP.FileRows) { $fr.OldPath = $fr.OldPath.Replace($oldGrand, $newGrand) }
                    }
                }
            }
        } catch {}
    }

    # Anchor recovery scan
    if (-not (Test-Path -LiteralPath $pJob.ProcessedAnchorPath)) {
        $recoveredAnchor = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '(?i)Full\.3mf$' -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($null -ne $recoveredAnchor) {
            $pJob.ProcessedAnchorPath = $recoveredAnchor.FullName
        } else {
            $pJob.IsDone = $true; $pJob.BtnApply.Content = "No Anchor Found"; $pJob.BtnApply.Background = Get-WpfColor "#D95F5F"
            $pJob.ProcessingOverlay.Visibility = "Collapsed"; $pJob.PickProcessingOverlay.Visibility = "Collapsed"
            $pJob.RowPanel.IsEnabled = $true
            $script:activeProcess = $null; $script:activeProcessJob = $null
            if ($script:processQueue.Count -gt 0) { Start-NextProcess }
            return
        }
    }

    # Build worker script
    $workerScript = Join-Path $env:TEMP "AsyncWorker_$([guid]::NewGuid().ToString().Substring(0,8)).ps1"
    $sb = New-Object System.Text.StringBuilder
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pJob.ProcessedAnchorPath)
    if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = $pJob.AnchorFile.BaseName }

    $dir        = $pJob.FolderPath
    $statusFile = Join-Path $dir "AsyncWorker_Status.txt"
    $basePrefix = if ($baseName.ToLower().EndsWith("full")) { $baseName.Substring(0, $baseName.Length - 4) } else { $baseName + "_" }

    $doMerge   = $pJob.ChkMerge.IsChecked
    $doSlice   = $pJob.ChkSlice.IsChecked
    $doExtract = $pJob.ChkExtract.IsChecked
    $doImage   = $pJob.ChkImage.IsChecked

    $anchorPath = $pJob.ProcessedAnchorPath
    $nestPath   = Join-Path $dir "$($basePrefix)Nest.3mf"
    $finalPath  = Join-Path $dir "$($basePrefix)Final.3mf"
    $tempOut    = Join-Path $dir "$($baseName)_merged_temp.3mf"
    $tempIso    = Join-Path $env:TEMP "iso_$([guid]::NewGuid().ToString().Substring(0,8))"
    $slicedFile = Join-Path $dir "$($baseName).gcode.3mf"
    $singleFile = Join-Path $dir "$($basePrefix)Final.gcode.3mf"
    $tsvBaseName = $baseName -replace '(?i)_Full$', ''
    $tsvFile    = Join-Path $dir "${tsvBaseName}_Data.tsv"

    [void]$sb.AppendLine("`$ErrorActionPreference = 'Continue'")
    [void]$sb.AppendLine("Add-Type -AssemblyName System.IO.Compression.FileSystem")

    if ($doMerge) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'MERGING...' -Force")
        [void]$sb.AppendLine("& `"$scriptDir\merge_3mf_worker.ps1`" -WorkDir `"$($pJob.TempWork)`" -InputPath `"$anchorPath`" -OutputPath `"$tempOut`" -DoColors `"0`"")
        [void]$sb.AppendLine("if (Test-Path `"$tempOut`") {")
        [void]$sb.AppendLine("    if (Test-Path `"$nestPath`") { Remove-Item `"$nestPath`" -Force }")
        [void]$sb.AppendLine("    Rename-Item -Path `"$anchorPath`" -NewName `"$($basePrefix)Nest.3mf`" -Force")
        [void]$sb.AppendLine("    Rename-Item -Path `"$tempOut`"    -NewName `"$($baseName).3mf`"      -Force")
        [void]$sb.AppendLine("    Get-ChildItem -Path `"$dir`" -Filter `"*MergeReport*.txt`" -ErrorAction SilentlyContinue | ForEach-Object { Remove-Item `$_.FullName -Force -ErrorAction SilentlyContinue }")
        [void]$sb.AppendLine("} else { Write-Host '[ERROR] Merge produced no output.' -ForegroundColor Red; exit 1 }")
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'ISOLATING FINAL...' -Force")
        [void]$sb.AppendLine("New-Item -ItemType Directory -Path `"$tempIso`" -Force | Out-Null")
        [void]$sb.AppendLine("[System.IO.Compression.ZipFile]::ExtractToDirectory(`"$nestPath`", `"$tempIso`")")
        [void]$sb.AppendLine("& `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `"$tempIso`" -OutputPath `"$finalPath`"")
        [void]$sb.AppendLine("Remove-Item `"$tempIso`" -Recurse -Force -ErrorAction SilentlyContinue")
    }

    if ($doSlice) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'SLICING...' -Force")
        [void]$sb.AppendLine("Start-Sleep -Seconds 3")
        [void]$sb.AppendLine("& `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$anchorPath`" -IsolatedPath `"$finalPath`"")
    } elseif ($doExtract -or $doImage) {
        # No full slice requested, but we need gcode for data extraction — re-slice Final only
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'RE-SLICING FINAL FOR DATA...' -Force")
        [void]$sb.AppendLine("if (-not (Test-Path `"$finalPath`") -and (Test-Path `"$nestPath`")) {")
        [void]$sb.AppendLine("    `$tmpR = Join-Path `$env:TEMP `"iso_reslice_$([guid]::NewGuid().ToString().Substring(0,8))`"")
        [void]$sb.AppendLine("    New-Item -ItemType Directory -Path `$tmpR -Force | Out-Null")
        [void]$sb.AppendLine("    [System.IO.Compression.ZipFile]::ExtractToDirectory(`"$nestPath`", `$tmpR)")
        [void]$sb.AppendLine("    & `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `$tmpR -OutputPath `"$finalPath`"")
        [void]$sb.AppendLine("    Remove-Item `$tmpR -Recurse -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("}")
        [void]$sb.AppendLine("if (Test-Path `"$finalPath`") {")
        [void]$sb.AppendLine("    Start-Sleep -Seconds 3")
        [void]$sb.AppendLine("    & `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$finalPath`"")
        [void]$sb.AppendLine("}")
    }

    if ($doExtract -or $doImage) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'EXTRACTING DATA...' -Force")
        $imageFlag = if ($doImage) { "-GenerateImage" } else { "" }
        [void]$sb.AppendLine("if (Test-Path `"$slicedFile`") {")
        [void]$sb.AppendLine("    Remove-Item `"$tsvFile`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("    & `"$scriptDir\Extract-3MFData.ps1`" -InputFile `"$slicedFile`" -SingleFile `"$singleFile`" -IndividualTsvPath `"$tsvFile`" $imageFlag")
        [void]$sb.AppendLine("    Remove-Item `"$singleFile`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("} else {")
        [void]$sb.AppendLine("    Set-Content -Path `"$statusFile`" -Value '[ERROR] SLICE FAILED - MISSING GCODE' -Force")
        [void]$sb.AppendLine("    Start-Sleep -Seconds 4")
        [void]$sb.AppendLine("}")
    }

    if ($doImage) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'IMAGE INJECTION...' -Force")
        [void]$sb.AppendLine("`$batPath = Join-Path `"$scriptDir`" 'ReplaceImageNew.bat'")
        [void]$sb.AppendLine("if (Test-Path `$batPath) {")
        [void]$sb.AppendLine("    `$argList = '/c `"`"' + `$batPath + '`" `"' + `"$dir`" + '`"`"'")
        [void]$sb.AppendLine("    Start-Process -FilePath 'cmd.exe' -ArgumentList `$argList -Wait -WindowStyle Hidden")
        [void]$sb.AppendLine("}")
    }

    [void]$sb.AppendLine("Get-ChildItem -Path `"$dir`" -Filter `"*ProcessLog*.txt`" -ErrorAction SilentlyContinue | Remove-Item -Force")
    [void]$sb.AppendLine("Remove-Item `"$statusFile`" -Force -ErrorAction SilentlyContinue")
    [void]$sb.AppendLine("Remove-Item `"$($pJob.TempWork)`" -Recurse -Force -ErrorAction SilentlyContinue")

    Set-Content -Path $workerScript -Value $sb.ToString()
    $script:activeProcess = Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$workerScript`"" -PassThru -WindowStyle Hidden
}

# --- 7. DYNAMIC UI GENERATION ---
function Build-PJob($parentPath, $anchorFile, $gpJob) {
    $tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -ItemType Directory -Path $tempWork | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($anchorFile.FullName, $tempWork)

    $activeSlots = New-Object System.Collections.ArrayList
    $projPath   = Join-Path $tempWork "Metadata\project_settings.config"
    $modSetPath = Join-Path $tempWork "Metadata\model_settings.config"

    $SlotMap = [ordered]@{}; $UsedSlots = New-Object System.Collections.Generic.HashSet[string]
    $UsedSlots.Add("1") | Out-Null

    if (Test-Path $projPath) {
        $projContent = [System.IO.File]::ReadAllText($projPath, [System.Text.Encoding]::UTF8)
        if ($projContent -match '(?is)"filament_colou?r"\s*:\s*\[(.*?)\]') {
            $hexMatches = [regex]::Matches($matches[1], '#[0-9a-fA-F]{6,8}')
            $si = 1
            foreach ($m in $hexMatches) {
                $hk = $m.Value.ToUpper()
                if (-not $SlotMap.Contains($hk)) { $SlotMap[$hk] = $si.ToString() }
                $si++
            }
        }
    }

    if (Test-Path $modSetPath) {
        try {
            [xml]$modXml = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
            foreach ($node in $modXml.SelectNodes('//metadata[contains(@key, "extruder")]')) {
                $val = $node.GetAttribute('value')
                if (-not [string]::IsNullOrWhiteSpace($val)) { $UsedSlots.Add($val) | Out-Null }
            }
        } catch {}
    }

    # Also check 3dmodel.model for materialid assignments
    $modelFile = Get-ChildItem -Path $tempWork -Filter '3dmodel.model' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($modelFile -and (Test-Path $modelFile.FullName)) {
        try {
            $modelContent = [System.IO.File]::ReadAllText($modelFile.FullName, [System.Text.Encoding]::UTF8)
            $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
            foreach ($m in $matMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
        } catch {}
    }

    foreach ($hex in $SlotMap.Keys) {
        $slotId = $SlotMap[$hex]
        if ($UsedSlots.Contains($slotId)) {
            $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
            $matchedName = if ($HexToName.Contains($checkHex)) { $HexToName[$checkHex] } else { "" }
            $activeSlots.Add([PSCustomObject]@{ OldHex = $checkHex; Name = $matchedName }) | Out-Null
        }
    }
    if ($activeSlots.Count -gt 4) { $activeSlots = $activeSlots[0..3] }

    $pJob = @{
        FolderPath = $parentPath; AnchorFile = $anchorFile; TempWork = $tempWork
        ProcessedAnchorPath = ""; CustomImagePath = $null
        UISlots = New-Object System.Collections.ArrayList
        FileRows = New-Object System.Collections.ArrayList
        IsDone = $false; IsQueued = $false; HasCollision = $false
    }

    # ── Outer row border (the RowPanel equivalent) ───────────────────────────
    $pBorder = New-Object System.Windows.Controls.Border
    $pBorder.Background = Get-WpfColor "#1A1C22"; $pBorder.BorderBrush = Get-WpfColor "#2A2C35"
    $pBorder.BorderThickness = New-Object System.Windows.Thickness(1)
    $pBorder.Margin = New-Object System.Windows.Thickness(0,0,0,10); $pBorder.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $pBorder.Padding = New-Object System.Windows.Thickness(10)
    $pJob.RowPanel = $pBorder

    $pGrid = New-Object System.Windows.Controls.Grid
    $pGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))
    $pGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $pBorder.Child = $pGrid

    # ── LEFT: Card panel + Pick panel (with overlays) ────────────────────────
    $leftStack = New-Object System.Windows.Controls.StackPanel
    $leftStack.Orientation = "Horizontal"; $leftStack.VerticalAlignment = "Top"
    [System.Windows.Controls.Grid]::SetColumn($leftStack, 0); $pGrid.Children.Add($leftStack) | Out-Null

    # Card panel — use a Grid so we can stack overlapping layers
    $cardGrid = New-Object System.Windows.Controls.Grid
    $cardGrid.Width = 350; $cardGrid.Height = 350
    $cardGrid.Margin = New-Object System.Windows.Thickness(0,0,15,0)

    $pbModel = New-Object System.Windows.Controls.Image
    $pbModel.Stretch = "Uniform"
    $baseImgPath = Join-Path $tempWork "Metadata\plate_1.png"
    $pbModel.Source = Load-WpfImage $baseImgPath
    $pbModel.AllowDrop = $true
    $cardGrid.Children.Add($pbModel) | Out-Null
    $pJob.PbPlate = $pbModel

    # PbPlateFinished — shown after job completes, covers whole card
    $pbPlateFinished = New-Object System.Windows.Controls.Image
    $pbPlateFinished.Stretch = "Uniform"; $pbPlateFinished.Visibility = "Collapsed"
    $cardGrid.Children.Add($pbPlateFinished) | Out-Null
    $pJob.PbPlateFinished = $pbPlateFinished

    # [CURRENT] gcode thumbnail (small, top-left corner)
    $currentGrid = New-Object System.Windows.Controls.Grid
    $currentGrid.HorizontalAlignment = "Left"; $currentGrid.VerticalAlignment = "Top"
    $currentGrid.Width = 110; $currentGrid.Height = 125; $currentGrid.Margin = New-Object System.Windows.Thickness(8,8,0,0)
    $pbCurrent = New-Object System.Windows.Controls.Image; $pbCurrent.Stretch = "Uniform"
    $lblCurrentTag = Create-TextBlock "[CURRENT]" "#E8A135" 8 "Bold"
    $lblCurrentTag.VerticalAlignment = "Top"; $lblCurrentTag.HorizontalAlignment = "Left"
    $currentGrid.Children.Add($pbCurrent) | Out-Null
    $currentGrid.Children.Add($lblCurrentTag) | Out-Null
    $cardGrid.Children.Add($currentGrid) | Out-Null
    $pJob.PbCurrent = $pbCurrent; $pJob.CurrentThumb = $currentGrid

    # Load [CURRENT] gcode thumbnail if available
    $gcodeFile = Get-ChildItem -Path $parentPath -Filter "*Full.gcode.3mf" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($gcodeFile) {
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $entry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
            if ($entry) {
                $gcodeImgPath = Join-Path $tempWork "plate_1_gcode.png"
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $gcodeImgPath, $true)
                $pbCurrent.Source = Load-WpfImage $gcodeImgPath
            } else { $currentGrid.Visibility = "Collapsed" }
            # --- OVERLAY LABELS ON IMAGE CARD ---
    $lblCharCard = Create-TextBlock "" "#E8A135" 16 "Bold"
    $lblCharCard.HorizontalAlignment = "Right"; $lblCharCard.VerticalAlignment = "Top"
    $lblCharCard.Margin = New-Object System.Windows.Thickness(0, 10, 10, 0)
    $cardGrid.Children.Add($lblCharCard) | Out-Null
    $pJob.LblCharCard = $lblCharCard

    $lblSkipTime = Create-TextBlock "Skip Time: 00" "#FFFFFF" 12 "Bold"
    $lblSkipTime.HorizontalAlignment = "Left"; $lblSkipTime.VerticalAlignment = "Bottom"
    $lblSkipTime.Margin = New-Object System.Windows.Thickness(10, 0, 0, 10)
    $cardGrid.Children.Add($lblSkipTime) | Out-Null

# --- COLOR SLOTS (OVERLAID ON RIGHT EDGE OF IMAGE) ---
    $colorsOverlayStack = New-Object System.Windows.Controls.StackPanel
    $colorsOverlayStack.Orientation = "Vertical"
    $colorsOverlayStack.HorizontalAlignment = "Right"
    $colorsOverlayStack.VerticalAlignment = "Center"
    $colorsOverlayStack.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)

    $slotIdx = 1
    foreach ($slotData in $activeSlots) {
        $rowStack = New-Object System.Windows.Controls.StackPanel
        $rowStack.Orientation = "Horizontal"; $rowStack.HorizontalAlignment = "Right"
        $rowStack.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)

        $textStack = New-Object System.Windows.Controls.StackPanel
        $textStack.Orientation = "Vertical"; $textStack.VerticalAlignment = "Center"
        $textStack.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)

        $lblStatus = Create-TextBlock "" "#A0A0A0" 10 "Bold"
        $lblStatus.HorizontalAlignment = "Right"
        if ($slotData.Name) { $lblStatus.Text = "[MATCHED]"; $lblStatus.Foreground = Get-WpfColor "#4CAF72" }
        else { $lblStatus.Text = "[UNMATCHED]"; $lblStatus.Foreground = Get-WpfColor "#D95F5F" }

        $combo = New-Object System.Windows.Controls.ComboBox; $combo.Width = 140; $combo.IsEditable = $true
        foreach ($k in $LibraryColors.Keys) { [void]$combo.Items.Add($k) }
        if ($slotData.Name) { $combo.Text = $slotData.Name } else { $combo.Text = "Select Color..." }

        $textStack.Children.Add($lblStatus) | Out-Null; $textStack.Children.Add($combo) | Out-Null

        $swatchColor = if ([string]::IsNullOrWhiteSpace($slotData.OldHex)) { "#333333" } else { $slotData.OldHex }
        $swatchBorder = New-Object System.Windows.Controls.Border
        $swatchBorder.Width = 40; $swatchBorder.Height = 40
        $swatchBorder.Background = Get-WpfColor $swatchColor
        $swatchBorder.BorderBrush = Get-WpfColor "#2A2C35"; $swatchBorder.BorderThickness = New-Object System.Windows.Thickness(1)

        # Calculate contrast for the number (Black text on light colors, White text on dark)
        $r = [Convert]::ToInt32($swatchColor.Substring(1,2), 16)
        $g = [Convert]::ToInt32($swatchColor.Substring(3,2), 16)
        $b = [Convert]::ToInt32($swatchColor.Substring(5,2), 16)
        $numColor = if ((0.299*$r + 0.587*$g + 0.114*$b) -gt 128) { "#000000" } else { "#FFFFFF" }

        $lblNum = Create-TextBlock $slotIdx.ToString() $numColor 14 "Bold"
        $lblNum.HorizontalAlignment = "Center"; $lblNum.VerticalAlignment = "Center"
        $swatchBorder.Child = $lblNum

        $rowStack.Children.Add($textStack) | Out-Null; $rowStack.Children.Add($swatchBorder) | Out-Null
        $colorsOverlayStack.Children.Add($rowStack) | Out-Null

        $pJob.UISlots.Add([PSCustomObject]@{ OldHex = $slotData.OldHex; Combo = $combo; StatusLbl = $lblStatus; SwatchBorder = $swatchBorder; LblNum = $lblNum }) | Out-Null

        $combo.Tag = @{ StatusLbl = $lblStatus; OrigName = $slotData.Name; P = $pJob; Swatch = $swatchBorder; LblNum = $lblNum }
        $combo.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, [System.Windows.Controls.TextChangedEventHandler]{
            param($s, $e)
            $data = $s.Tag
            if ($LibraryColors.Contains($s.Text)) {
                $newHex = $LibraryColors[$s.Text]
                $data.Swatch.Background = Get-WpfColor $newHex

                # Update text contrast when color changes
                $r = [Convert]::ToInt32($newHex.Substring(1,2), 16)
                $g = [Convert]::ToInt32($newHex.Substring(3,2), 16)
                $b = [Convert]::ToInt32($newHex.Substring(5,2), 16)
                $numColor = if ((0.299*$r + 0.587*$g + 0.114*$b) -gt 128) { "#000000" } else { "#FFFFFF" }
                $data.LblNum.Foreground = Get-WpfColor $numColor
            }
            if ($s.Text -eq $data.OrigName) {
                if ($data.OrigName) { $data.StatusLbl.Text = "[MATCHED]"; $data.StatusLbl.Foreground = Get-WpfColor "#4CAF72" }
                else { $data.StatusLbl.Text = "[UNMATCHED]"; $data.StatusLbl.Foreground = Get-WpfColor "#D95F5F" }
            } else { $data.StatusLbl.Text = "[CHANGED]"; $data.StatusLbl.Foreground = Get-WpfColor "#E8A135" }
            Validate-PJob $data.P
        })
        $slotIdx++
    }
    $cardGrid.Children.Add($colorsOverlayStack) | Out-Null
        } catch { $currentGrid.Visibility = "Collapsed" }
        finally { if ($null -ne $zip) { $zip.Dispose() } }
    } else { $currentGrid.Visibility = "Collapsed" }

    # Load custom PNG if present (Explicitly ignore the generated card image!)
    $diParent = [System.IO.DirectoryInfo]::new($parentPath)
    $customPng = Get-ChildItem -Path $parentPath -Filter "*.png" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" -and $_.Name -notmatch "(?i)^$($diParent.Name)\.png$" } | Select-Object -First 1
    if ($customPng) {
        $pJob.CustomImagePath = $customPng.FullName
        $pbModel.Source = Load-WpfImage $customPng.FullName
    }

    # Processing overlay for card
    $cardOverlay = New-Object System.Windows.Controls.TextBlock
    $cardOverlay.Text = "[ PROCESSING ]"
    $cardOverlay.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220,232,161,53))
    $cardOverlay.Foreground = Get-WpfColor "#000000"
    $cardOverlay.FontSize = 14; $cardOverlay.FontWeight = [System.Windows.FontWeights]::Bold
    $cardOverlay.TextAlignment = "Center"; $cardOverlay.VerticalAlignment = "Center"; $cardOverlay.HorizontalAlignment = "Stretch"
    $cardOverlay.Visibility = "Collapsed"; $cardOverlay.Padding = New-Object System.Windows.Thickness(0,155,0,0)
    $cardGrid.Children.Add($cardOverlay) | Out-Null
    $pJob.ProcessingOverlay = $cardOverlay

    $leftStack.Children.Add($cardGrid) | Out-Null

    # Drag-drop for custom PNG on plate image
    $pbModel.Tag = @{ P = $pJob }
    $pbModel.Add_DragOver({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
            if ($files[0] -match '(?i)\.(png|jpg|jpeg)$') { $e.Effects = [System.Windows.DragDropEffects]::Copy } else { $e.Effects = [System.Windows.DragDropEffects]::None }
        }
        $e.Handled = $true
    })
    $pbModel.Add_Drop({
        param($s, $e)
        $p = $s.Tag.P
        $files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        $dropped = $files[0]
        if ($dropped -match '(?i)\.(png|jpg|jpeg)$') {
            $dest = Join-Path $p.FolderPath (Split-Path $dropped -Leaf)
            if ($dropped -ne $dest) { Copy-Item -Path $dropped -Destination $dest -Force }
            $s.Source = Load-WpfImage $dest
            $p.CustomImagePath = $dest
        }
    })

    # Pick panel (with overlay)
    $pickGrid = New-Object System.Windows.Controls.Grid
    $pickGrid.Width = 350; $pickGrid.Height = 350

    $pbPick = New-Object System.Windows.Controls.Image; $pbPick.Stretch = "Uniform"
    $pickGrid.Children.Add($pbPick) | Out-Null
    $pJob.PbPick = $pbPick

    $pickOverlay = New-Object System.Windows.Controls.TextBlock
    $pickOverlay.Text = "[ PROCESSING ]"
    $pickOverlay.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220,232,161,53))
    $pickOverlay.Foreground = Get-WpfColor "#000000"
    $pickOverlay.FontSize = 14; $pickOverlay.FontWeight = [System.Windows.FontWeights]::Bold
    $pickOverlay.TextAlignment = "Center"; $pickOverlay.VerticalAlignment = "Center"; $pickOverlay.HorizontalAlignment = "Stretch"
    $pickOverlay.Visibility = "Collapsed"; $pickOverlay.Padding = New-Object System.Windows.Thickness(0,155,0,0)
    $pickGrid.Children.Add($pickOverlay) | Out-Null
    $pJob.PickProcessingOverlay = $pickOverlay

    # Load pick image
    if ($gcodeFile) {
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)(^|/)pick_1\.png$" } | Select-Object -First 1
            if ($pickEntry) {
                $rawPickPath = Join-Path $tempWork "pick_1_raw.png"
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $rawPickPath, $true)
                $pickPath = Join-Path $tempWork "pick_1.png"
                Invoke-RandomizePickColors $rawPickPath $pickPath | Out-Null
                if (-not (Test-Path $pickPath)) { $pickPath = $rawPickPath }
                $pbPick.Source = Load-WpfImage $pickPath
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    }
    $leftStack.Children.Add($pickGrid) | Out-Null

    # ── RIGHT: Controls column ───────────────────────────────────────────────
    $rightStack = New-Object System.Windows.Controls.StackPanel
    $rightStack.Orientation = "Vertical"; $rightStack.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($rightStack, 1); $pGrid.Children.Add($rightStack) | Out-Null

    # Header row: folder label + Refresh + Remove Folder
    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))

    $lblFolder = Create-TextBlock "Folder: $(Split-Path $parentPath -Leaf)" "#FFFFFF" 14 "Bold"
    [System.Windows.Controls.Grid]::SetColumn($lblFolder, 0); $headerGrid.Children.Add($lblFolder) | Out-Null
    $pJob.LblFolder = $lblFolder

    $btnHdrStack = New-Object System.Windows.Controls.StackPanel; $btnHdrStack.Orientation = "Horizontal"
    [System.Windows.Controls.Grid]::SetColumn($btnHdrStack, 1)

    $btnRefresh = New-Object System.Windows.Controls.Button
    $btnRefresh.Content = "Refresh"; $btnRefresh.Background = Get-WpfColor "#2A2C35"; $btnRefresh.Foreground = Get-WpfColor "#FFFFFF"
    $btnRefresh.Width = 100; $btnRefresh.Height = 25; $btnRefresh.BorderThickness = 0
    $btnRefresh.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnRefresh.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRefresh.Tag = @{ P = $pJob; G = $gpJob }
    $btnRefresh.Add_Click({ $t = $this.Tag; Refresh-PJob $t.P $t.G })
    $btnHdrStack.Children.Add($btnRefresh) | Out-Null

    $btnRemoveP = New-Object System.Windows.Controls.Button
    $btnRemoveP.Content = "Remove Folder"; $btnRemoveP.Background = Get-WpfColor "#D95F5F"; $btnRemoveP.Foreground = Get-WpfColor "#FFFFFF"
    $btnRemoveP.Width = 120; $btnRemoveP.Height = 25; $btnRemoveP.BorderThickness = 0; $btnRemoveP.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRemoveP.Tag = @{ P = $pJob; G = $gpJob; Border = $pBorder }
    $btnRemoveP.Add_Click({
        $t = $this.Tag
        $t.G.ParentListStack.Children.Remove($t.Border) | Out-Null
        $t.G.Parents.Remove($t.P) | Out-Null
        if ($t.G.Parents.Count -eq 0) {
            $mainStack.Children.Remove($t.G.Container) | Out-Null
            $script:jobs.Remove($t.G) | Out-Null
        }
    })
    $btnHdrStack.Children.Add($btnRemoveP) | Out-Null

    $headerGrid.Children.Add($btnHdrStack) | Out-Null
    $rightStack.Children.Add($headerGrid) | Out-Null

    # Tasks box
    $tasksBox = New-Object System.Windows.Controls.Border
    $tasksBox.Background = Get-WpfColor "#1C1D23"; $tasksBox.BorderBrush = Get-WpfColor "#2A2C35"
    $tasksBox.BorderThickness = New-Object System.Windows.Thickness(1)
    $tasksBox.Margin = New-Object System.Windows.Thickness(0,10,0,0); $tasksBox.Padding = New-Object System.Windows.Thickness(10)

    $tasksOuter = New-Object System.Windows.Controls.StackPanel

    $tasksRow1 = New-Object System.Windows.Controls.StackPanel; $tasksRow1.Orientation = "Horizontal"
    $chkMerge   = New-Object System.Windows.Controls.CheckBox; $chkMerge.Content   = "Merge";                $chkMerge.IsChecked   = $true; $chkMerge.Foreground   = Get-WpfColor "#FFFFFF"; $chkMerge.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkSlice   = New-Object System.Windows.Controls.CheckBox; $chkSlice.Content   = "Slice / Export Gcode"; $chkSlice.IsChecked   = $true; $chkSlice.Foreground   = Get-WpfColor "#FFFFFF"; $chkSlice.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkExtract = New-Object System.Windows.Controls.CheckBox; $chkExtract.Content = "Extract Data";         $chkExtract.IsChecked = $true; $chkExtract.Foreground = Get-WpfColor "#FFFFFF"; $chkExtract.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkImage   = New-Object System.Windows.Controls.CheckBox; $chkImage.Content   = "Generate Image Card";  $chkImage.IsChecked   = $true; $chkImage.Foreground   = Get-WpfColor "#FFFFFF"
    $tasksRow1.Children.Add($chkMerge) | Out-Null; $tasksRow1.Children.Add($chkSlice) | Out-Null
    $tasksRow1.Children.Add($chkExtract) | Out-Null; $tasksRow1.Children.Add($chkImage) | Out-Null
    $tasksOuter.Children.Add($tasksRow1) | Out-Null

    $tasksRow2 = New-Object System.Windows.Controls.StackPanel
    $tasksRow2.Orientation = "Horizontal"; $tasksRow2.Margin = New-Object System.Windows.Thickness(0,8,0,0)
    $btnSelAll = New-Object System.Windows.Controls.Button; $btnSelAll.Content = "Select All"; $btnSelAll.Background = Get-WpfColor "#2A2C35"; $btnSelAll.Foreground = Get-WpfColor "#FFFFFF"; $btnSelAll.Width = 100; $btnSelAll.Height = 25; $btnSelAll.BorderThickness = 0; $btnSelAll.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnSelAll.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnDeselAll = New-Object System.Windows.Controls.Button; $btnDeselAll.Content = "Deselect All"; $btnDeselAll.Background = Get-WpfColor "#2A2C35"; $btnDeselAll.Foreground = Get-WpfColor "#FFFFFF"; $btnDeselAll.Width = 100; $btnDeselAll.Height = 25; $btnDeselAll.BorderThickness = 0; $btnDeselAll.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnDeselAll.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRevertMerge = New-Object System.Windows.Controls.Button; $btnRevertMerge.Content = "Revert Merge"; $btnRevertMerge.Background = Get-WpfColor "#D95F5F"; $btnRevertMerge.Foreground = Get-WpfColor "#FFFFFF"; $btnRevertMerge.Width = 110; $btnRevertMerge.Height = 25; $btnRevertMerge.BorderThickness = 0; $btnRevertMerge.Cursor = [System.Windows.Input.Cursors]::Hand
    $tasksRow2.Children.Add($btnSelAll) | Out-Null; $tasksRow2.Children.Add($btnDeselAll) | Out-Null; $tasksRow2.Children.Add($btnRevertMerge) | Out-Null
    $tasksOuter.Children.Add($tasksRow2) | Out-Null

    $tasksBox.Child = $tasksOuter
    $rightStack.Children.Add($tasksBox) | Out-Null

    $pJob.ChkMerge = $chkMerge; $pJob.ChkSlice = $chkSlice; $pJob.ChkExtract = $chkExtract; $pJob.ChkImage = $chkImage

    # Checkbox interdependencies
    $tasksData = @{ Merge = $chkMerge; Slice = $chkSlice; Extract = $chkExtract; Image = $chkImage; PJob = $pJob; GpJob = $gpJob }

    $chkSlice.Tag = $tasksData
    $chkSlice.Add_Checked({ if ($this.IsChecked) { $this.Tag.Extract.IsChecked = $true } })

    $chkImage.Tag = $tasksData
    $chkImage.Add_Checked({
        $t = $this.Tag
        $t.Extract.IsChecked = $true
        $tsvExists = (Get-ChildItem -Path $t.PJob.FolderPath -Filter "*_Data.tsv" -ErrorAction SilentlyContinue).Count -gt 0
        if (-not $tsvExists) { $t.Slice.IsChecked = $true }
    })

    $chkExtract.Tag = $tasksData
    $chkExtract.Add_Unchecked({
        $t = $this.Tag
        if (-not $this.IsChecked) {
            $t.Slice.IsChecked = $false
            $t.Image.IsChecked = $false
        }
    })

    $btnSelAll.Tag = $tasksData
    $btnSelAll.Add_Click({
        $t = $this.Tag
        $t.Merge.IsEnabled = $true; $t.Slice.IsEnabled = $true; $t.Extract.IsEnabled = $true; $t.Image.IsEnabled = $true
        $t.Merge.IsChecked = $true; $t.Slice.IsChecked = $true; $t.Extract.IsChecked = $true; $t.Image.IsChecked = $true
    })

    $btnDeselAll.Tag = $tasksData
    $btnDeselAll.Add_Click({
        $t = $this.Tag
        $t.Merge.IsEnabled = $true; $t.Slice.IsEnabled = $true; $t.Extract.IsEnabled = $true; $t.Image.IsEnabled = $true
        $t.Merge.IsChecked = $false; $t.Slice.IsChecked = $false; $t.Extract.IsChecked = $false; $t.Image.IsChecked = $false
    })

    $btnRevertMerge.Tag = @{ P = $pJob; G = $gpJob }
    $btnRevertMerge.Add_Click({
        $t = $this.Tag; $pj = $t.P; $gp = $t.G
        $batPath = Join-Path $scriptDir "RevertMerge.bat"
        if (-not (Test-Path $batPath)) { [System.Windows.MessageBox]::Show("RevertMerge.bat not found.", "Error") | Out-Null; return }
        $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }
        $pj.BtnApply.Content = "Reverting..."; $pj.BtnApply.Width = 150
        if ($pj.BtnRevertDone) { $pj.BtnRevertDone.Visibility = "Collapsed" }
        $pj.RowPanel.IsEnabled = $false
        try {
            $argList = '/c ""' + $batPath + '" "' + $targetPath + '""'
            $proc = Start-Process -FilePath "cmd.exe" -ArgumentList $argList -WindowStyle Hidden -PassThru
            $timeout = 100
            while (-not $proc.HasExited -and $timeout -gt 0) {
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
                Start-Sleep -Milliseconds 100; $timeout--
            }
            if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
        } catch {}
        Refresh-PJob $pj $gp
    })

    # Edit boxes
    $gpName = if ($gpJob.DiGrand) { $gpJob.DiGrand.Name } else { "" }
    $fills = SmartFill $anchorFile.Name $gpName

    $editBox = New-Object System.Windows.Controls.Border
    $editBox.Background = Get-WpfColor "#1C1D23"; $editBox.BorderBrush = Get-WpfColor "#2A2C35"
    $editBox.BorderThickness = New-Object System.Windows.Thickness(1)
    $editBox.Margin = New-Object System.Windows.Thickness(0,10,0,0); $editBox.Padding = New-Object System.Windows.Thickness(10)
    $editStack = New-Object System.Windows.Controls.StackPanel; $editStack.Orientation = "Horizontal"

    $charStack = New-Object System.Windows.Controls.StackPanel; $charStack.Margin = New-Object System.Windows.Thickness(0,0,20,0)
    $charStack.Children.Add((Create-TextBlock "Character *" "#A0A0A0" 12 "Normal")) | Out-Null
    $tbChar = New-Object System.Windows.Controls.TextBox; $tbChar.Text = $fills.Char; $tbChar.Width = 200; $tbChar.Background = Get-WpfColor "#1E2028"; $tbChar.Foreground = Get-WpfColor "#FFFFFF"
    $charStack.Children.Add($tbChar) | Out-Null; $editStack.Children.Add($charStack) | Out-Null
    $pJob.TBChar = $tbChar

    $adjStack = New-Object System.Windows.Controls.StackPanel
    $adjStack.Children.Add((Create-TextBlock "Adjective (Optional)" "#A0A0A0" 12 "Normal")) | Out-Null
    $tbAdj = New-Object System.Windows.Controls.TextBox; $tbAdj.Text = $fills.Adj; $tbAdj.Width = 200; $tbAdj.Background = Get-WpfColor "#1E2028"; $tbAdj.Foreground = Get-WpfColor "#FFFFFF"
    $adjStack.Children.Add($tbAdj) | Out-Null; $editStack.Children.Add($adjStack) | Out-Null
    $pJob.TBAdj = $tbAdj
    $editBox.Child = $editStack; $rightStack.Children.Add($editBox) | Out-Null

    # Files list
    $pnlFiles = New-Object System.Windows.Controls.StackPanel; $pnlFiles.Margin = New-Object System.Windows.Thickness(0,10,0,0)
    $rightStack.Children.Add($pnlFiles) | Out-Null; $pJob.PnlFiles = $pnlFiles
    $files = Get-ChildItem -Path $parentPath -File -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($fi in $files) { Add-FileRow $pJob $gpJob $fi }

    # Apply + Revert Done buttons
    $applyRow = New-Object System.Windows.Controls.StackPanel; $applyRow.Orientation = "Horizontal"; $applyRow.HorizontalAlignment = "Right"; $applyRow.Margin = New-Object System.Windows.Thickness(0,15,0,0)

    $btnApply = New-Object System.Windows.Controls.Button
    $btnApply.Content = "Add to Queue"; $btnApply.Background = Get-WpfColor "#4CAF72"; $btnApply.Foreground = Get-WpfColor "#FFFFFF"
    $btnApply.FontWeight = [System.Windows.FontWeights]::Bold; $btnApply.Width = 150; $btnApply.Height = 35; $btnApply.BorderThickness = 0; $btnApply.Cursor = [System.Windows.Input.Cursors]::Hand
    $applyRow.Children.Add($btnApply) | Out-Null; $pJob.BtnApply = $btnApply

    $btnRevertDone = New-Object System.Windows.Controls.Button
    $btnRevertDone.Content = "REVERT"; $btnRevertDone.Background = Get-WpfColor "#D95F5F"; $btnRevertDone.Foreground = Get-WpfColor "#FFFFFF"
    $btnRevertDone.FontWeight = [System.Windows.FontWeights]::Bold; $btnRevertDone.Width = 75; $btnRevertDone.Height = 35; $btnRevertDone.BorderThickness = 0
    $btnRevertDone.Margin = New-Object System.Windows.Thickness(10,0,0,0); $btnRevertDone.Visibility = "Collapsed"; $btnRevertDone.Cursor = [System.Windows.Input.Cursors]::Hand
    $applyRow.Children.Add($btnRevertDone) | Out-Null; $pJob.BtnRevertDone = $btnRevertDone
    $rightStack.Children.Add($applyRow) | Out-Null

    $btnApply.Tag = @{ P = $pJob; G = $gpJob }
    $btnApply.Add_Click({
        $t = $this.Tag
        if ($this.Content -eq "KEEP") {
            $this.Content = "Finished"; $this.Background = Get-WpfColor "#333333"; $this.IsEnabled = $false; $this.Width = 150
            if ($t.P.BtnRevertDone) { $t.P.BtnRevertDone.Visibility = "Collapsed" }
        } else {
            Enqueue-PJob $t.P $t.G
        }
    })

    $btnRevertDone.Tag = @{ P = $pJob; G = $gpJob }
    $btnRevertDone.Add_Click({
        $t = $this.Tag; $pj = $t.P; $gp = $t.G
        $batPath = Join-Path $scriptDir "RevertMerge.bat"
        if (-not (Test-Path $batPath)) { return }
        $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }
        $pj.BtnApply.Content = "Reverting..."; $pj.BtnApply.Width = 150; $pj.BtnRevertDone.Visibility = "Collapsed"
        $pj.RowPanel.IsEnabled = $false
        try {
            $argList = '/c ""' + $batPath + '" "' + $targetPath + '""'
            $proc = Start-Process -FilePath "cmd.exe" -ArgumentList $argList -WindowStyle Hidden -PassThru
            $timeout = 100
            while (-not $proc.HasExited -and $timeout -gt 0) {
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
                Start-Sleep -Milliseconds 100; $timeout--
            }
            if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
        } catch {}
        Refresh-PJob $pj $gp
    })

    $tbChar.Tag = @{ P = $pJob; G = $gpJob }; $tbChar.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })
    $tbAdj.Tag = @{ P = $pJob; G = $gpJob };  $tbAdj.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })

    Update-ParentPreview $pJob $gpJob
    $gpJob.ParentListStack.Children.Add($pBorder) | Out-Null
    return $pJob
}

function Build-GpJob($gpPath, $parentDict) {
    $diGrand = if ($gpPath -notlike "ROOT_*") { [System.IO.DirectoryInfo]::new($gpPath) } else { $null }
    $gpName = if ($diGrand) { $diGrand.Name } else { "(No Parent Folder)" }

    $gpJob = @{ GpPath = $gpPath; DiGrand = $diGrand; Parents = New-Object System.Collections.ArrayList }
    $script:jobs.Add($gpJob) | Out-Null

    $container = New-Object System.Windows.Controls.Border
    $container.Background = Get-WpfColor "#1C1D23"; $container.BorderBrush = Get-WpfColor "#2A2C35"
    $container.BorderThickness = New-Object System.Windows.Thickness(1)
    $container.Margin = New-Object System.Windows.Thickness(0,0,0,20); $container.CornerRadius = New-Object System.Windows.CornerRadius(6)
    $gpJob.Container = $container

    $gpStack = New-Object System.Windows.Controls.StackPanel; $container.Child = $gpStack

    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.Background = Get-WpfColor "#2A2C35"; $headerGrid.Height = 60

    $headerStack = New-Object System.Windows.Controls.StackPanel; $headerStack.Orientation = "Horizontal"
    $lblGP = Create-TextBlock "Grandparent Theme: " "#E8A135" 14 "Bold"
    $lblGP.Margin = New-Object System.Windows.Thickness(15,0,0,0); $headerStack.Children.Add($lblGP) | Out-Null

    $tbTheme = New-Object System.Windows.Controls.TextBox; $tbTheme.Text = $gpName; $tbTheme.Width = 250
    $tbTheme.Background = Get-WpfColor "#1E2028"; $tbTheme.Foreground = Get-WpfColor "#FFFFFF"
    $tbTheme.VerticalAlignment = "Center"; $tbTheme.Margin = New-Object System.Windows.Thickness(10,0,0,0)
    $headerStack.Children.Add($tbTheme) | Out-Null; $gpJob.TBTheme = $tbTheme

    $chkSkip = New-Object System.Windows.Controls.CheckBox; $chkSkip.Content = "Don't rename folder"
    $chkSkip.Foreground = Get-WpfColor "#FFFFFF"; $chkSkip.VerticalAlignment = "Center"
    $chkSkip.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    $headerStack.Children.Add($chkSkip) | Out-Null; $gpJob.ChkSkip = $chkSkip

    $headerGrid.Children.Add($headerStack) | Out-Null

    $btnRemoveGp = New-Object System.Windows.Controls.Button
    $btnRemoveGp.Content = "Remove Group"; $btnRemoveGp.Background = Get-WpfColor "#D95F5F"; $btnRemoveGp.Foreground = Get-WpfColor "#FFFFFF"
    $btnRemoveGp.Width = 140; $btnRemoveGp.Height = 30; $btnRemoveGp.BorderThickness = 0
    $btnRemoveGp.HorizontalAlignment = "Right"; $btnRemoveGp.Margin = New-Object System.Windows.Thickness(0,0,15,0); $btnRemoveGp.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRemoveGp.Tag = @{ Container = $container; GpJob = $gpJob }
    $btnRemoveGp.Add_Click({
        $t = $this.Tag
        $mainStack.Children.Remove($t.Container) | Out-Null
        $script:jobs.Remove($t.GpJob) | Out-Null
    })
    $headerGrid.Children.Add($btnRemoveGp) | Out-Null
    $gpStack.Children.Add($headerGrid) | Out-Null

    $parentListStack = New-Object System.Windows.Controls.StackPanel
    $parentListStack.Margin = New-Object System.Windows.Thickness(15)
    $gpStack.Children.Add($parentListStack) | Out-Null
    $gpJob.ParentListStack = $parentListStack

    foreach ($pKey in $parentDict.Keys) {
        $pJob = Build-PJob $pKey $parentDict[$pKey] $gpJob
        $gpJob.Parents.Add($pJob) | Out-Null
    }

    $tbTheme.Tag = $gpJob
    $tbTheme.Add_TextChanged({ foreach ($p in $this.Tag.Parents) { Update-ParentPreview $p $this.Tag } })
    $mainStack.Children.Add($container) | Out-Null
}

# --- 8. WPF DISPATCHER TIMER ---
$script:queueTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:queueTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:queueTimer.Add_Tick({
    if ($script:activeProcess -ne $null) {
        if (-not $script:activeProcess.HasExited) {
            # Poll status file for progress text
            $pJob = $script:activeProcessJob.PJob
            $statusFile = Join-Path $pJob.FolderPath "AsyncWorker_Status.txt"
            if (Test-Path $statusFile) {
                try {
                    $fs = [System.IO.File]::Open($statusFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $sr = New-Object System.IO.StreamReader($fs)
                    $statusText = $sr.ReadToEnd(); $sr.Dispose(); $fs.Dispose()
                    if ($statusText) {
                        $txt = "[ $($statusText.Trim()) ]"
                        $pJob.ProcessingOverlay.Text = $txt
                        $pJob.PickProcessingOverlay.Text = $txt
                        $pJob.BtnApply.Content = $statusText.Trim()
                    }
                } catch {}
            }
        } else {
            # Job finished
            $pJob  = $script:activeProcessJob.PJob
            $gpJob = $script:activeProcessJob.GpJob

            $dir      = $pJob.FolderPath
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pJob.ProcessedAnchorPath)
            $gcodeFile = Join-Path $dir "$($baseName).gcode.3mf"
            if (-not (Test-Path $gcodeFile)) { $gcodeFile = $pJob.ProcessedAnchorPath }

            # Try to load the finished plate thumbnail from the new gcode file
            $tempExtract = Join-Path $env:TEMP "finalRead_$([guid]::NewGuid().ToString().Substring(0,8))"
            New-Item -ItemType Directory -Path $tempExtract -Force | Out-Null

            if (Test-Path $gcodeFile) {
                $retry = 0; $zip = $null
                while ($retry -lt 10 -and $null -eq $zip) {
                    try { $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile) }
                    catch { $retry++; Start-Sleep -Milliseconds 200 }
                }

                if ($null -ne $zip) {
                    try {
                        $plateEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
                        if ($plateEntry) {
                            $newPlate = Join-Path $tempExtract "plate_1.png"
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($plateEntry, $newPlate, $true)
                            $bmp = Load-WpfImage $newPlate
                            $pJob.PbPlateFinished.Source = $bmp
                            $pJob.PbPlate.Source = $bmp
                            $pJob.PbPlateFinished.Visibility = "Visible"
                        }

                        $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)(^|/)pick_1\.png$" } | Select-Object -First 1
                        if ($pickEntry) {
                            $rawPick = Join-Path $tempExtract "pick_1_raw.png"
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $rawPick, $true)
                            $newPick = Join-Path $tempExtract "pick_1.png"
                            Invoke-RandomizePickColors $rawPick $newPick | Out-Null
                            if (-not (Test-Path $newPick)) { $newPick = $rawPick }
                            $pJob.PbPick.Source = Load-WpfImage $newPick
                        }

                        # Verify color slots
                        foreach ($slot in $pJob.UISlots) {
                            $selectedName = $slot.Combo.Text
                            if ($LibraryColors.Contains($selectedName)) {
                                $verifiedHex = $LibraryColors[$selectedName]
                                $slot.SwatchBorder.Background = Get-WpfColor $verifiedHex
                                $slot.StatusLbl.Text = "[VERIFIED]"; $slot.StatusLbl.Foreground = Get-WpfColor "#4CAF72"
                            }
                        }
                    } finally { $zip.Dispose() }
                }
            }

            # Reload file rows
            $pJob.PnlFiles.Children.Clear(); $pJob.FileRows.Clear()
            $files = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue | Sort-Object Name
            foreach ($fi in $files) { Add-FileRow $pJob $gpJob $fi }
            Update-ParentPreview $pJob $gpJob

            # Update UI to KEEP/REVERT state
            $pJob.ProcessingOverlay.Visibility = "Collapsed"
            $pJob.PickProcessingOverlay.Visibility = "Collapsed"
            $pJob.RowPanel.IsEnabled = $true
            $pJob.IsDone = $true
            $pJob.ChkMerge.IsEnabled = $true; $pJob.ChkSlice.IsEnabled = $true
            $pJob.ChkExtract.IsEnabled = $true; $pJob.ChkImage.IsEnabled = $true

            $pJob.BtnApply.Content = "KEEP"; $pJob.BtnApply.Background = Get-WpfColor "#4CAF72"
            $pJob.BtnApply.IsEnabled = $true; $pJob.BtnApply.Width = 70
            $pJob.BtnRevertDone.Visibility = "Visible"

            $script:activeProcess = $null; $script:activeProcessJob = $null

            # Start next job immediately rather than waiting for the next tick
            if ($script:processQueue.Count -gt 0) { Start-NextProcess }
        }
    } else {
        if ($script:processQueue.Count -gt 0) { Start-NextProcess }
    }
})

# --- 9. TOP BUTTON HANDLERS ---
$btnCombineData.Add_Click({
    $targetDirs = @()
    foreach ($gpJob in $script:jobs) {
        if ($gpJob.GpPath -like "ROOT_*") {
            foreach ($p in $gpJob.Parents) { if (-not $targetDirs.Contains($p.FolderPath)) { $targetDirs += $p.FolderPath } }
        } else {
            if (-not $targetDirs.Contains($gpJob.GpPath)) { $targetDirs += $gpJob.GpPath }
        }
    }
    if ($targetDirs.Count -eq 0) { return }

    $combinedCount = 0
    $clipboardArray = New-Object System.Collections.ArrayList

    foreach ($targetDir in $targetDirs) {
        if (-not (Test-Path $targetDir)) { continue }
        $tsvFiles = Get-ChildItem -Path $targetDir -Filter "*_Data.tsv" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch "(?i)^.*_Design_Data\.tsv$" }
        if ($tsvFiles.Count -eq 0) { continue }

        $folderName = Split-Path $targetDir -Leaf
        $outTsvPath = Join-Path $targetDir "${folderName}_Data.tsv"

        $combined = [ordered]@{}
        foreach ($tsv in $tsvFiles) {
            if ($tsv.FullName -eq $outTsvPath) { continue }
            $line = Get-Content $tsv.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $key = ($line -split "`t")[0]
            $combined[$key] = $line
        }
        if ($combined.Count -gt 0) {
            $combined.Values | Set-Content -Path $outTsvPath -Encoding UTF8
            $combinedCount++
            foreach ($val in $combined.Values) { [void]$clipboardArray.Add($val) }
        }
    }

    if ($clipboardArray.Count -gt 0) {
        try {
            $clipboardText = $clipboardArray.ToArray() -join "`r`n"
            [System.Windows.Clipboard]::SetText($clipboardText)
            [System.Windows.MessageBox]::Show("Combined TSV data for $combinedCount group(s).`n`n$($clipboardArray.Count) rows copied to clipboard!", "Combine Data Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Combined $combinedCount group(s), but clipboard copy failed.", "Partial Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    } else {
        [System.Windows.MessageBox]::Show("No new data found to combine.", "Nothing to Combine", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
})

$btnProcessAll.Add_Click({
    foreach ($gpJob in $script:jobs) {
        foreach ($pJob in $gpJob.Parents) { Enqueue-PJob $pJob $gpJob }
    }
})

# --- MAIN WINDOW DRAG & DROP HANDLER ---
$window.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { $e.Effects = [System.Windows.DragDropEffects]::Copy }
    else { $e.Effects = [System.Windows.DragDropEffects]::None }
    $e.Handled = $true
})

$window.Add_Drop({
    param($s, $e)
    if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
    $dropped = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)

    $lblGlobalTitle.Text = "Scanning dropped folders..."
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    $newFound = @()
    foreach ($p in $dropped) {
        if (Test-Path $p -PathType Container) { $newFound += Get-ChildItem -Path $p -Filter "*Full.3mf" -Recurse -File }
        elseif ($p -match '(?i)Full\.3mf$') { $newFound += Get-Item $p }
    }

    if ($newFound.Count -gt 0) {
        $newGpQueue = [ordered]@{}
        foreach ($f in $newFound) {
            $parentPath = $f.DirectoryName
            $gp = $f.Directory.Parent
            $gpPath = if ($gp) { $gp.FullName } else { "ROOT_" + $parentPath }

            # Skip if this folder is already in the UI
            $exists = $false
            foreach ($j in $script:jobs) { foreach ($parentJob in $j.Parents) { if ($parentJob.FolderPath -eq $parentPath) { $exists = $true; break } } }
            if ($exists) { continue }

            if (-not $newGpQueue.Contains($gpPath)) { $newGpQueue[$gpPath] = [ordered]@{} }
            if (-not $newGpQueue[$gpPath].Contains($parentPath)) { $newGpQueue[$gpPath][$parentPath] = $f }
        }

        foreach ($gpPath in $newGpQueue.Keys) {
            $existingGp = $null
            foreach ($j in $script:jobs) { if ($j.GpPath -eq $gpPath) { $existingGp = $j; break } }

            if ($existingGp) {
                # Append to existing Grandparent group
                foreach ($pKey in $newGpQueue[$gpPath].Keys) {
                    $pJob = Build-PJob $pKey $newGpQueue[$gpPath][$pKey] $existingGp
                    $existingGp.Parents.Add($pJob) | Out-Null
                }
            } else {
                # Create a brand new Grandparent group
                Build-GpJob $gpPath $newGpQueue[$gpPath]
            }
        }
    }

    $lblGlobalTitle.Text = "Queue Dashboard ($($script:jobs.Count) Theme(s) found)"
    if ($script:jobs.Count -gt 0) { Update-GlobalProcessAllStatus }
})
# --- 10. STARTUP ---
$window.Add_Loaded({
    $idx = 1
    foreach ($gpPath in $gpQueue.Keys) {
        $lblGlobalTitle.Text = "Extracting & Analyzing Group $idx of $($gpQueue.Count)..."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Build-GpJob $gpPath $gpQueue[$gpPath]
        $idx++
    }
    $lblGlobalTitle.Text = "Queue Dashboard ($($gpQueue.Count) Theme(s) found)"
    Update-GlobalProcessAllStatus
    $script:queueTimer.Start()
})

$window.Add_Closed({
    $script:queueTimer.Stop()
    if ($script:activeProcess -ne $null -and -not $script:activeProcess.HasExited) {
        Stop-Process -Id $script:activeProcess.Id -Force -ErrorAction SilentlyContinue
    }
    foreach ($gpJob in $script:jobs) {
        foreach ($pJob in $gpJob.Parents) {
            if (Test-Path $pJob.TempWork) { Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
})

$window.ShowDialog() | Out-Null