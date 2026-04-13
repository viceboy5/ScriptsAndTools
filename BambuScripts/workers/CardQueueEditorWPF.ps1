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
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class NativeFolderBrowser {
    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogImpl {}

    [ComImport, Guid("B63EA76D-1F85-456F-A19C-48159EFA858B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemArray {
        void BindToHandler([In] IntPtr pbc, [In] ref Guid bhid, [In] ref Guid riid, out IntPtr ppvOut);
        void GetPropertyStore([In] int flags, [In] ref Guid riid, out IntPtr ppv);
        void GetPropertyDescriptionList([In] ref Guid keyType, [In] ref Guid riid, out IntPtr ppv);
        void GetAttributes([In] int AttribFlags, [In] uint sfgaoMask, out uint psfgaoAttribs);
        void GetCount(out uint pdwNumItems);
        void GetItemAt([In] uint dwIndex, [MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
        void EnumItems(out IntPtr ppenumShellItems);
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem {
        void BindToHandler([In] IntPtr pbc, [In] ref Guid bhid, [In] ref Guid riid, out IntPtr ppv);
        void GetParent([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
        void GetDisplayName([In] uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes([In] uint sfgaoMask, out uint psfgaoAttribs);
        void Compare([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi, [In] uint hint, out int piOrder);
    }

    // Upgraded to IFileOpenDialog to unlock GetResults()
    [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
        [PreserveSig] int Show([In] IntPtr parent);
        void SetFileTypes([In] uint cFileTypes, [In] IntPtr rgFilterSpec);
        void SetFileTypeIndex([In] uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise([In, MarshalAs(UnmanagedType.Interface)] IntPtr pfde, out uint pdwCookie);
        void Unadvise([In] uint dwCookie);
        void SetOptions([In] uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder([In, MarshalAs(UnmanagedType.Interface)] IntPtr psi);
        void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IntPtr psi);
        void GetFolder([MarshalAs(UnmanagedType.Interface)] out IntPtr ppsi);
        void GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IntPtr ppsi);
        void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
        void AddPlace([In, MarshalAs(UnmanagedType.Interface)] IntPtr psi, int fdap);
        void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close([MarshalAs(UnmanagedType.Error)] int hr);
        void SetClientGuid([In] ref Guid guid);
        void ClearClientData();
        void SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
        void GetResults([MarshalAs(UnmanagedType.Interface)] out IShellItemArray ppenum);
        void GetSelectedItems([MarshalAs(UnmanagedType.Interface)] out IShellItemArray ppsai);
    }

    public static string[] ShowDialog(IntPtr ownerHandle, string title) {
        try {
            IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogImpl();
            uint options;
            dialog.GetOptions(out options);

            // Apply FOS_PICKFOLDERS (0x20) | FOS_FORCEFILESYSTEM (0x40) | FOS_ALLOWMULTISELECT (0x200)
            dialog.SetOptions(options | 0x00000020 | 0x00000040 | 0x00000200);
            if (!string.IsNullOrEmpty(title)) dialog.SetTitle(title);

            int hr = dialog.Show(ownerHandle);
            if (hr != 0) return null;

            IShellItemArray resultsArray;
            dialog.GetResults(out resultsArray);

            uint count;
            resultsArray.GetCount(out count);
            string[] paths = new string[count];

            for (uint i = 0; i < count; i++) {
                IShellItem item;
                resultsArray.GetItemAt(i, out item);
                string path;
                item.GetDisplayName(0x80058000, out path); // 0x80058000 = SIGDN_FILESYSPATH
                paths[i] = path;
                Marshal.ReleaseComObject(item);
            }

            Marshal.ReleaseComObject(resultsArray);
            Marshal.ReleaseComObject(dialog);

            return paths;
        } catch { return null; }
    }
}
"@

try {
    Add-Type -TypeDefinition @'
    using System;
    using System.Drawing;
    using System.Collections.Generic;
    public class MapLine  { public Point Start; public Point End; public Color LineColor; }
    public class MapBounds { public Rectangle Bounds; public Color BoxColor; }
    public class FastMergeMap {
        public static List<MapLine> GetMergeLines(Bitmap pre, Bitmap post) {
            var preAnchors  = new Dictionary<int, Point>();
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
        public static List<MapBounds> GetUnmatchedBounds(Bitmap pre, Bitmap post) {
            // Rebuild the same anchor maps and positional links as GetMergeLines
            var preAnchors  = new Dictionary<int, Point>();
            var postAnchors = new Dictionary<int, Point>();
            int w = pre.Width; int h = pre.Height;
            for (int y = 0; y < h; y++)
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
            // A post color is "matched" if a pre-anchor position lands on it
            var matchedPost = new HashSet<int>();
            foreach (var oldPt in preAnchors.Values) {
                if (oldPt.X < post.Width && oldPt.Y < post.Height) {
                    Color c = post.GetPixel(oldPt.X, oldPt.Y);
                    int argb = c.ToArgb();
                    if (postAnchors.ContainsKey(argb)) matchedPost.Add(argb);
                }
            }
            // Any post color never reached by a pre-anchor = unmatched
            var unmatched = new HashSet<int>();
            foreach (int argb in postAnchors.Keys)
                if (!matchedPost.Contains(argb)) unmatched.Add(argb);
            if (unmatched.Count == 0) return new List<MapBounds>();
            var minX = new Dictionary<int,int>(); var minY = new Dictionary<int,int>();
            var maxX = new Dictionary<int,int>(); var maxY = new Dictionary<int,int>();
            foreach (int a in unmatched) { minX[a] = int.MaxValue; minY[a] = int.MaxValue; maxX[a] = 0; maxY[a] = 0; }
            for (int y = 0; y < post.Height; y++)
                for (int x = 0; x < post.Width; x++) {
                    Color c = post.GetPixel(x, y);
                    int a = c.ToArgb();
                    if (!unmatched.Contains(a)) continue;
                    if (x < minX[a]) minX[a] = x; if (y < minY[a]) minY[a] = y;
                    if (x > maxX[a]) maxX[a] = x; if (y > maxY[a]) maxY[a] = y;
                }
            var result = new List<MapBounds>();
            foreach (int a in unmatched)
                result.Add(new MapBounds { Bounds = new Rectangle(minX[a]-2, minY[a]-2, maxX[a]-minX[a]+4, maxY[a]-minY[a]+4), BoxColor = Color.FromArgb(a) });
            return result;
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
    $stem = $anchorName -replace '(?i)\.gcode\.3mf$|\.3mf$|\.stl$|\.png$', ''
    $stem = $stem -replace '(?i)_(Full|Final|Nest)$', ''

    # Strip all leading qualifier tokens (printer prefix + tags) from gpName
    # so we match only the bare theme token against the filename stem.
    # e.g. "P2S_KC_Puppies" -> "Puppies", "KC_Puppies" -> "Puppies"
    $cleanGpName = $gpName
    if ($gpName -ne '') {
        $gpTokens = [System.Collections.Generic.List[string]]($gpName -split '_' | Where-Object { $_ -ne '' })
        while ($gpTokens.Count -gt 1) {
            $head = $gpTokens[0]
            if (($script:PrinterPrefixes -icontains $head) -or ($script:Tags -icontains $head)) {
                $gpTokens.RemoveAt(0)
            } else { break }
        }
        $cleanGpName = $gpTokens -join '_'
    }

    # Try to strip the theme by matching the cleaned gpName at the end of the stem
    $themeStripped = $false
    $prefix = $stem
    if ($cleanGpName -ne '') {
        $escaped = [regex]::Escape($cleanGpName)
        if ($stem -imatch "^(.+)_${escaped}$") { $prefix = $Matches[1]; $themeStripped = $true }
    }

    # Strip known printer prefix (e.g. P2S_, X1C_) so it doesn't bleed into Character/Adjective
    $prefParts = $prefix -split '_'
    if ($prefParts.Count -gt 1 -and $script:PrinterPrefixes -icontains $prefParts[0]) {
        $prefix = ($prefParts[1..($prefParts.Count-1)] -join '_')
    }

    # Positional fallback: if theme wasn't matched via gpName AND we still have 3+ tokens,
    # treat the last token as the theme and strip it.
    # Convention: Prefix_Char_Adj_Theme_Suffix — after stripping prefix+suffix, max 3 remain.
    $parts = [string[]]($prefix -split '_' | Where-Object { $_ -ne '' })
    if (-not $themeStripped -and $parts.Count -ge 3) {
        $parts = $parts[0..($parts.Count-2)]
    }

    # Never let a known theme name appear in Character or Adjective regardless of how
    # it ended up in the token list (e.g. theme not set at the GP level, positional
    # fallback missed, etc.)
    $parts = [string[]]($parts | Where-Object {
        $clean = $_ -replace '[^a-zA-Z0-9]', ''
        -not ($script:GpThemes | Where-Object { ($_ -replace '[^a-zA-Z0-9]', '') -ieq $clean })
    })

    # Detect a known tag prefix at the start of the first part
    # e.g. "KCFrankenstein" -> Tag="KC", Char="Frankenstein"
    $detectedTag = ''
    if ($parts.Count -ge 1) {
        foreach ($t in $script:Tags) {
            if ($parts[0].Length -gt $t.Length -and $parts[0].Substring(0, $t.Length) -ieq $t) {
                $detectedTag = $t
                $parts[0] = $parts[0].Substring($t.Length)
                break
            }
        }
    }

    if ($parts.Count -ge 2) { return @{ Char = $parts[0]; Adj = ($parts[1..($parts.Count-1)] -join ''); Tag = $detectedTag } }
    return @{ Char = ($parts -join ''); Adj = ''; Tag = $detectedTag }
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

function Get-ImageBasedMergeMap {
    param([string]$preMergePath, [string]$postMergePath)
    if (-not (Test-Path $preMergePath) -or -not (Test-Path $postMergePath)) { return $null }
    try {
        $bmpPre  = New-Object System.Drawing.Bitmap($preMergePath)
        $bmpPost = New-Object System.Drawing.Bitmap($postMergePath)
        $lines = [FastMergeMap]::GetMergeLines($bmpPre, $bmpPost)
        $g = [System.Drawing.Graphics]::FromImage($bmpPost)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        foreach ($l in $lines) {
            $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 5)
            $colorPen  = New-Object System.Drawing.Pen($l.LineColor, 2)
            # Outline then color for the connecting line
            $g.DrawLine($borderPen, $l.Start, $l.End)
            $g.DrawLine($colorPen,  $l.Start, $l.End)
            # Crosshairs at each end
            $g.DrawLine($borderPen, ($l.Start.X - 3), $l.Start.Y, ($l.Start.X + 3), $l.Start.Y)
            $g.DrawLine($borderPen, $l.Start.X, ($l.Start.Y - 3), $l.Start.X, ($l.Start.Y + 3))
            $g.DrawLine($colorPen,  ($l.Start.X - 3), $l.Start.Y, ($l.Start.X + 3), $l.Start.Y)
            $g.DrawLine($colorPen,  $l.Start.X, ($l.Start.Y - 3), $l.Start.X, ($l.Start.Y + 3))
            $g.DrawLine($borderPen, ($l.End.X - 3), $l.End.Y, ($l.End.X + 3), $l.End.Y)
            $g.DrawLine($borderPen, $l.End.X, ($l.End.Y - 3), $l.End.X, ($l.End.Y + 3))
            $g.DrawLine($colorPen,  ($l.End.X - 3), $l.End.Y, ($l.End.X + 3), $l.End.Y)
            $g.DrawLine($colorPen,  $l.End.X, ($l.End.Y - 3), $l.End.X, ($l.End.Y + 3))
            $borderPen.Dispose(); $colorPen.Dispose()
        }
        $g.Dispose(); $bmpPre.Dispose()
        return $bmpPost
    } catch { return $null }
}

# --- 4. BACKEND QUEUE DATA ---
$script:jobs = New-Object System.Collections.ArrayList
$script:processQueue = New-Object System.Collections.Queue
$script:activeProcess = $null
$script:activeProcessJob = $null
. "$PSScriptRoot\BambuConfig.ps1"
$script:AdjPresets = @('Common','RARE','EPIC','LEGENDARY','Default')

function Find-AnchorFile($folderPath) {
    $files = Get-ChildItem -Path $folderPath -File -ErrorAction SilentlyContinue
    $f = $files | Where-Object { $_.Name -match '(?i)Full\.3mf$'  -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Name -match '(?i)Final\.3mf$' -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.3mf$'    -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.stl$' }                                               | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.png$'    -and $_.Name -notmatch '(?i)_slicePreview\.png$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    return $f
}

function Get-AnchorQueue($paths) {
    $queue = [ordered]@{}
    foreach ($p in $paths) {
        if (Test-Path $p -PathType Container) {
            $relevantDirs = @($p) + @(Get-ChildItem -Path $p -Directory -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
            foreach ($dir in $relevantDirs) {
                $anchor = Find-AnchorFile $dir
                if ($anchor) {
                    $gp = [System.IO.DirectoryInfo]::new($dir).Parent
                    $gpPath = if ($gp) { $gp.FullName } else { "ROOT_" + $dir }
                    if (-not $queue.Contains($gpPath)) { $queue[$gpPath] = [ordered]@{} }
                    if (-not $queue[$gpPath].Contains($dir)) { $queue[$gpPath][$dir] = $anchor }
                }
            }
        } elseif ($p -match '(?i)\.(3mf|stl|png)$' -and $p -notmatch '(?i)\.gcode\.3mf$' -and $p -notmatch '(?i)_slicePreview\.png$') {
            $f = Get-Item $p -ErrorAction SilentlyContinue
            if ($f) {
                $gp = $f.Directory.Parent
                $gpPath = if ($gp) { $gp.FullName } else { "ROOT_" + $f.DirectoryName }
                if (-not $queue.Contains($gpPath)) { $queue[$gpPath] = [ordered]@{} }
                if (-not $queue[$gpPath].Contains($f.DirectoryName)) { $queue[$gpPath][$f.DirectoryName] = $f }
            }
        }
    }
    return $queue
}

# Check for files passed by the VBScript, otherwise launch empty
$gpQueue = [ordered]@{}
if ($args.Count -gt 0) {
    $gpQueue = Get-AnchorQueue $args
}

# --- 5. THE XAML LAYOUT ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Batch Pre-Flight Editor - WPF Engine"
        Width="1550" Height="850" MinWidth="1100" MinHeight="600"
        Background="#16171B" WindowStartupLocation="CenterScreen" AllowDrop="True">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Background="#1C1D23" Grid.Row="0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>

                <TextBlock Grid.Column="0" Name="LblGlobalTitle" Text="Loading files into queue..." Foreground="White" FontSize="18" FontWeight="Bold" VerticalAlignment="Center" Margin="15,0,0,0"/>

                <StackPanel Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center" Orientation="Vertical">
                    <Button Name="BtnBrowse" Content="Browse Files" Background="#5A78C4" Foreground="White" FontWeight="Bold" Width="140" Height="30" BorderThickness="0" Cursor="Hand"/>
                    <TextBlock Text="Browse or drop files to add" Foreground="#888888" FontSize="10" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="2" HorizontalAlignment="Right" VerticalAlignment="Center" Orientation="Horizontal" Margin="0,0,15,0">
                    <Button Name="BtnProcessAll" Content="Process All Tasks" Background="#4CAF72" Foreground="White" FontWeight="Bold" Width="150" Height="30" BorderThickness="0" Cursor="Hand"/>
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
$btnBrowse      = $window.FindName("BtnBrowse") # <--- ADD THIS LINE BACK
$mainStack      = $window.FindName("MainStack")

function Update-GlobalProcessAllStatus {
    $hasAnyIssue = $false
    foreach ($gp in $script:jobs) {
        foreach ($p in $gp.Parents) {
            if ($p.IsQueued -or $p.IsDone) { continue }
            if ($p.HasCollision) { $hasAnyIssue = $true; break }
            foreach ($slot in $p.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $hasAnyIssue = $true; break } }
            if ($hasAnyIssue) { break }
        }
        if ($hasAnyIssue) { break }
    }
    if ($hasAnyIssue) {
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
    $th = ("$($gpJob.TBTheme.SelectedItem)" -replace '[^a-zA-Z0-9]', '')
    $tg = ''
    if ($null -ne $pJob.TBTag -and $null -ne $pJob.TBTag.SelectedItem -and "$($pJob.TBTag.SelectedItem)" -ne '(none)') {
        $tg = "$($pJob.TBTag.SelectedItem)" -replace '[^a-zA-Z0-9]', ''
    }
    $pf = ""
    if ($null -ne $gpJob.CbPrefix -and $null -ne $gpJob.CbPrefix.SelectedItem) {
        $pf = $gpJob.CbPrefix.SelectedItem.ToString()
        if ($pf -eq "(none)") { $pf = "" }
    }

    # 1. Split CamelCase into spaces (e.g., "BabyDragon" -> "Baby Dragon")
    $chSpaced = $ch -creplace '([a-z])([A-Z])', '$1 $2'
    $adSpaced = $ad -creplace '([a-z])([A-Z])', '$1 $2'

    # 2. Format title: "Tag - Character (Adj)" when tag set, else "Character (Adj)"
    $displayTitle = if ($tg) { "$tg - $chSpaced" } else { $chSpaced }
    if (-not [string]::IsNullOrWhiteSpace($adSpaced)) {
        $displayTitle += " ($adSpaced)"
    }

    # 3. Apply to the UI card in ALL CAPS
    if ($null -ne $pJob.LblCharCard) { $pJob.LblCharCard.Text = $displayTitle.ToUpper() }

    # Tag is prepended directly to Character (no separator) in filenames: Printer_TagChar_Adj_Theme_Suffix
    $combined = "$tg$ch"

    $nameCounts = @{}
    foreach ($r in $pJob.FileRows) {
        $sf = $r.SuffixBox.Text -replace '[^a-zA-Z0-9]', ''
        $parts = New-Object System.Collections.ArrayList
        if ($pf)       { [void]$parts.Add($pf) }
        if ($combined) { [void]$parts.Add($combined) }; if ($ad) { [void]$parts.Add($ad) }
        if ($th)       { [void]$parts.Add($th) }; if ($sf) { [void]$parts.Add($sf) }
        $r.TargetName = ($parts.ToArray() -join '_') + $r.Ext
        if (-not $nameCounts.ContainsKey($r.TargetName)) { $nameCounts[$r.TargetName] = 0 }
        $nameCounts[$r.TargetName]++
    }

    $hasCollision = $false
    foreach ($r in $pJob.FileRows) {
        $r.NewLbl.Inlines.Clear()
        $sfx     = $r.SuffixBox.Text -replace '[^a-zA-Z0-9]', ''
        $sfxPart = if ($sfx) { "_${sfx}$($r.Ext)" } else { $r.Ext }
        $baseLen = $r.TargetName.Length - $sfxPart.Length
        $basePart = if ($baseLen -gt 0) { $r.TargetName.Substring(0, $baseLen) } else { "" }
        if ($nameCounts[$r.TargetName] -gt 1) {
            $run = New-Object System.Windows.Documents.Run($r.TargetName); $run.Foreground = Get-WpfColor "#D95F5F"
            $r.NewLbl.Inlines.Add($run)
            if ($r.OldLbl) { $r.OldLbl.Foreground = Get-WpfColor "#D95F5F" }
            $hasCollision = $true
        } else {
            $runBase = New-Object System.Windows.Documents.Run($basePart); $runBase.Foreground = Get-WpfColor "#90B8C8"
            $runSfx  = New-Object System.Windows.Documents.Run($sfxPart);  $runSfx.Foreground  = Get-WpfColor $r.BaseColor
            $r.NewLbl.Inlines.Add($runBase); $r.NewLbl.Inlines.Add($runSfx)
            if ($r.OldLbl) { $r.OldLbl.Foreground = Get-WpfColor "#6B6E7A" }
        }
    }
    # Update folder label to preview the renamed folder name (Prefix_TagCharacter_Adjective_Theme)
    $folderParts = @()
    if ($pf)       { $folderParts += $pf }
    if ($combined) { $folderParts += $combined }
    if ($ad)       { $folderParts += $ad }
    if ($th)       { $folderParts += $th }
    $previewFolderName = $folderParts -join '_'
    if ($null -ne $pJob.LblFolder) {
        if ($previewFolderName) { $pJob.LblFolder.Text = "Folder: $previewFolderName" }
        else { $pJob.LblFolder.Text = "Folder: $(Split-Path $pJob.FolderPath -Leaf)" }
    }

    # Update grandparent folder name preview (Prefix_Theme or just Theme)
    if ($null -ne $gpJob.LblGpPreview) {
        $gpPreview = if ($pf) { "${pf}_${th}" } else { $th }
        $gpJob.LblGpPreview.Text = if ($gpPreview) { [char]0x2192 + " $gpPreview" } else { "" }
    }

    $pJob.HasCollision = $hasCollision
    Validate-PJob $pJob
    Update-GlobalProcessAllStatus
}

function Add-FileRow($pJob, $gpJob, $fi) {
    $parsed = ParseFile $fi.Name

    $fRow = New-Object System.Windows.Controls.Border
    $fRow.Background = Get-WpfColor "#16171B"; $fRow.BorderBrush = Get-WpfColor "#2A2C35"
    $fRow.BorderThickness = New-Object System.Windows.Thickness(0,0,0,1)
    $fRow.MinHeight = 40; $fRow.Padding = New-Object System.Windows.Thickness(0, 5, 0, 5)

    $fGrid = New-Object System.Windows.Controls.Grid
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(70))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(30))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(45))}))
    $fGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(40))}))
    $fRow.Child = $fGrid

    $sBadge = New-Object System.Windows.Controls.TextBox
    $sBadge.Text = $parsed.Suffix; $sBadge.Background = Get-WpfColor "#1E2028"; $sBadge.Foreground = Get-WpfColor "#E8A135"
    $sBadge.VerticalAlignment = "Center"; $sBadge.Margin = New-Object System.Windows.Thickness(10,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($sBadge, 0); $fGrid.Children.Add($sBadge) | Out-Null

    $lOld = Create-TextBlock $fi.Name "#6B6E7A" 11 "Normal"
    $lOld.Margin = New-Object System.Windows.Thickness(10,0,0,0)
    $lOld.TextWrapping = "Wrap"
    [System.Windows.Controls.Grid]::SetColumn($lOld, 1); $fGrid.Children.Add($lOld) | Out-Null

    $lArr = Create-TextBlock "->" "#A0A0A0" 12 "Normal"
    [System.Windows.Controls.Grid]::SetColumn($lArr, 2); $fGrid.Children.Add($lArr) | Out-Null

    $newNameColor = if     ($fi.Name -imatch 'Nest\.3mf$')        { "#FF69B4" }  # Pink
                   elseif ($fi.Name -imatch 'Full\.gcode\.3mf$') { "#4CAF72" }  # Green
                   elseif ($fi.Name -imatch 'Full\.3mf$')        { "#B57BFF" }  # Purple
                   elseif ($fi.Name -imatch 'Final\.3mf$')       { "#FFD700" }  # Yellow
                   else                                           { "#90B8C8" }  # Blue-grey
    $lNew = Create-TextBlock "" $newNameColor 11 "Bold"
    $lNew.Margin = New-Object System.Windows.Thickness(5,0,5,0)
    $lNew.TextWrapping = "Wrap"
    [System.Windows.Controls.Grid]::SetColumn($lNew, 3); $fGrid.Children.Add($lNew) | Out-Null

    $btnOpen = New-Object System.Windows.Controls.Button
    $btnOpen.Content = "Open"; $btnOpen.Background = Get-WpfColor "#2A2C35"; $btnOpen.Foreground = Get-WpfColor "#A0C4FF"
    $btnOpen.BorderThickness = 0; $btnOpen.Width = 40; $btnOpen.Height = 20; $btnOpen.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnOpen.ToolTip = $fi.FullName
    [System.Windows.Controls.Grid]::SetColumn($btnOpen, 4); $fGrid.Children.Add($btnOpen) | Out-Null
    $btnOpen.Tag = $fi.FullName
    $btnOpen.Add_Click({ Start-Process $this.Tag })

    $btnDel = New-Object System.Windows.Controls.Button
    $btnDel.Content = "X"; $btnDel.Background = Get-WpfColor "#D95F5F"; $btnDel.Foreground = Get-WpfColor "#FFFFFF"
    $btnDel.BorderThickness = 0; $btnDel.Width = 20; $btnDel.Height = 20; $btnDel.Cursor = [System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnDel, 5); $fGrid.Children.Add($btnDel) | Out-Null

    $frObj = [PSCustomObject]@{ OldPath = $fi.FullName; SuffixBox = $sBadge; OldLbl = $lOld; NewLbl = $lNew; Ext = $parsed.Extension; TargetName = ""; BaseColor = $newNameColor }
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
    $folderPath = $pJob.FolderPath

    # Record position so we can reinsert at the same slot
    $idx = $gpJob.ParentListStack.Children.IndexOf($pJob.RowPanel)

    # Clean up old temp work
    if (Test-Path $pJob.TempWork) { Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue }

    # Tear down old UI row and data entry
    $gpJob.ParentListStack.Children.Remove($pJob.RowPanel) | Out-Null
    $gpJob.Parents.Remove($pJob) | Out-Null

    $newAnchor = Find-AnchorFile $folderPath
    if (-not $newAnchor) { Update-GpFileCount $gpJob; return }

    # Rebuild from scratch — re-parses slots, SmartFill, full card UI
    $newPJob = Build-PJob $folderPath $newAnchor $gpJob

    # Reinsert at the original position
    $lastIdx = $gpJob.ParentListStack.Children.Count - 1
    if ($idx -ge 0 -and $idx -lt $lastIdx) {
        $gpJob.ParentListStack.Children.Remove($newPJob.RowPanel) | Out-Null
        $gpJob.ParentListStack.Children.Insert($idx, $newPJob.RowPanel) | Out-Null
        $gpJob.Parents.Insert($idx, $newPJob) | Out-Null
    } else {
        $gpJob.Parents.Add($newPJob) | Out-Null
    }

    Update-GpFileCount $gpJob
}

function Enqueue-PJob($pJob, $gpJob) {
    if ($pJob.IsQueued -or $pJob.IsDone -or $pJob.HasCollision) { return }
    foreach ($slot in $pJob.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { return } }

    # First plate queued for this grandparent: confirm rename if needed, then lock UI
    if (-not $gpJob.GpRenameConfirmed) {
        if (-not $gpJob.ChkSkip.IsChecked) {
            $th = ("$($gpJob.TBTheme.SelectedItem)" -replace '[^a-zA-Z0-9]', '')
            $pf = if ($null -ne $gpJob.CbPrefix -and "$($gpJob.CbPrefix.SelectedItem)" -ne "(none)") { "$($gpJob.CbPrefix.SelectedItem)" } else { "" }
            $newGpFolderName = if ($pf) { "${pf}_${th}" } else { $th }
            $currentLeaf = if ($gpJob.DiGrand) { Split-Path $gpJob.GpPath -Leaf } else { "" }

            if ($newGpFolderName -ne '' -and $currentLeaf -ne '' -and $newGpFolderName -ne $currentLeaf) {
                # Strip printer prefix to get the bare theme part of the current folder name
                $currentThemePart = $currentLeaf
                foreach ($pfx in $script:PrinterPrefixes) {
                    if ($currentLeaf -imatch "^${pfx}_") { $currentThemePart = $currentLeaf.Substring($pfx.Length + 1); break }
                }
                $currentThemeClean = $currentThemePart -replace '[^a-zA-Z0-9]', ''
                $alreadyKnown = [bool]($script:GpThemes | Where-Object { ($_ -replace '[^a-zA-Z0-9]','') -ieq $currentThemeClean })

                if (-not $alreadyKnown) {
                    $res = [System.Windows.MessageBox]::Show(
                        "Are you sure you want to rename:`n`n    $currentLeaf`n    -> $newGpFolderName",
                        "Confirm Folder Rename",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question
                    )
                    if ($res -ne 'Yes') { return }
                }
            }
        }
        $gpJob.GpRenameConfirmed = $true
        if ($null -ne $gpJob.TBTheme)  { $gpJob.TBTheme.IsEnabled  = $false }
        if ($null -ne $gpJob.CbPrefix) { $gpJob.CbPrefix.IsEnabled = $false }
        if ($null -ne $gpJob.ChkSkip)  { $gpJob.ChkSkip.IsEnabled  = $false }
    }

    $pJob.IsQueued = $true
    $pJob.BtnApply.Content = "Queued..."; $pJob.BtnApply.Background = Get-WpfColor "#E8A135"
    $pJob.RowPanel.IsEnabled = $false
    $pJob.CardStatusLabel.Text = "[ PREPARING ]"; $pJob.ProcessingOverlay.Visibility = "Visible"
    $pJob.PickStatusLabel.Text = "[ PREPARING ]"; $pJob.PickProcessingOverlay.Visibility = "Visible"

    $script:processQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob })
}

function Start-NextProcess {
    if ($script:activeProcess -ne $null -or $script:processQueue.Count -eq 0) { return }

    $jobWrapper = $script:processQueue.Dequeue()
    $pJob = $jobWrapper.PJob; $gpJob = $jobWrapper.GpJob
    $script:activeProcessJob = $jobWrapper
    $pJob.BtnApply.Content = "Processing..."; $pJob.IsQueued = $false

    $th = ("$($gpJob.TBTheme.SelectedItem)" -replace '[^a-zA-Z0-9]', '')
    $pf = ""
    if ($null -ne $gpJob.CbPrefix -and $null -ne $gpJob.CbPrefix.SelectedItem) {
        $pf = $gpJob.CbPrefix.SelectedItem.ToString()
        if ($pf -eq "(none)") { $pf = "" }
    }
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

    $doRename  = [bool]$pJob.ChkRename.IsChecked
    $doMerge   = [bool]$pJob.ChkMerge.IsChecked
    $doSlice   = [bool]$pJob.ChkSlice.IsChecked
    $doExtract = [bool]$pJob.ChkExtract.IsChecked
    $doImage   = [bool]$pJob.ChkImage.IsChecked
    $doLogs    = [bool]$pJob.ChkLogs.IsChecked

    $anchorIsZip3mf = $pJob.AnchorFile.Extension -imatch '\.3mf$' -and $pJob.AnchorFile.Name -notmatch '(?i)\.gcode\.3mf$'

    # Patch color changes into the source 3MF before any renaming or merging
    if ($modifiedFiles.Count -gt 0 -and $anchorIsZip3mf) {
        $srcPath = $pJob.AnchorFile.FullName
        if (Test-Path -LiteralPath $srcPath) {
            try {
                $srcZip = [System.IO.Compression.ZipFile]::Open($srcPath, 'Update')
                foreach ($mf in $modifiedFiles) {
                    $rel = $mf.FullName.Substring($pJob.TempWork.Length).TrimStart('\','/').Replace('\','/')
                    $e = $srcZip.GetEntry($rel)
                    if ($e) { $e.Delete() }
                    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($srcZip, $mf.FullName, $rel) | Out-Null
                }
                $srcZip.Dispose()
            } catch {}
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        }
    }

    $newFilePath = $pJob.AnchorFile.FullName

    if ($doRename) {
        # Rename anchor file
        $anchorTargetName = ""; $anchorFileRow = $null
        $currentAnchorLocation = $pJob.AnchorFile.FullName
        foreach ($r in $pJob.FileRows) {
            if ((Split-Path $r.OldPath -Leaf) -eq $pJob.AnchorFile.Name) {
                $anchorTargetName = $r.TargetName; $anchorFileRow = $r; $currentAnchorLocation = $r.OldPath; break
            }
        }
        if ($anchorTargetName -eq "") { $anchorTargetName = $pJob.AnchorFile.Name }
        $newFilePath = Join-Path $pJob.FolderPath $anchorTargetName

        if ($currentAnchorLocation -ne $newFilePath) {
            if (Test-Path $newFilePath) { Remove-Item $newFilePath -Force -ErrorAction SilentlyContinue }
            try { Rename-Item $currentAnchorLocation $anchorTargetName -Force } catch {}
        }
        if ($null -ne $anchorFileRow) { $anchorFileRow.OldPath = $newFilePath }
    }

    # Inject color changes back into the zip (after rename so path is correct)
    if ($modifiedFiles.Count -gt 0 -and $anchorIsZip3mf) {
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

    if ($doRename) {
        Start-Sleep -Milliseconds 100

        # Rename all other files
        foreach ($r in $pJob.FileRows) {
            if ($r.OldPath -eq $pJob.ProcessedAnchorPath) { continue }
            $targetName = $r.TargetName
            $newPath = Join-Path $pJob.FolderPath $targetName
            if ($r.OldPath -ne $newPath -and (Test-Path $r.OldPath)) {
                Rename-Item $r.OldPath $targetName -Force; $r.OldPath = $newPath
            }
        }

        # Rename parent folder
        $cleanChar = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
        $cleanAdj  = $pJob.TBAdj.Text  -replace '[^a-zA-Z0-9]', ''
        $cleanTag  = if ($null -ne $pJob.TBTag -and $null -ne $pJob.TBTag.SelectedItem -and "$($pJob.TBTag.SelectedItem)" -ne '(none)') { "$($pJob.TBTag.SelectedItem)" -replace '[^a-zA-Z0-9]', '' } else { '' }
        $cleanCombined = "$cleanTag$cleanChar"
        $pParts = New-Object System.Collections.ArrayList
        if ($pf)           { $pParts.Add($pf)           | Out-Null }
        if ($cleanCombined){ $pParts.Add($cleanCombined) | Out-Null }
        if ($cleanAdj)     { $pParts.Add($cleanAdj)     | Out-Null }
        if ($th)           { $pParts.Add($th)            | Out-Null }
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
        $newGpFolderName = if ($pf) { "${pf}_${th}" } else { $th }
        if (-not $gpJob.ChkSkip.IsChecked -and $newGpFolderName -ne '' -and $oldGrand -ne '' -and $newGpFolderName -ne (Split-Path $oldGrand -Leaf)) {
            $newGrand = Join-Path (Split-Path $oldGrand -Parent) $newGpFolderName
            try {
                Rename-Item $oldGrand $newGpFolderName -Force -ErrorAction Stop
                $gpJob.GpPath = $newGrand; $gpJob.DiGrand = [System.IO.DirectoryInfo]::new($newGrand)
                foreach ($p in $gpJob.Parents) {
                    $p.FolderPath = $p.FolderPath.Replace($oldGrand, $newGrand)
                    if ($p.ProcessedAnchorPath) { $p.ProcessedAnchorPath = $p.ProcessedAnchorPath.Replace($oldGrand, $newGrand) }
                    if ($p.CustomImagePath) { $p.CustomImagePath = $p.CustomImagePath.Replace($oldGrand, $newGrand) }
                    foreach ($fr in $p.FileRows) { $fr.OldPath = $fr.OldPath.Replace($oldGrand, $newGrand) }
                }
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
    }

    # If rename-only (no heavy tasks), complete immediately without launching worker
    if (-not ($doMerge -or $doSlice -or $doExtract -or $doImage)) {
        $pJob.ProcessingOverlay.Visibility = "Collapsed"; $pJob.PickProcessingOverlay.Visibility = "Collapsed"
        $pJob.RowPanel.IsEnabled = $true; $pJob.IsDone = $true
        $pJob.BtnApply.Content = "Done"; $pJob.BtnApply.Background = Get-WpfColor "#4CAF72"
        $pJob.BtnApply.IsEnabled = $true; $pJob.BtnApply.Width = 100
        $pJob.PnlFiles.Children.Clear(); $pJob.FileRows.Clear()
        $files = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue | Sort-Object { switch -Regex ($_.Name) { 'Final\.3mf$' {0} 'Nest\.3mf$' {1} 'Full\.3mf$' {2} 'Full\.gcode\.3mf$' {3} default {4} } }, Name
        foreach ($fi in $files) { Add-FileRow $pJob $gpJob $fi }
        Update-ParentPreview $pJob $gpJob
        if (Test-Path $pJob.TempWork) { Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue }
        $script:activeProcessJob = $null
        if ($script:processQueue.Count -gt 0) { Start-NextProcess }
        return
    }

    # Anchor recovery scan (only needed when launching the worker)
    if (-not (Test-Path -LiteralPath $pJob.ProcessedAnchorPath)) {
        $recoveredAnchor = Find-AnchorFile $pJob.FolderPath
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
    if ($doLogs) { [void]$sb.AppendLine("Start-Transcript -Path `"$dir\Worker_PS_Log.txt`" -Force") }
    [void]$sb.AppendLine("Add-Type -AssemblyName System.IO.Compression.FileSystem")

    if ($doMerge) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'MERGING...' -Force")
        [void]$sb.AppendLine("& `"$scriptDir\merge_3mf_worker.ps1`" -WorkDir `"$($pJob.TempWork)`" -InputPath `"$anchorPath`" -OutputPath `"$tempOut`" -DoColors `"0`"")
        [void]$sb.AppendLine("if (Test-Path `"$tempOut`") {")
        [void]$sb.AppendLine("    if (Test-Path `"$nestPath`") { Remove-Item `"$nestPath`" -Force }")
        [void]$sb.AppendLine("    Rename-Item -Path `"$anchorPath`" -NewName `"$($basePrefix)Nest.3mf`" -Force")
        [void]$sb.AppendLine("    Rename-Item -Path `"$tempOut`"    -NewName `"$($baseName).3mf`"      -Force")
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
        [void]$sb.AppendLine("& `"$scriptDir\Slice_worker.ps1`" -InputPath `"$anchorPath`" -IsolatedPath `"$finalPath`"")
    } elseif ($doExtract -or $doImage) {
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
        [void]$sb.AppendLine("    & `"$scriptDir\Slice_worker.ps1`" -InputPath `"$finalPath`"")
        [void]$sb.AppendLine("}")
    }

    if ($doExtract -or $doImage) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'EXTRACTING DATA...' -Force")
        $imageFlag = if ($doImage) { "-GenerateImage" } else { "" }
        [void]$sb.AppendLine("if (Test-Path `"$slicedFile`") {")
        [void]$sb.AppendLine("    Remove-Item `"$tsvFile`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("    & `"$scriptDir\DataExtract_worker.ps1`" -InputFile `"$slicedFile`" -SingleFile `"$singleFile`" -IndividualTsvPath `"$tsvFile`" $imageFlag")
        [void]$sb.AppendLine("    Remove-Item `"$singleFile`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("} else {")
        [void]$sb.AppendLine("    Set-Content -Path `"$statusFile`" -Value '[ERROR] SLICE FAILED - MISSING GCODE' -Force")
        [void]$sb.AppendLine("    Start-Sleep -Seconds 4")
        [void]$sb.AppendLine("}")
    }

    if ($doImage) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'IMAGE INJECTION...' -Force")
        [void]$sb.AppendLine("`$batPath = Join-Path `"$scriptDir`" '..\callers\ReplaceImageNew.bat'")
        [void]$sb.AppendLine("if (Test-Path `$batPath) {")
        [void]$sb.AppendLine("    `$argList = '/c `"`"' + `$batPath + '`" `"' + `"$dir`" + '`"`"'")
        if ($doLogs) {
            [void]$sb.AppendLine("    Start-Process -FilePath 'cmd.exe' -ArgumentList `$argList -Wait -WindowStyle Hidden -RedirectStandardOutput `"$dir\ImageGen_Log.txt`" -RedirectStandardError `"$dir\ImageGen_Errors.txt`"")
        } else {
            [void]$sb.AppendLine("    Start-Process -FilePath 'cmd.exe' -ArgumentList `$argList -Wait -WindowStyle Hidden")
        }
        [void]$sb.AppendLine("}")
    }

    if ($doLogs) {
        [void]$sb.AppendLine("Stop-Transcript")
    } else {
        [void]$sb.AppendLine("Get-ChildItem -Path `"$dir`" -Filter `"*ProcessLog*.txt`" -ErrorAction SilentlyContinue | Remove-Item -Force")
        [void]$sb.AppendLine("Get-ChildItem -Path `"$dir`" -Filter `"*_Log.txt`" -ErrorAction SilentlyContinue | Remove-Item -Force")
        [void]$sb.AppendLine("Get-ChildItem -Path `"$dir`" -Filter `"*_Errors.txt`" -ErrorAction SilentlyContinue | Remove-Item -Force")
    }
    [void]$sb.AppendLine("Remove-Item `"$statusFile`" -Force -ErrorAction SilentlyContinue")
    [void]$sb.AppendLine("Remove-Item `"$($pJob.TempWork)`" -Recurse -Force -ErrorAction SilentlyContinue")

    Set-Content -Path $workerScript -Value $sb.ToString()
    $script:activeProcess = Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$workerScript`"" -PassThru -WindowStyle Hidden
}

# --- 7. DYNAMIC UI GENERATION ---
function Build-PJob($parentPath, $anchorFile, $gpJob) {
    $tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -ItemType Directory -Path $tempWork | Out-Null

    $isZip3mf = $anchorFile.Extension -imatch '\.3mf$' -and $anchorFile.Name -notmatch '(?i)\.gcode\.3mf$'
    if ($isZip3mf) {
        try { [System.IO.Compression.ZipFile]::ExtractToDirectory($anchorFile.FullName, $tempWork) } catch {}
    }

    $activeSlots = New-Object System.Collections.ArrayList

    if ($isZip3mf) {
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

        # plate_N.json contains exactly which filament slots are used on each plate
        # (filament_ids is 0-based). Use this when available — it's more precise than
        # scanning the whole model, and correctly excludes slots only used on other plates.
        $plateSlots = New-Object System.Collections.Generic.HashSet[string]
        $metaDir = Join-Path $tempWork "Metadata"
        if (Test-Path $metaDir) {
            foreach ($pjf in (Get-ChildItem -Path $metaDir -Filter 'plate_*.json' -ErrorAction SilentlyContinue)) {
                try {
                    $pjObj = ([System.IO.File]::ReadAllText($pjf.FullName, [System.Text.Encoding]::UTF8)) | ConvertFrom-Json
                    if ($pjObj.filament_ids) {
                        foreach ($fid in $pjObj.filament_ids) { $plateSlots.Add(($fid + 1).ToString()) | Out-Null }
                    }
                } catch {}
            }
        }

        if ($plateSlots.Count -gt 0) {
            # Plate data found — use it directly; it already represents only what's on the plates
            $UsedSlots = $plateSlots
        } else {
            # No plate JSON — fall back to scanning model_settings.config and 3dmodel.model
            if (Test-Path $modSetPath) {
                try {
                    [xml]$modXml = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
                    foreach ($node in $modXml.SelectNodes('//metadata[contains(@key, "extruder")]')) {
                        $val = $node.GetAttribute('value')
                        if (-not [string]::IsNullOrWhiteSpace($val)) { $UsedSlots.Add($val) | Out-Null }
                    }
                } catch {}
            }

            $modelFile = Get-ChildItem -Path $tempWork -Filter '3dmodel.model' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($modelFile -and (Test-Path $modelFile.FullName)) {
                try {
                    $modelContent = [System.IO.File]::ReadAllText($modelFile.FullName, [System.Text.Encoding]::UTF8)
                    $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
                    foreach ($m in $matMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
                } catch {}
            }
        }

        foreach ($hex in $SlotMap.Keys) {
            $slotId = $SlotMap[$hex]
            if ($UsedSlots.Contains($slotId)) {
                $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
                $matchedName = if ($HexToName.Contains($checkHex)) { $HexToName[$checkHex] } else { "" }
                $activeSlots.Add([PSCustomObject]@{ OldHex = $checkHex; Name = $matchedName }) | Out-Null
            }
        }
        if ($activeSlots.Count -gt 8) { $activeSlots = $activeSlots[0..7] }
    }

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
    # Left column absorbs all the resizing
    $pGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    # Right column stays strictly rigid at 560px
    $pGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(560))}))
    $pBorder.Child = $pGrid

# ── LEFT COLUMN: Card panel + Pick panel ────────────────────────
    $leftGrid = New-Object System.Windows.Controls.Grid
    $leftGrid.VerticalAlignment = "Top"
    $leftGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))
    $leftGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))

    # The Viewbox perfectly scales everything inside it to fit the shrinking column
    $viewbox = New-Object System.Windows.Controls.Viewbox
    $viewbox.Stretch = "Uniform"
    $viewbox.StretchDirection = [System.Windows.Controls.StretchDirection]::DownOnly
    $viewbox.HorizontalAlignment = "Left"
    $viewbox.VerticalAlignment = "Top"
    $viewbox.Child = $leftGrid

    [System.Windows.Controls.Grid]::SetColumn($viewbox, 0); $pGrid.Children.Add($viewbox) | Out-Null

    # Card panel
    $cardGrid = New-Object System.Windows.Controls.Grid
    $cardGrid.Height = 438; $cardGrid.Width = 438
    $cardGrid.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $cardGrid.Background = [System.Windows.Media.Brushes]::Transparent

    $pbModel = New-Object System.Windows.Controls.Image
    $pbModel.Stretch = "Uniform"
    $pbModel.HorizontalAlignment = "Left"
    $pbModel.Margin = New-Object System.Windows.Thickness(10, 20, 150, 10)
    $pbModel.Cursor = [System.Windows.Input.Cursors]::Hand
    $cardGrid.Children.Add($pbModel) | Out-Null
    $pJob.PbPlate = $pbModel
    $pbModel.Add_MouseLeftButtonDown({
        if ($_.ClickCount -ge 2 -and $null -ne $this.Source) {
            $viewer = New-Object System.Windows.Window
            $viewer.Title = "Card Image"; $viewer.Background = Get-WpfColor "#0D0E10"
            $viewer.SizeToContent = "WidthAndHeight"; $viewer.WindowStartupLocation = "CenterScreen"
            $viewer.ResizeMode = "CanResizeWithGrip"
            $imgView = New-Object System.Windows.Controls.Image
            $imgView.Source = $this.Source; $imgView.MaxWidth = 900; $imgView.MaxHeight = 900; $imgView.Stretch = "Uniform"
            $imgView.Margin = New-Object System.Windows.Thickness(10)
            $viewer.Content = $imgView
            $viewer.ShowDialog() | Out-Null
        }
    })

    # Determine Main Image (Custom PNG vs Extracted Anchor 3MF)
    $diParent = [System.IO.DirectoryInfo]::new($parentPath)
    $gpName = if ($gpJob.DiGrand) { $gpJob.DiGrand.Name } else { "" }
    $customPng = Get-ChildItem -Path $parentPath -Filter "*.png" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" } | Select-Object -First 1

    if ($customPng) {
        $pJob.CustomImagePath = $customPng.FullName
        $pbModel.Source = Load-WpfImage $customPng.FullName
    } elseif ($isZip3mf) {
        # Fallback to the Anchor 3MF extracted images (Checks both plate_1 and thumbnail!)
        $pJob.CustomImagePath = $null
        $baseImgPath = Join-Path $tempWork "Metadata\plate_1.png"
        if (-not (Test-Path $baseImgPath)) { $baseImgPath = Join-Path $tempWork "Metadata\thumbnail.png" }
        $pbModel.Source = Load-WpfImage $baseImgPath
    } elseif ($anchorFile.Extension -imatch '\.png$') {
        $pJob.CustomImagePath = $anchorFile.FullName
        $pbModel.Source = Load-WpfImage $anchorFile.FullName
    } else {
        $pJob.CustomImagePath = $null  # STL or unknown — no preview available
    }

    # (PbPlateFinished defined below as a full-width overlay on leftGrid)

    # [CURRENT] gcode thumbnail
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
    $pbCurrent.Cursor = [System.Windows.Input.Cursors]::Hand
    $pbCurrent.Add_MouseLeftButtonDown({
        if ($_.ClickCount -ge 2 -and $null -ne $this.Source) {
            $viewer = New-Object System.Windows.Window
            $viewer.Title = "Current Plate"; $viewer.Background = Get-WpfColor "#0D0E10"
            $viewer.SizeToContent = "WidthAndHeight"; $viewer.WindowStartupLocation = "CenterScreen"
            $viewer.ResizeMode = "CanResizeWithGrip"
            $imgView = New-Object System.Windows.Controls.Image
            $imgView.Source = $this.Source; $imgView.MaxWidth = 900; $imgView.MaxHeight = 900; $imgView.Stretch = "Uniform"
            $imgView.Margin = New-Object System.Windows.Thickness(10)
            $viewer.Content = $imgView
            $viewer.ShowDialog() | Out-Null
        }
    })

    $gcodeFile = Get-ChildItem -Path $parentPath -Filter "*Full.gcode.3mf" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($gcodeFile) {
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            # Find either plate_1.png or thumbnail.png inside the gcode file
            $entry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/(plate_1|thumbnail)\.png$" } | Select-Object -First 1
            if ($entry) {
                $gcodeImgPath = Join-Path $tempWork "plate_1_gcode.png"
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $gcodeImgPath, $true)
                $pbCurrent.Source = Load-WpfImage $gcodeImgPath
            } else { $currentGrid.Visibility = "Collapsed" }
        } catch { $currentGrid.Visibility = "Collapsed" }
        finally { if ($null -ne $zip) { $zip.Dispose() } }
    } else { $currentGrid.Visibility = "Collapsed" }

    # Overlay Labels
    $lblCharCard = Create-TextBlock "" "#E8A135" 20 "Bold"
    $lblCharCard.HorizontalAlignment = "Right"; $lblCharCard.VerticalAlignment = "Top"
    $lblCharCard.Margin = New-Object System.Windows.Thickness(0, 10, 10, 0)
    $cardGrid.Children.Add($lblCharCard) | Out-Null
    $pJob.LblCharCard = $lblCharCard

    $lblSkipTime = Create-TextBlock "Skip Time: 00" "#FFFFFF" 12 "Bold"
    $lblSkipTime.HorizontalAlignment = "Left"; $lblSkipTime.VerticalAlignment = "Bottom"
    $lblSkipTime.Margin = New-Object System.Windows.Thickness(10, 0, 0, 10)
    $cardGrid.Children.Add($lblSkipTime) | Out-Null
    $btnBrowseImg = New-Object System.Windows.Controls.Button
    $btnBrowseImg.Content = "Browse Images"
    $btnBrowseImg.Foreground = Get-WpfColor "#FFFFFF"
    $btnBrowseImg.FontWeight = [System.Windows.FontWeights]::Bold
    $btnBrowseImg.Width = 110; $btnBrowseImg.Height = 22
    $btnBrowseImg.BorderThickness = 0; $btnBrowseImg.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnBrowseImg.HorizontalAlignment = "Left"; $btnBrowseImg.VerticalAlignment = "Bottom"
    $btnBrowseImg.Margin = New-Object System.Windows.Thickness(10, 0, 0, 30)

    if ($pJob.CustomImagePath) { $btnBrowseImg.Background = Get-WpfColor "#4CAF72" }
    else { $btnBrowseImg.Background = Get-WpfColor "#E8A135" }

    $btnBrowseImg.Tag = $pJob
    $btnBrowseImg.Add_Click({
        $job = $this.Tag
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Title = "Select Custom Card Image"
        $dialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg"
        if ($dialog.ShowDialog() -eq $true) {
            $srcPath = $dialog.FileName
            $destPath = Join-Path $job.FolderPath ([System.IO.Path]::GetFileName($srcPath))
            if ($srcPath -ne $destPath) { Copy-Item -Path $srcPath -Destination $destPath -Force }
            $job.CustomImagePath = $destPath
            $job.PbPlate.Source = Load-WpfImage $destPath
            $this.Background = Get-WpfColor "#4CAF72"
        }
    })
    $cardGrid.Children.Add($btnBrowseImg) | Out-Null
    $pJob.BtnBrowseImg = $btnBrowseImg

    # Color Slots (Protected against malformed hex codes)
    $colorsOverlayStack = New-Object System.Windows.Controls.StackPanel
    $colorsOverlayStack.Orientation = "Vertical"
    $colorsOverlayStack.HorizontalAlignment = "Right"
    $colorsOverlayStack.VerticalAlignment = "Center"
    $colorsOverlayStack.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)

    # Scale swatch size and spacing to fit 5–8 color slots inside the 438px card
    $slotCount    = $activeSlots.Count
    $swatchSize   = if ($slotCount -le 4) { 52 } elseif ($slotCount -le 5) { 44 } elseif ($slotCount -le 6) { 38 } else { 32 }
    $rowMarginBtm = if ($slotCount -le 4) { 15 } elseif ($slotCount -le 5) { 10 } elseif ($slotCount -le 6) { 8  } else { 6  }
    $numFontSize  = if ($slotCount -le 4) { 14 } elseif ($slotCount -le 6) { 12 } else { 10 }
    $comboMinW    = if ($slotCount -ge 5) { 90 } else { 110 }

    $slotIdx = 1
    foreach ($slotData in $activeSlots) {
        $rowStack = New-Object System.Windows.Controls.StackPanel
        $rowStack.Orientation = "Horizontal"; $rowStack.HorizontalAlignment = "Right"
        $rowStack.Margin = New-Object System.Windows.Thickness(0, 0, 0, $rowMarginBtm)

        $textStack = New-Object System.Windows.Controls.StackPanel
        $textStack.Orientation = "Vertical"; $textStack.VerticalAlignment = "Center"
        $textStack.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)

        $lblStatus = Create-TextBlock "" "#A0A0A0" 10 "Bold"
        $lblStatus.HorizontalAlignment = "Right"
        if ($slotData.Name) { $lblStatus.Text = "[MATCHED]"; $lblStatus.Foreground = Get-WpfColor "#4CAF72" }
        else { $lblStatus.Text = "[UNMATCHED]"; $lblStatus.Foreground = Get-WpfColor "#D95F5F" }

        $combo = New-Object System.Windows.Controls.ComboBox; $combo.IsEditable = $true
        $combo.MinWidth = $comboMinW; $combo.MaxWidth = 210 # Safe limits so it doesn't hit the image
        $combo.Width = [System.Double]::NaN # Tells WPF to Auto-size dynamically!
        foreach ($k in $LibraryColors.Keys) { [void]$combo.Items.Add($k) }
        if ($slotData.Name) { $combo.Text = $slotData.Name } else { $combo.Text = "Select Color..." }

        $textStack.Children.Add($lblStatus) | Out-Null; $textStack.Children.Add($combo) | Out-Null

        $swatchColor = if ([string]::IsNullOrWhiteSpace($slotData.OldHex) -or $slotData.OldHex.Length -lt 7) { "#333333" } else { $slotData.OldHex }
        $swatchBorder = New-Object System.Windows.Controls.Border
        $swatchBorder.Width = $swatchSize; $swatchBorder.Height = $swatchSize
        $swatchBorder.Background = Get-WpfColor $swatchColor
        $swatchBorder.BorderBrush = Get-WpfColor "#2A2C35"; $swatchBorder.BorderThickness = New-Object System.Windows.Thickness(1)

        $r = 51; $g = 51; $b = 51
        if ($swatchColor.Length -ge 7) {
            try {
                $r = [Convert]::ToInt32($swatchColor.Substring(1,2), 16)
                $g = [Convert]::ToInt32($swatchColor.Substring(3,2), 16)
                $b = [Convert]::ToInt32($swatchColor.Substring(5,2), 16)
            } catch {}
        }
        $numColor = if ((0.299*$r + 0.587*$g + 0.114*$b) -gt 128) { "#000000" } else { "#FFFFFF" }

        $lblNum = Create-TextBlock $slotIdx.ToString() $numColor $numFontSize "Bold"
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
                try {
                    $r = [Convert]::ToInt32($newHex.Substring(1,2), 16)
                    $g = [Convert]::ToInt32($newHex.Substring(3,2), 16)
                    $b = [Convert]::ToInt32($newHex.Substring(5,2), 16)
                    $numColor = if ((0.299*$r + 0.587*$g + 0.114*$b) -gt 128) { "#000000" } else { "#FFFFFF" }
                    $data.LblNum.Foreground = Get-WpfColor $numColor
                } catch {}
            }
            if ($s.Text -eq $data.OrigName -and $data.OrigName) {
                $data.StatusLbl.Text = "[MATCHED]"; $data.StatusLbl.Foreground = Get-WpfColor "#4CAF72"
            } elseif ($LibraryColors.Contains($s.Text)) {
                $data.StatusLbl.Text = "[CHANGED]"; $data.StatusLbl.Foreground = Get-WpfColor "#E8A135"
            } else {
                $data.StatusLbl.Text = "[UNMATCHED]"; $data.StatusLbl.Foreground = Get-WpfColor "#D95F5F"
            }
            Validate-PJob $data.P
        })
        $slotIdx++
    }
    $cardGrid.Children.Add($colorsOverlayStack) | Out-Null

    # Processing Overlay (border + bottom label)
    $cardBorderOverlay = New-Object System.Windows.Controls.Border
    $cardBorderOverlay.BorderThickness = New-Object System.Windows.Thickness(6)
    $cardBorderOverlay.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220,232,161,53))
    $cardBorderOverlay.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(30,232,161,53))
    $cardBorderOverlay.Visibility = "Collapsed"
    $cardStatusLbl = New-Object System.Windows.Controls.TextBlock
    $cardStatusLbl.Text = "[ PROCESSING ]"
    $cardStatusLbl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255,232,161,53))
    $cardStatusLbl.FontSize = 13; $cardStatusLbl.FontWeight = [System.Windows.FontWeights]::Bold
    $cardStatusLbl.TextAlignment = "Center"; $cardStatusLbl.VerticalAlignment = "Bottom"; $cardStatusLbl.HorizontalAlignment = "Stretch"
    $cardStatusLbl.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(180,0,0,0))
    $cardStatusLbl.Padding = New-Object System.Windows.Thickness(5,4,5,6)
    $cardBorderOverlay.Child = $cardStatusLbl
    $cardGrid.Children.Add($cardBorderOverlay) | Out-Null
    $pJob.ProcessingOverlay = $cardBorderOverlay
    $pJob.CardStatusLabel   = $cardStatusLbl

    $cardGrid.AllowDrop = $true
    $cardGrid.Tag = $pJob
    $cardGrid.Add_DragOver({
        $files = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($files -and $files.Count -gt 0 -and [System.IO.Path]::GetExtension($files[0]) -imatch '\.png') {
            $_.Effects = [System.Windows.DragDropEffects]::Copy
        } else {
            $_.Effects = [System.Windows.DragDropEffects]::None
        }
        $_.Handled = $true
    })
    $cardGrid.Add_Drop({
        $files = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($files -and $files.Count -gt 0 -and [System.IO.Path]::GetExtension($files) -imatch '\.(png|jpg|jpeg)$') {
            $job = $this.Tag
            $srcPath = $files
            $destPath = Join-Path $job.FolderPath ([System.IO.Path]::GetFileName($srcPath))
            if ($srcPath -ne $destPath) { Copy-Item -Path $srcPath -Destination $destPath -Force }
            $job.CustomImagePath = $destPath
            $job.PbPlate.Source = Load-WpfImage $destPath
            if ($job.BtnBrowseImg) { $job.BtnBrowseImg.Background = Get-WpfColor "#4CAF72" }
        }
        $_.Handled = $true
    })
    [System.Windows.Controls.Grid]::SetColumn($cardGrid, 0); $leftGrid.Children.Add($cardGrid) | Out-Null

    # Pick Panel
    $pickGrid = New-Object System.Windows.Controls.Grid
    $pickGrid.Height = 438; $pickGrid.Width = 438

    $pbPick = New-Object System.Windows.Controls.Image; $pbPick.Stretch = "Uniform"
    $pbPick.Cursor = [System.Windows.Input.Cursors]::Hand
    $pickGrid.Children.Add($pbPick) | Out-Null
    $pJob.PbPick = $pbPick
    $pbPick.Add_MouseLeftButtonDown({
        if ($_.ClickCount -ge 2) {
            $t = $this.Tag
            if ($null -eq $t -or -not $t.Path) { return }
            $preMergePath = $null
            try {
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                $dir      = Split-Path $t.Original3mf -Parent
                $base     = (Split-Path $t.Original3mf -Leaf) -replace '(?i)_?Full\.gcode\.3mf$|_?Full\.3mf$', ''
                $nestFile = Join-Path $dir "${base}_Nest.3mf"
                if (-not (Test-Path $nestFile)) { [System.Windows.MessageBox]::Show("Nest.3mf not found:`n$nestFile", "Merge Map Error"); return }
                $preMergePath = Join-Path $env:TEMP "pre_verify_$([guid]::NewGuid().ToString().Substring(0,8)).png"
                $zip = [System.IO.Compression.ZipFile]::OpenRead($nestFile)
                $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "pick_1\.png$" } | Select-Object -First 1
                if ($null -eq $pickEntry) { $zip.Dispose(); [System.Windows.MessageBox]::Show("No pick_1.png found in:`n$nestFile", "Merge Map Error"); return }
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $preMergePath, $true)
                $zip.Dispose()
                $annotatedBmp = Get-ImageBasedMergeMap -preMergePath $preMergePath -postMergePath $t.Path
                if ($null -ne $annotatedBmp) {
                    $ms = New-Object System.IO.MemoryStream
                    $annotatedBmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
                    $annotatedBmp.Dispose()
                    $ms.Position = 0
                    $bmpSource = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bmpSource.BeginInit()
                    $bmpSource.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $bmpSource.StreamSource = $ms
                    $bmpSource.EndInit()
                    $ms.Dispose()
                    $viewer = New-Object System.Windows.Window
                    $viewer.Title = "Merged RGB Verification Overlay"; $viewer.Background = Get-WpfColor "#000000"
                    $viewer.SizeToContent = "WidthAndHeight"; $viewer.WindowStartupLocation = "CenterScreen"
                    $viewer.ResizeMode = "CanResizeWithGrip"
                    $imgView = New-Object System.Windows.Controls.Image
                    $imgView.Source = $bmpSource; $imgView.MaxWidth = 900; $imgView.MaxHeight = 900; $imgView.Stretch = "Uniform"
                    $imgView.Margin = New-Object System.Windows.Thickness(10)
                    $viewer.Content = $imgView
                    $viewer.ShowDialog() | Out-Null
                }
            } catch {
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Merge Map Error")
            } finally {
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                if ($preMergePath -and (Test-Path $preMergePath)) { Remove-Item $preMergePath -Force -ErrorAction SilentlyContinue }
            }
        }
    })

    # Pick processing border overlay
    $pickBorderOverlay = New-Object System.Windows.Controls.Border
    $pickBorderOverlay.BorderThickness = New-Object System.Windows.Thickness(6)
    $pickBorderOverlay.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220,232,161,53))
    $pickBorderOverlay.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(30,232,161,53))
    $pickBorderOverlay.Visibility = "Collapsed"
    $pickStatusLbl = New-Object System.Windows.Controls.TextBlock
    $pickStatusLbl.Text = "[ PROCESSING ]"
    $pickStatusLbl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255,232,161,53))
    $pickStatusLbl.FontSize = 13; $pickStatusLbl.FontWeight = [System.Windows.FontWeights]::Bold
    $pickStatusLbl.TextAlignment = "Center"; $pickStatusLbl.VerticalAlignment = "Bottom"; $pickStatusLbl.HorizontalAlignment = "Stretch"
    $pickStatusLbl.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(180,0,0,0))
    $pickStatusLbl.Padding = New-Object System.Windows.Thickness(5,4,5,6)
    $pickBorderOverlay.Child = $pickStatusLbl
    $pickGrid.Children.Add($pickBorderOverlay) | Out-Null
    $pJob.PickProcessingOverlay = $pickBorderOverlay
    $pJob.PickStatusLabel = $pickStatusLbl

    # Merge detected banner (top of pick image)
    $nestExists = Get-ChildItem -Path $parentPath -Filter "*Nest.3mf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    $mergeBanner = New-Object System.Windows.Controls.TextBlock
    $mergeBanner.Text = "MERGE DETECTED"
    $mergeBanner.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(210,30,140,60))
    $mergeBanner.Foreground = Get-WpfColor "#FFFFFF"
    $mergeBanner.FontSize = 12; $mergeBanner.FontWeight = [System.Windows.FontWeights]::Bold
    $mergeBanner.TextAlignment = "Center"; $mergeBanner.VerticalAlignment = "Top"; $mergeBanner.HorizontalAlignment = "Stretch"
    $mergeBanner.Padding = New-Object System.Windows.Thickness(0,5,0,5)
    $mergeBanner.Visibility = if ($nestExists) { "Visible" } else { "Collapsed" }
    $pickGrid.Children.Add($mergeBanner) | Out-Null
    $pJob.MergeBanner = $mergeBanner

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
                $pbPick.Tag = @{ Path = $pickPath; Original3mf = $anchorFile.FullName }
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    }
    [System.Windows.Controls.Grid]::SetColumn($pickGrid, 1); $leftGrid.Children.Add($pickGrid) | Out-Null

    # Finished image overlay — spans both card and pick columns, shown after processing completes
    $finishedOverlay = New-Object System.Windows.Controls.Border
    $finishedOverlay.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(230, 16, 17, 23))
    $finishedOverlay.Visibility = "Collapsed"
    [System.Windows.Controls.Grid]::SetColumn($finishedOverlay, 0)
    $pbPlateFinished = New-Object System.Windows.Controls.Image
    $pbPlateFinished.Stretch = "Uniform"
    $pbPlateFinished.Cursor = [System.Windows.Input.Cursors]::Hand
    $pbPlateFinished.Add_MouseLeftButtonDown({
        if ($_.ClickCount -ge 2 -and $null -ne $this.Source) {
            $viewer = New-Object System.Windows.Window
            $viewer.Title = "Finished Plate"; $viewer.Background = Get-WpfColor "#0D0E10"
            $viewer.SizeToContent = "WidthAndHeight"; $viewer.WindowStartupLocation = "CenterScreen"
            $viewer.ResizeMode = "CanResizeWithGrip"
            $imgView = New-Object System.Windows.Controls.Image
            $imgView.Source = $this.Source; $imgView.MaxWidth = 900; $imgView.MaxHeight = 900; $imgView.Stretch = "Uniform"
            $imgView.Margin = New-Object System.Windows.Thickness(10)
            $viewer.Content = $imgView
            $viewer.ShowDialog() | Out-Null
        }
    })
    $finishedOverlay.Child = $pbPlateFinished
    $leftGrid.Children.Add($finishedOverlay) | Out-Null
    $pJob.PbPlateFinished = $pbPlateFinished
    $pJob.FinishedOverlay  = $finishedOverlay

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

    $btnOpenFolder = New-Object System.Windows.Controls.Button
    $btnOpenFolder.Content = "Open Folder"; $btnOpenFolder.Background = Get-WpfColor "#2A2C35"; $btnOpenFolder.Foreground = Get-WpfColor "#FFFFFF"
    $btnOpenFolder.Width = 100; $btnOpenFolder.Height = 25; $btnOpenFolder.BorderThickness = 0
    $btnOpenFolder.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnOpenFolder.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnOpenFolder.Tag = $pJob
    $btnOpenFolder.Add_Click({ Start-Process "explorer.exe" $this.Tag.FolderPath })
    $btnHdrStack.Children.Add($btnOpenFolder) | Out-Null

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
        } else {
            Update-GpFileCount $t.G
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

    $tasksRow1 = New-Object System.Windows.Controls.WrapPanel; $tasksRow1.Orientation = "Horizontal"
    $chkRename  = New-Object System.Windows.Controls.CheckBox; $chkRename.Content  = "Rename Files/Folders"; $chkRename.IsChecked  = $false; $chkRename.Foreground  = Get-WpfColor "#FFFFFF"; $chkRename.Margin  = New-Object System.Windows.Thickness(0,0,15,0)
    $chkMerge   = New-Object System.Windows.Controls.CheckBox; $chkMerge.Content   = "Merge";                $chkMerge.IsChecked   = $false; $chkMerge.Foreground   = Get-WpfColor "#FFFFFF"; $chkMerge.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkSlice   = New-Object System.Windows.Controls.CheckBox; $chkSlice.Content   = "Slice / Export Gcode"; $chkSlice.IsChecked   = $false; $chkSlice.Foreground   = Get-WpfColor "#FFFFFF"; $chkSlice.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkExtract = New-Object System.Windows.Controls.CheckBox; $chkExtract.Content = "Extract Data";         $chkExtract.IsChecked = $false; $chkExtract.Foreground = Get-WpfColor "#FFFFFF"; $chkExtract.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkImage   = New-Object System.Windows.Controls.CheckBox; $chkImage.Content   = "Generate Image Card";  $chkImage.IsChecked   = $false; $chkImage.Foreground   = Get-WpfColor "#FFFFFF"; $chkImage.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkLogs    = New-Object System.Windows.Controls.CheckBox; $chkLogs.Content    = "Create Logs";          $chkLogs.IsChecked    = $false; $chkLogs.Foreground    = Get-WpfColor "#FFFFFF"
    $tasksRow1.Children.Add($chkRename) | Out-Null; $tasksRow1.Children.Add($chkMerge) | Out-Null; $tasksRow1.Children.Add($chkSlice) | Out-Null
    $tasksRow1.Children.Add($chkExtract) | Out-Null; $tasksRow1.Children.Add($chkImage) | Out-Null
    $tasksRow1.Children.Add($chkLogs) | Out-Null
    $tasksOuter.Children.Add($tasksRow1) | Out-Null

    $tasksRow2 = New-Object System.Windows.Controls.StackPanel
    $tasksRow2.Orientation = "Horizontal"; $tasksRow2.Margin = New-Object System.Windows.Thickness(0,8,0,0)
    $btnSelAll = New-Object System.Windows.Controls.Button; $btnSelAll.Content = "Select All"; $btnSelAll.Background = Get-WpfColor "#2A2C35"; $btnSelAll.Foreground = Get-WpfColor "#FFFFFF"; $btnSelAll.Width = 100; $btnSelAll.Height = 25; $btnSelAll.BorderThickness = 0; $btnSelAll.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnSelAll.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnDeselAll = New-Object System.Windows.Controls.Button; $btnDeselAll.Content = "Deselect All"; $btnDeselAll.Background = Get-WpfColor "#2A2C35"; $btnDeselAll.Foreground = Get-WpfColor "#FFFFFF"; $btnDeselAll.Width = 100; $btnDeselAll.Height = 25; $btnDeselAll.BorderThickness = 0; $btnDeselAll.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnDeselAll.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRevertMerge = New-Object System.Windows.Controls.Button; $btnRevertMerge.Content = "Revert Merge"; $btnRevertMerge.Width = 110; $btnRevertMerge.Height = 25; $btnRevertMerge.BorderThickness = 0; $btnRevertMerge.Cursor = [System.Windows.Input.Cursors]::Hand
    if ($nestExists) {
        $btnRevertMerge.Background = Get-WpfColor "#D95F5F"; $btnRevertMerge.Foreground = Get-WpfColor "#FFFFFF"; $btnRevertMerge.IsEnabled = $true
    } else {
        $btnRevertMerge.Background = Get-WpfColor "#3A3A3A"; $btnRevertMerge.Foreground = Get-WpfColor "#666666"; $btnRevertMerge.IsEnabled = $false
        $btnRevertMerge.ToolTip = "No merged file detected"
    }
    $pJob.BtnRevertMerge = $btnRevertMerge
    $tasksRow2.Children.Add($btnSelAll) | Out-Null; $tasksRow2.Children.Add($btnDeselAll) | Out-Null; $tasksRow2.Children.Add($btnRevertMerge) | Out-Null
    $tasksOuter.Children.Add($tasksRow2) | Out-Null

    $tasksBox.Child = $tasksOuter
    $rightStack.Children.Add($tasksBox) | Out-Null

    $pJob.ChkRename = $chkRename; $pJob.ChkMerge = $chkMerge; $pJob.ChkSlice = $chkSlice; $pJob.ChkExtract = $chkExtract; $pJob.ChkImage = $chkImage; $pJob.ChkLogs = $chkLogs
    if ($nestExists) {
        $chkMerge.IsChecked = $false; $chkMerge.IsEnabled = $false; $chkMerge.Foreground = Get-WpfColor "#555555"
        $chkMerge.ToolTip = "Remove Nest.3mf or Revert Merge before merging again"
    }

    # Checkbox interdependencies
    $tasksData = @{ Rename = $chkRename; Merge = $chkMerge; Slice = $chkSlice; Extract = $chkExtract; Image = $chkImage; Logs = $chkLogs; PJob = $pJob; GpJob = $gpJob }

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
        $t.Rename.IsChecked = $true; $t.Slice.IsChecked = $true; $t.Extract.IsChecked = $true; $t.Image.IsChecked = $true
        if ($t.Merge.IsEnabled) { $t.Merge.IsChecked = $true }
    })

    $btnDeselAll.Tag = $tasksData
    $btnDeselAll.Add_Click({
        $t = $this.Tag
        $t.Rename.IsChecked = $false; $t.Slice.IsChecked = $false; $t.Extract.IsChecked = $false; $t.Image.IsChecked = $false; $t.Logs.IsChecked = $false
        if ($t.Merge.IsEnabled) { $t.Merge.IsChecked = $false }
    })

    $btnRevertMerge.Tag = @{ P = $pJob; G = $gpJob }
    $btnRevertMerge.Add_Click({
        $t = $this.Tag; $pj = $t.P; $gp = $t.G
        $batPath = Join-Path $scriptDir "..\callers\RevertMerge.bat"
        if (-not (Test-Path $batPath)) { [System.Windows.MessageBox]::Show("RevertMerge.bat not found.", "Error") | Out-Null; return }
        $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }
        $pj.BtnApply.Content = "Reverting..."; $pj.BtnApply.Width = 150
        if ($pj.BtnRevertDone) { $pj.BtnRevertDone.Visibility = "Collapsed" }
        $pj.RowPanel.IsEnabled = $false
        $redBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220, 217, 95, 95))
        $pj.ProcessingOverlay.BorderBrush = $redBrush
        $pj.ProcessingOverlay.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(30, 217, 95, 95))
        $pj.CardStatusLabel.Text       = "[REVERTING...]"
        $pj.CardStatusLabel.Foreground = $redBrush
        $pj.ProcessingOverlay.Visibility = "Visible"
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
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

    $charStack = New-Object System.Windows.Controls.StackPanel; $charStack.Margin = New-Object System.Windows.Thickness(0,0,12,0)
    $charStack.Children.Add((Create-TextBlock "Character *" "#A0A0A0" 12 "Normal")) | Out-Null
    $tbChar = New-Object System.Windows.Controls.TextBox; $tbChar.Text = $fills.Char; $tbChar.Width = 175; $tbChar.Background = Get-WpfColor "#1E2028"; $tbChar.Foreground = Get-WpfColor "#FFFFFF"
    $charStack.Children.Add($tbChar) | Out-Null; $editStack.Children.Add($charStack) | Out-Null
    $pJob.TBChar = $tbChar

    $tagStack = New-Object System.Windows.Controls.StackPanel; $tagStack.Margin = New-Object System.Windows.Thickness(0,0,12,0)
    $tagStack.Children.Add((Create-TextBlock "Tag" "#A0A0A0" 12 "Normal")) | Out-Null
    $cbTag = New-Object System.Windows.Controls.ComboBox; $cbTag.IsEditable = $false; $cbTag.Width = 80
    $cbTag.Background = Get-WpfColor "#1E2028"; $cbTag.Foreground = Get-WpfColor "#E8A135"
    $cbTag.BorderBrush = Get-WpfColor "#5A78C4"; $cbTag.BorderThickness = New-Object System.Windows.Thickness(1)
    $cbTag.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, [System.Windows.Controls.ToolBar]::ComboBoxStyleKey)
    $cbTag.Resources[[System.Windows.SystemColors]::WindowBrushKey]        = Get-WpfColor "#1E2028"
    $cbTag.Resources[[System.Windows.SystemColors]::WindowTextBrushKey]    = Get-WpfColor "#E8A135"
    $cbTag.Resources[[System.Windows.SystemColors]::HighlightBrushKey]     = Get-WpfColor "#5A78C4"
    $cbTag.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey] = Get-WpfColor "#FFFFFF"
    $cbTagItemStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
    $cbTagItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, (Get-WpfColor "#1E2028"))))
    $cbTagItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, (Get-WpfColor "#E8A135"))))
    $cbTag.ItemContainerStyle = $cbTagItemStyle
    [void]$cbTag.Items.Add("(none)")
    foreach ($tag in $script:Tags) { [void]$cbTag.Items.Add($tag) }
    $cbTag.SelectedItem = if ($fills.Tag -ne '') { $fills.Tag } else { "(none)" }
    $tagStack.Children.Add($cbTag) | Out-Null; $editStack.Children.Add($tagStack) | Out-Null
    $pJob.TBTag = $cbTag

    $adjStack = New-Object System.Windows.Controls.StackPanel
    $adjStack.Children.Add((Create-TextBlock "Adjective (Optional)" "#A0A0A0" 12 "Normal")) | Out-Null
    $cbAdj = New-Object System.Windows.Controls.ComboBox; $cbAdj.IsEditable = $true; $cbAdj.Width = 220
    $cbAdj.Background = Get-WpfColor "#1E2028"; $cbAdj.Foreground = Get-WpfColor "#FFFFFF"
    $cbAdj.BorderBrush = Get-WpfColor "#5A78C4"; $cbAdj.BorderThickness = New-Object System.Windows.Thickness(1)
    $cbAdj.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, [System.Windows.Controls.ToolBar]::ComboBoxStyleKey)
    $cbAdj.Resources[[System.Windows.SystemColors]::WindowBrushKey]        = Get-WpfColor "#1E2028"
    $cbAdj.Resources[[System.Windows.SystemColors]::WindowTextBrushKey]    = Get-WpfColor "#FFFFFF"
    $cbAdj.Resources[[System.Windows.SystemColors]::HighlightBrushKey]     = Get-WpfColor "#5A78C4"
    $cbAdj.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey] = Get-WpfColor "#FFFFFF"
    $cbAdjItemStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
    $cbAdjItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, (Get-WpfColor "#1E2028"))))
    $cbAdjItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, (Get-WpfColor "#FFFFFF"))))
    $cbAdj.ItemContainerStyle = $cbAdjItemStyle
    [void]$cbAdj.Items.Add("")
    foreach ($adj in $script:AdjPresets) { [void]$cbAdj.Items.Add($adj) }
    $cbAdj.Text = $fills.Adj
    $adjStack.Children.Add($cbAdj) | Out-Null; $editStack.Children.Add($adjStack) | Out-Null
    $pJob.TBAdj = $cbAdj
    $editBox.Child = $editStack; $rightStack.Children.Add($editBox) | Out-Null

    # Files list
    $pnlFiles = New-Object System.Windows.Controls.StackPanel; $pnlFiles.Margin = New-Object System.Windows.Thickness(0,10,0,0)
    $rightStack.Children.Add($pnlFiles) | Out-Null; $pJob.PnlFiles = $pnlFiles
    $files = Get-ChildItem -Path $parentPath -File -ErrorAction SilentlyContinue | Sort-Object { switch -Regex ($_.Name) { 'Final\.3mf$' {0} 'Nest\.3mf$' {1} 'Full\.3mf$' {2} 'Full\.gcode\.3mf$' {3} default {4} } }, Name
    foreach ($fi in $files) { Add-FileRow $pJob $gpJob $fi }

    # Apply + Revert + Delete Logs buttons
    $applyRow = New-Object System.Windows.Controls.StackPanel; $applyRow.Orientation = "Horizontal"; $applyRow.HorizontalAlignment = "Right"; $applyRow.Margin = New-Object System.Windows.Thickness(0,15,0,0)

    $btnDeleteLogs = New-Object System.Windows.Controls.Button
    $btnDeleteLogs.Content = "Delete Logs"; $btnDeleteLogs.Background = Get-WpfColor "#555555"; $btnDeleteLogs.Foreground = Get-WpfColor "#FFFFFF"
    $btnDeleteLogs.FontWeight = [System.Windows.FontWeights]::Bold; $btnDeleteLogs.Width = 100; $btnDeleteLogs.Height = 35; $btnDeleteLogs.BorderThickness = 0
    $btnDeleteLogs.Margin = New-Object System.Windows.Thickness(0,0,15,0); $btnDeleteLogs.Cursor = [System.Windows.Input.Cursors]::Hand
    $applyRow.Children.Add($btnDeleteLogs) | Out-Null

    $btnDeleteLogs.Tag = @{ P = $pJob }
    $btnDeleteLogs.Add_Click({
        $t = $this.Tag
        $logs = Get-ChildItem -Path $t.P.FolderPath -Include "*_Log.txt", "*_Errors.txt", "*ProcessLog*.txt" -Recurse -File -ErrorAction SilentlyContinue
        if ($logs.Count -gt 0) {
            $res = [System.Windows.MessageBox]::Show("Found $($logs.Count) log files in this folder. Are you sure you want to delete them?", "Confirm Delete Logs", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($res -eq 'Yes') {
                $logs | Remove-Item -Force -ErrorAction SilentlyContinue
                [System.Windows.MessageBox]::Show("Logs deleted successfully.", "Logs Cleared", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        } else {
            [System.Windows.MessageBox]::Show("No log files found in this folder.", "No Logs", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
    })

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
        if ($this.Content -eq "Done") {
            Refresh-PJob $t.P $t.G
        } elseif ($this.Content -eq "KEEP") {
            # Move the finished plate image to the [CURRENT] thumbnail
            if ($t.P.PbPlateFinished.Source -ne $null) {
                $t.P.PbCurrent.Source = $t.P.PbPlateFinished.Source
                if ($t.P.CurrentThumb) { $t.P.CurrentThumb.Visibility = "Visible" }
            }
            # Reload card editor with folder PNG or 3MF plate_1.png
            $diParent = [System.IO.DirectoryInfo]::new($t.P.FolderPath)
            $gpName = if ($t.G.DiGrand) { $t.G.DiGrand.Name } else { "" }
            $customPng = Get-ChildItem -Path $t.P.FolderPath -Filter "*.png" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" } | Select-Object -First 1
            if ($customPng) {
                $t.P.CustomImagePath = $customPng.FullName
                $t.P.PbPlate.Source = Load-WpfImage $customPng.FullName
            } else {
                $t.P.CustomImagePath = $null
                $baseImgPath = Join-Path $t.P.TempWork "Metadata\plate_1.png"
                if (-not (Test-Path $baseImgPath)) { $baseImgPath = Join-Path $t.P.TempWork "Metadata\thumbnail.png" }
                $t.P.PbPlate.Source = Load-WpfImage $baseImgPath
            }
            $this.Content = "Finished"; $this.Background = Get-WpfColor "#333333"; $this.IsEnabled = $false; $this.Width = 150
            if ($t.P.BtnRevertDone) { $t.P.BtnRevertDone.Visibility = "Collapsed" }
            if ($t.P.FinishedOverlay) { $t.P.FinishedOverlay.Visibility = "Collapsed" }
        } else {
            Enqueue-PJob $t.P $t.G
        }
    })

    $btnRevertDone.Tag = @{ P = $pJob; G = $gpJob }
    $btnRevertDone.Add_Click({
        $t = $this.Tag; $pj = $t.P; $gp = $t.G
        $batPath = Join-Path $scriptDir "..\callers\RevertMerge.bat"
        if (-not (Test-Path $batPath)) { return }
        $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }
        $pj.BtnApply.Content = "Reverting..."; $pj.BtnApply.Width = 150; $pj.BtnRevertDone.Visibility = "Collapsed"
        $pj.RowPanel.IsEnabled = $false
        $redBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220, 217, 95, 95))
        $pj.ProcessingOverlay.BorderBrush = $redBrush
        $pj.ProcessingOverlay.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(30, 217, 95, 95))
        $pj.CardStatusLabel.Text       = "[REVERTING...]"
        $pj.CardStatusLabel.Foreground = $redBrush
        $pj.ProcessingOverlay.Visibility = "Visible"
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
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
    $cbTag.Tag  = @{ P = $pJob; G = $gpJob }; $cbTag.Add_SelectionChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })
    $cbAdj.Tag = @{ P = $pJob; G = $gpJob }
    $cbAdj.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, [System.Windows.Controls.TextChangedEventHandler]{ param($s,$e); $t = $s.Tag; Update-ParentPreview $t.P $t.G })

    Update-ParentPreview $pJob $gpJob
    $gpJob.ParentListStack.Children.Add($pBorder) | Out-Null
    return $pJob
}

function Update-GpFileCount($gpJob) {
    if ($null -ne $gpJob.LblFileCount) {
        $c = $gpJob.Parents.Count
        $gpJob.LblFileCount.Text = "($c plate$(if ($c -ne 1) { 's' }))"
    }
}

function Build-GpJob($gpPath, $parentDict) {
    $diGrand = if ($gpPath -notlike "ROOT_*") { [System.IO.DirectoryInfo]::new($gpPath) } else { $null }
    $gpName = if ($diGrand) { $diGrand.Name } else { "(No Parent Folder)" }

    # Detect printer prefix and strip all leading qualifier tokens (printer prefix + tags)
    # from the grandparent folder name to isolate the bare theme.
    # Handles any combination/order, e.g.:
    #   "P2S_KC_Licensing"  -> prefix "P2S", theme "Licensing"
    #   "KC_Licensing"      -> prefix "",    theme "Licensing"
    #   "P2S_Licensing"     -> prefix "P2S", theme "Licensing"
    #   "Licensing"         -> prefix "",    theme "Licensing"
    $gpDetectedPrefix = ""; $gpNameForTheme = $gpName
    if ($gpName -ne "(No Parent Folder)") {
        $gpTokens = [System.Collections.Generic.List[string]]($gpName -split '_' | Where-Object { $_ -ne '' })
        # Peel tokens from the front as long as they are a known printer prefix or tag
        while ($gpTokens.Count -gt 1) {
            $head = $gpTokens[0]
            if ($script:PrinterPrefixes -icontains $head) {
                if ($gpDetectedPrefix -eq '') { $gpDetectedPrefix = $head }   # keep first printer prefix
                $gpTokens.RemoveAt(0)
            } elseif ($script:Tags -icontains $head) {
                $gpTokens.RemoveAt(0)   # strip tag qualifiers (KC, Big, Huge, …)
            } else {
                break
            }
        }
        $gpNameForTheme = $gpTokens -join '_'

        # Fallback: detect printer prefix from first anchor file stem if still not found
        if ($gpDetectedPrefix -eq "") {
            foreach ($pKey in $parentDict.Keys) {
                $stemTest = ($parentDict[$pKey]).BaseName -replace '(?i)_Full$', ''
                $afParts = $stemTest -split '_'
                if ($afParts.Count -gt 0 -and $script:PrinterPrefixes -icontains $afParts[0]) {
                    $gpDetectedPrefix = $afParts[0]; break
                }
            }
        }
    }

    $gpJob = @{ GpPath = $gpPath; DiGrand = $diGrand; Parents = New-Object System.Collections.ArrayList; CbPrefix = $null; GpRenameConfirmed = $false }
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

    # Current folder name (far left)
    $lblCurrentName = Create-TextBlock $gpName "#CCCCCC" 13 "Bold"
    $lblCurrentName.VerticalAlignment = "Center"; $lblCurrentName.Margin = New-Object System.Windows.Thickness(15,0,20,0)
    $headerStack.Children.Add($lblCurrentName) | Out-Null

    # Printer prefix dropdown
    $lblPrefix = Create-TextBlock "Printer: " "#E8A135" 14 "Bold"
    $lblPrefix.Margin = New-Object System.Windows.Thickness(0,0,0,0); $headerStack.Children.Add($lblPrefix) | Out-Null
    $cbPrefix = New-Object System.Windows.Controls.ComboBox; $cbPrefix.Width = 85
    $cbPrefix.Background = Get-WpfColor "#2A2C35"; $cbPrefix.Foreground = Get-WpfColor "#FFFFFF"
    $cbPrefix.BorderBrush = Get-WpfColor "#5A78C4"; $cbPrefix.BorderThickness = New-Object System.Windows.Thickness(1)
    $cbPrefix.VerticalAlignment = "Center"; $cbPrefix.Margin = New-Object System.Windows.Thickness(5,0,20,0)

    # Strip the Windows Light-Grey gradient and force it to use our Background color!
    $cbPrefix.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, [System.Windows.Controls.ToolBar]::ComboBoxStyleKey)

    $cbPrefix.Resources[[System.Windows.SystemColors]::WindowBrushKey]          = Get-WpfColor "#2A2C35"
    $cbPrefix.Resources[[System.Windows.SystemColors]::WindowTextBrushKey]      = Get-WpfColor "#FFFFFF"
    $cbPrefix.Resources[[System.Windows.SystemColors]::HighlightBrushKey]       = Get-WpfColor "#5A78C4"
    $cbPrefix.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey]   = Get-WpfColor "#FFFFFF"
    $cbItemStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
    $cbItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, (Get-WpfColor "#2A2C35"))))
    $cbItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, (Get-WpfColor "#FFFFFF"))))
    $cbPrefix.ItemContainerStyle = $cbItemStyle
    [void]$cbPrefix.Items.Add("(none)")
    foreach ($pfx in $script:PrinterPrefixes) { [void]$cbPrefix.Items.Add($pfx) }
    if ($gpDetectedPrefix -ne "" -and $script:PrinterPrefixes -icontains $gpDetectedPrefix) {
        $cbPrefix.SelectedItem = $gpDetectedPrefix
    } else { $cbPrefix.SelectedIndex = 0 }
    $headerStack.Children.Add($cbPrefix) | Out-Null; $gpJob.CbPrefix = $cbPrefix

    # Grandparent theme label + dropdown
    $lblGP = Create-TextBlock "Theme: " "#E8A135" 14 "Bold"
    $lblGP.Margin = New-Object System.Windows.Thickness(0,0,0,0); $headerStack.Children.Add($lblGP) | Out-Null

    $cbTheme = New-Object System.Windows.Controls.ComboBox; $cbTheme.Width = 175
    $cbTheme.Background = Get-WpfColor "#1E2028"; $cbTheme.Foreground = Get-WpfColor "#FFFFFF"
    $cbTheme.BorderBrush = Get-WpfColor "#5A78C4"; $cbTheme.BorderThickness = New-Object System.Windows.Thickness(1)
    $cbTheme.VerticalAlignment = "Center"; $cbTheme.Margin = New-Object System.Windows.Thickness(10,0,0,0)
    $cbTheme.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, [System.Windows.Controls.ToolBar]::ComboBoxStyleKey)
    $cbTheme.Resources[[System.Windows.SystemColors]::WindowBrushKey]        = Get-WpfColor "#1E2028"
    $cbTheme.Resources[[System.Windows.SystemColors]::WindowTextBrushKey]    = Get-WpfColor "#FFFFFF"
    $cbTheme.Resources[[System.Windows.SystemColors]::HighlightBrushKey]     = Get-WpfColor "#5A78C4"
    $cbTheme.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey] = Get-WpfColor "#FFFFFF"
    $cbThemeItemStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
    $cbThemeItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, (Get-WpfColor "#1E2028"))))
    $cbThemeItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, (Get-WpfColor "#FFFFFF"))))
    $cbTheme.ItemContainerStyle = $cbThemeItemStyle
    foreach ($theme in $script:GpThemes) { [void]$cbTheme.Items.Add($theme) }
    $matchedTheme = $script:GpThemes | Where-Object { ($_ -replace '[^a-zA-Z0-9]','') -ieq ($gpNameForTheme -replace '[^a-zA-Z0-9]','') } | Select-Object -First 1
    if ($matchedTheme) { $cbTheme.SelectedItem = $matchedTheme } else { $cbTheme.SelectedIndex = -1 }
    $headerStack.Children.Add($cbTheme) | Out-Null; $gpJob.TBTheme = $cbTheme

    $chkSkip = New-Object System.Windows.Controls.CheckBox; $chkSkip.Content = "Don't rename folder"
    $chkSkip.Foreground = Get-WpfColor "#FFFFFF"; $chkSkip.VerticalAlignment = "Center"
    $chkSkip.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    $headerStack.Children.Add($chkSkip) | Out-Null; $gpJob.ChkSkip = $chkSkip

    # Live preview of the full grandparent folder name (Prefix_Theme or just Theme)
    $lblGpPreview = Create-TextBlock "" "#6B9FD4" 14 "Bold"
    $lblGpPreview.VerticalAlignment = "Center"; $lblGpPreview.Margin = New-Object System.Windows.Thickness(20,0,0,0)
    $initGpPreview = if ($gpDetectedPrefix -ne "") { "$gpDetectedPrefix`_$gpNameForTheme" } else { $gpNameForTheme }
    $lblGpPreview.Text = if ($initGpPreview) { [char]0x2192 + " $initGpPreview" } else { "" }
    $headerStack.Children.Add($lblGpPreview) | Out-Null; $gpJob.LblGpPreview = $lblGpPreview

    $lblFileCount = Create-TextBlock "" "#888888" 11 "Normal"
    $lblFileCount.VerticalAlignment = "Center"; $lblFileCount.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    $headerStack.Children.Add($lblFileCount) | Out-Null; $gpJob.LblFileCount = $lblFileCount

    $headerGrid.Children.Add($headerStack) | Out-Null

    $gpRightBtnStack = New-Object System.Windows.Controls.StackPanel
    $gpRightBtnStack.Orientation = "Horizontal"; $gpRightBtnStack.HorizontalAlignment = "Right"
    $gpRightBtnStack.VerticalAlignment = "Center"; $gpRightBtnStack.Margin = New-Object System.Windows.Thickness(0,0,15,0)

    $btnCombineGp = New-Object System.Windows.Controls.Button
    $btnCombineGp.Content = "Combine TSV Data"; $btnCombineGp.Background = Get-WpfColor "#7B4FBF"; $btnCombineGp.Foreground = Get-WpfColor "#FFFFFF"
    $btnCombineGp.FontWeight = [System.Windows.FontWeights]::Bold; $btnCombineGp.Width = 140; $btnCombineGp.Height = 30; $btnCombineGp.BorderThickness = 0
    $btnCombineGp.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnCombineGp.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnCombineGp.Tag = $gpJob
    $btnCombineGp.Add_Click({
        $gp = $this.Tag
        $targetDir = $gp.GpPath
        if ([string]::IsNullOrWhiteSpace($targetDir) -or -not (Test-Path $targetDir)) {
            [System.Windows.MessageBox]::Show("Grandparent folder path is not valid.", "Combine TSV Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $folderName = Split-Path $targetDir -Leaf
        $outTsvPath = Join-Path $targetDir "${folderName}_Data.tsv"
        $tsvFiles = Get-ChildItem -Path $targetDir -Filter "*_Data.tsv" -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notmatch "(?i)^.*_Design_Data\.tsv$" }
        if ($tsvFiles.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No TSV data files found in:`n$targetDir", "Nothing to Combine", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $combined = [ordered]@{}
        foreach ($tsv in $tsvFiles) {
            if ($tsv.FullName -eq $outTsvPath) { continue }
            $line = Get-Content $tsv.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $key = ($line -split "`t")[0]
            $combined[$key] = $line
        }
        if ($combined.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No TSV data rows found to combine.", "Nothing to Combine", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $combined.Values | Set-Content -Path $outTsvPath -Encoding UTF8
        [System.Windows.Clipboard]::SetText($combined.Values -join "`r`n")
        [System.Windows.MessageBox]::Show(
            "Combined $($combined.Count) rows into:`n${folderName}_Data.tsv`n`nAll $($combined.Count) rows have been copied to your clipboard.",
            "Combine Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    })
    $gpRightBtnStack.Children.Add($btnCombineGp) | Out-Null

    $btnRemoveGp = New-Object System.Windows.Controls.Button
    $btnRemoveGp.Content = "Remove Group"; $btnRemoveGp.Background = Get-WpfColor "#D95F5F"; $btnRemoveGp.Foreground = Get-WpfColor "#FFFFFF"
    $btnRemoveGp.Width = 140; $btnRemoveGp.Height = 30; $btnRemoveGp.BorderThickness = 0
    $btnRemoveGp.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRemoveGp.Tag = @{ Container = $container; GpJob = $gpJob }
    $btnRemoveGp.Add_Click({
        $t = $this.Tag
        $mainStack.Children.Remove($t.Container) | Out-Null
        $script:jobs.Remove($t.GpJob) | Out-Null
    })
    $gpRightBtnStack.Children.Add($btnRemoveGp) | Out-Null
    $headerGrid.Children.Add($gpRightBtnStack) | Out-Null
    $gpStack.Children.Add($headerGrid) | Out-Null

    # --- THEME TASK BAR ---
    $themeBar = New-Object System.Windows.Controls.Border
    $themeBar.Background = Get-WpfColor "#21222B"
    $themeBar.BorderBrush = Get-WpfColor "#2A2C35"
    $themeBar.BorderThickness = New-Object System.Windows.Thickness(0,0,0,1)
    $themeBar.Padding = New-Object System.Windows.Thickness(15,10,15,10)

    $themeBarStack = New-Object System.Windows.Controls.StackPanel
    $themeBarStack.Orientation = "Horizontal"

    $chkThRename  = New-Object System.Windows.Controls.CheckBox; $chkThRename.Content  = "Rename";  $chkThRename.IsChecked  = $false; $chkThRename.Foreground  = Get-WpfColor "#CCCCCC"; $chkThRename.VerticalAlignment  = "Center"; $chkThRename.Margin  = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThMerge   = New-Object System.Windows.Controls.CheckBox; $chkThMerge.Content   = "Merge";   $chkThMerge.IsChecked   = $false; $chkThMerge.Foreground   = Get-WpfColor "#CCCCCC"; $chkThMerge.VerticalAlignment   = "Center"; $chkThMerge.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThSlice   = New-Object System.Windows.Controls.CheckBox; $chkThSlice.Content   = "Slice";   $chkThSlice.IsChecked   = $false; $chkThSlice.Foreground   = Get-WpfColor "#CCCCCC"; $chkThSlice.VerticalAlignment   = "Center"; $chkThSlice.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThExtract = New-Object System.Windows.Controls.CheckBox; $chkThExtract.Content = "Extract"; $chkThExtract.IsChecked = $false; $chkThExtract.Foreground = Get-WpfColor "#CCCCCC"; $chkThExtract.VerticalAlignment = "Center"; $chkThExtract.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThImage   = New-Object System.Windows.Controls.CheckBox; $chkThImage.Content   = "Image";   $chkThImage.IsChecked   = $false; $chkThImage.Foreground   = Get-WpfColor "#CCCCCC"; $chkThImage.VerticalAlignment   = "Center"; $chkThImage.Margin   = New-Object System.Windows.Thickness(0,0,20,0)

    $btnThSelAll   = New-Object System.Windows.Controls.Button; $btnThSelAll.Content   = "Select All";   $btnThSelAll.Background   = Get-WpfColor "#2A2C35"; $btnThSelAll.Foreground   = Get-WpfColor "#FFFFFF"; $btnThSelAll.Width   = 85;  $btnThSelAll.Height   = 25; $btnThSelAll.BorderThickness   = 0; $btnThSelAll.Cursor   = [System.Windows.Input.Cursors]::Hand; $btnThSelAll.Margin   = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThDeselAll = New-Object System.Windows.Controls.Button; $btnThDeselAll.Content = "Deselect All"; $btnThDeselAll.Background = Get-WpfColor "#2A2C35"; $btnThDeselAll.Foreground = Get-WpfColor "#FFFFFF"; $btnThDeselAll.Width = 85;  $btnThDeselAll.Height = 25; $btnThDeselAll.BorderThickness = 0; $btnThDeselAll.Cursor = [System.Windows.Input.Cursors]::Hand; $btnThDeselAll.Margin = New-Object System.Windows.Thickness(0,0,20,0)
    $btnThRevert   = New-Object System.Windows.Controls.Button; $btnThRevert.Content   = "Revert Merge"; $btnThRevert.Background   = Get-WpfColor "#D95F5F"; $btnThRevert.Foreground   = Get-WpfColor "#FFFFFF"; $btnThRevert.Width   = 110; $btnThRevert.Height   = 25; $btnThRevert.BorderThickness   = 0; $btnThRevert.Cursor   = [System.Windows.Input.Cursors]::Hand; $btnThRevert.Margin   = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThProcess  = New-Object System.Windows.Controls.Button; $btnThProcess.Content  = "Process Theme"; $btnThProcess.Background  = Get-WpfColor "#4CAF72"; $btnThProcess.Foreground  = Get-WpfColor "#FFFFFF"; $btnThProcess.Width  = 115; $btnThProcess.Height  = 25; $btnThProcess.BorderThickness  = 0; $btnThProcess.Cursor  = [System.Windows.Input.Cursors]::Hand; $btnThProcess.Margin  = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThRefresh  = New-Object System.Windows.Controls.Button; $btnThRefresh.Content  = "Refresh Theme"; $btnThRefresh.Background  = Get-WpfColor "#5A78C4"; $btnThRefresh.Foreground  = Get-WpfColor "#FFFFFF"; $btnThRefresh.Width  = 115; $btnThRefresh.Height  = 25; $btnThRefresh.BorderThickness  = 0; $btnThRefresh.Cursor  = [System.Windows.Input.Cursors]::Hand

    $themeBarStack.Children.Add($chkThRename)  | Out-Null
    $themeBarStack.Children.Add($chkThMerge)   | Out-Null
    $themeBarStack.Children.Add($chkThSlice)   | Out-Null
    $themeBarStack.Children.Add($chkThExtract) | Out-Null
    $themeBarStack.Children.Add($chkThImage)   | Out-Null
    $themeBarStack.Children.Add($btnThSelAll)   | Out-Null
    $themeBarStack.Children.Add($btnThDeselAll) | Out-Null
    $themeBarStack.Children.Add($btnThRevert)   | Out-Null
    $themeBarStack.Children.Add($btnThProcess)  | Out-Null
    $themeBarStack.Children.Add($btnThRefresh)  | Out-Null
    $themeBar.Child = $themeBarStack
    $gpStack.Children.Add($themeBar) | Out-Null

    $parentListStack = New-Object System.Windows.Controls.StackPanel
    $parentListStack.Margin = New-Object System.Windows.Thickness(15)
    $gpStack.Children.Add($parentListStack) | Out-Null
    $gpJob.ParentListStack = $parentListStack

    foreach ($pKey in $parentDict.Keys) {
        $pJob = Build-PJob $pKey $parentDict[$pKey] $gpJob
        $gpJob.Parents.Add($pJob) | Out-Null
    }
    Update-GpFileCount $gpJob

    # --- THEME TASK BAR HANDLERS (wired after Parents are populated) ---
    $chkThRename.Tag = $gpJob
    $chkThRename.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkRename.IsChecked = $s } })

    $chkThMerge.Tag = $gpJob
    $chkThMerge.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { if ($p.ChkMerge.IsEnabled) { $p.ChkMerge.IsChecked = $s } } })

    $chkThSlice.Tag = $gpJob
    $chkThSlice.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkSlice.IsChecked = $s } })

    $chkThExtract.Tag = $gpJob
    $chkThExtract.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkExtract.IsChecked = $s } })

    $chkThImage.Tag = $gpJob
    $chkThImage.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkImage.IsChecked = $s } })

    $btnThSelAll.Tag = @{ GpJob = $gpJob; Chks = @{ Rename = $chkThRename; Merge = $chkThMerge; Slice = $chkThSlice; Extract = $chkThExtract; Image = $chkThImage } }
    $btnThSelAll.Add_Click({
        $t = $this.Tag
        foreach ($p in $t.GpJob.Parents) {
            $p.ChkRename.IsChecked = $true; $p.ChkSlice.IsChecked = $true
            $p.ChkExtract.IsChecked = $true; $p.ChkImage.IsChecked = $true
            if ($p.ChkMerge.IsEnabled) { $p.ChkMerge.IsChecked = $true }
        }
        $t.Chks.Rename.IsChecked = $true; $t.Chks.Merge.IsChecked = $true
        $t.Chks.Slice.IsChecked = $true; $t.Chks.Extract.IsChecked = $true; $t.Chks.Image.IsChecked = $true
    })

    $btnThDeselAll.Tag = @{ GpJob = $gpJob; Chks = @{ Rename = $chkThRename; Merge = $chkThMerge; Slice = $chkThSlice; Extract = $chkThExtract; Image = $chkThImage } }
    $btnThDeselAll.Add_Click({
        $t = $this.Tag
        foreach ($p in $t.GpJob.Parents) {
            $p.ChkRename.IsChecked = $false; $p.ChkMerge.IsChecked = $false; $p.ChkSlice.IsChecked = $false
            $p.ChkExtract.IsChecked = $false; $p.ChkImage.IsChecked = $false; $p.ChkLogs.IsChecked = $false
        }
        $t.Chks.Rename.IsChecked = $false; $t.Chks.Merge.IsChecked = $false
        $t.Chks.Slice.IsChecked = $false; $t.Chks.Extract.IsChecked = $false; $t.Chks.Image.IsChecked = $false
    })

    $btnThRevert.Tag = $gpJob
    $btnThRevert.Add_Click({
        $gp = $this.Tag
        $batPath = Join-Path $scriptDir "..\callers\RevertMerge.bat"
        if (-not (Test-Path $batPath)) { [System.Windows.MessageBox]::Show("RevertMerge.bat not found.", "Error") | Out-Null; return }
        $targets = @()
        foreach ($p in $gp.Parents) {
            if ($p.BtnRevertMerge -and $p.BtnRevertMerge.IsEnabled) {
                $tp = if ($p.ProcessedAnchorPath -ne "") { $p.ProcessedAnchorPath } else { $p.AnchorFile.FullName }
                if ($tp) { $targets += $tp }
            }
        }
        if ($targets.Count -eq 0) { return }
        foreach ($tp in $targets) {
            $argList = '/c ""' + $batPath + '" "' + $tp + '""'
            $proc = Start-Process -FilePath "cmd.exe" -ArgumentList $argList -WindowStyle Hidden -PassThru
            $timeout = 100
            while (-not $proc.HasExited -and $timeout -gt 0) {
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
                Start-Sleep -Milliseconds 100; $timeout--
            }
            if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
        }
        foreach ($p in @($gp.Parents)) { Refresh-PJob $p $gp }
    })

    $btnThProcess.Tag = $gpJob
    $btnThProcess.Add_Click({
        $gp = $this.Tag
        foreach ($p in $gp.Parents) { Enqueue-PJob $p $gp }
        if ($script:activeProcess -eq $null -and $script:processQueue.Count -gt 0) { Start-NextProcess }
    })

    $btnThRefresh.Tag = $gpJob
    $btnThRefresh.Add_Click({
        $gp = $this.Tag
        foreach ($p in @($gp.Parents)) { Refresh-PJob $p $gp }
    })

    $cbTheme.Tag = $gpJob
    $cbTheme.Add_SelectionChanged({ foreach ($p in $this.Tag.Parents) { Update-ParentPreview $p $this.Tag } })
    $cbPrefix.Tag = $gpJob
    $cbPrefix.Add_SelectionChanged({ foreach ($p in $this.Tag.Parents) { Update-ParentPreview $p $this.Tag } })
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
                        $pJob.CardStatusLabel.Text = $txt
                        $pJob.PickStatusLabel.Text = $txt
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
                            $pJob.FinishedOverlay.Visibility = "Visible"
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
            $files = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue | Sort-Object { switch -Regex ($_.Name) { 'Final\.3mf$' {0} 'Nest\.3mf$' {1} 'Full\.3mf$' {2} 'Full\.gcode\.3mf$' {3} default {4} } }, Name
            foreach ($fi in $files) { Add-FileRow $pJob $gpJob $fi }
            Update-ParentPreview $pJob $gpJob

            # Update UI to KEEP/REVERT state
            $pJob.ProcessingOverlay.Visibility = "Collapsed"
            $pJob.PickProcessingOverlay.Visibility = "Collapsed"
            $pJob.RowPanel.IsEnabled = $true
            $pJob.IsDone = $true
            $pJob.ChkRename.IsEnabled = $true; $pJob.ChkMerge.IsEnabled = $true; $pJob.ChkSlice.IsEnabled = $true
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
$btnBrowse.Add_Click({
    # 1. Grab the native Window Handle directly from the global $window object
    $interop = New-Object System.Windows.Interop.WindowInteropHelper($window)
    $hwnd = $interop.Handle

    # 2. Call the new NativeFolderBrowser (Now returns an array of paths!)
    $selectedPaths = [NativeFolderBrowser]::ShowDialog($hwnd, "Select folders containing Full.3mf files")

    # 3. Stop if the user closed the window or clicked Cancel
    if ($null -eq $selectedPaths -or $selectedPaths.Count -eq 0) { return }

    $lblGlobalTitle.Text = "Scanning selected folders..."
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    # 4. Standard scanning & queueing logic (Looping through every folder you checked!)
    $newFound = @()
    foreach ($path in $selectedPaths) {
        $newFound += @(Get-ChildItem -Path $path -Filter "*Full.3mf" -Recurse -File -ErrorAction SilentlyContinue)
    }

    $newGpQueue = [ordered]@{}
    foreach ($f in $newFound) {
        $parentPath = $f.DirectoryName
        $gp = $f.Directory.Parent
        $gpPath = if ($gp) { $gp.FullName } else { "ROOT_" + $parentPath }

        # Prevent duplicates
        $exists = $false
        foreach ($j in $script:jobs) {
            foreach ($parentJob in $j.Parents) {
                if ($parentJob.FolderPath -eq $parentPath) { $exists = $true; break }
            }
        }
        if ($exists) { continue }

        if (-not $newGpQueue.Contains($gpPath)) { $newGpQueue[$gpPath] = [ordered]@{} }
        if (-not $newGpQueue[$gpPath].Contains($parentPath)) { $newGpQueue[$gpPath][$parentPath] = $f }
    }

    foreach ($gpPath in $newGpQueue.Keys) {
        $existingGp = $null
        foreach ($j in $script:jobs) { if ($j.GpPath -eq $gpPath) { $existingGp = $j; break } }

        if ($existingGp) {
            foreach ($pKey in $newGpQueue[$gpPath].Keys) {
                $pJob = Build-PJob $pKey $newGpQueue[$gpPath][$pKey] $existingGp
                $existingGp.Parents.Add($pJob) | Out-Null
            }
            Update-GpFileCount $existingGp
        } else {
            Build-GpJob $gpPath $newGpQueue[$gpPath]
        }
    }

    $lblGlobalTitle.Text = "Queue Dashboard ($($script:jobs.Count) Theme(s) found)"
    if ($script:jobs.Count -gt 0) { Update-GlobalProcessAllStatus }
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

    $newGpQueue = Get-AnchorQueue $dropped

    if ($newGpQueue.Count -gt 0) {
        # Remove folders already loaded in the UI
        foreach ($gpPath in @($newGpQueue.Keys)) {
            foreach ($parentPath in @($newGpQueue[$gpPath].Keys)) {
                $exists = $false
                foreach ($j in $script:jobs) { foreach ($parentJob in $j.Parents) { if ($parentJob.FolderPath -eq $parentPath) { $exists = $true; break } } }
                if ($exists) { $newGpQueue[$gpPath].Remove($parentPath) }
            }
            if ($newGpQueue[$gpPath].Count -eq 0) { $newGpQueue.Remove($gpPath) }
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

$btnProcessAll.Add_Click({
    foreach ($gpJob in $script:jobs) {
        foreach ($pJob in $gpJob.Parents) {
            Enqueue-PJob $pJob $gpJob
        }
    }
    if ($script:activeProcess -eq $null -and $script:processQueue.Count -gt 0) {
        Start-NextProcess
    }
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