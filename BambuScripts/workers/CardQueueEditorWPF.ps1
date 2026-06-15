# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# FULL WPF BATCH PRE-FLIGHT EDITOR ENGINE
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'

# â”€â”€ DEBUG LOGGING â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
$script:DebugLog = Join-Path $PSScriptRoot "CardQueueEditor_debug.txt"
Set-Content -Path $script:DebugLog -Value "=== CardQueueEditor Debug Log  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" -Encoding UTF8
function Write-Log {
    param([string]$msg, [string]$level = "INFO")
    $ts  = (Get-Date).ToString("HH:mm:ss.fff")
    $line = "[$ts][$level] $msg"
    Add-Content -Path $script:DebugLog -Value $line -Encoding UTF8
    Write-Host $line
}
Start-Transcript -Path ($script:DebugLog -replace '\.txt$', '_transcript.txt') -Force | Out-Null
Write-Log "Script started"

[System.AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $e)
    $ex = $e.ExceptionObject
    Write-Log "APPDOMAIN UNHANDLED: $($ex.GetType().FullName): $($ex.Message)" "FATAL"
    if ($ex.InnerException) { Write-Log "  INNER: $($ex.InnerException.Message)" "FATAL" }
    Write-Log "  STACK: $($ex.StackTrace)" "FATAL"
})

# --- 1. ENVIRONMENT & LIBRARY ---
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = [System.Environment]::CurrentDirectory }
$colorCsvPath = Join-Path $scriptDir "..\libraries\FilamentLibrary.csv"
$namesLibPath  = Join-Path $scriptDir "..\libraries\NamesLibrary.ps1"
$purgeDictPath = Join-Path $scriptDir "..\libraries\PurgeDictionary.csv"

# Row type for the Purge Dictionary grid â€” tracks original values so edited
# cells can be highlighted and the Save button enabled only when something changed.
Add-Type -TypeDefinition @"
using System.ComponentModel;

public class PurgeDictRow : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    private void Raise(string name) {
        var h = PropertyChanged;
        if (h != null) h(this, new PropertyChangedEventArgs(name));
    }

    private string _source = "";
    private string _target = "";
    private bool   _tuned;
    private string _tunedVolume = "";
    private string _baseVolume = "";
    private string _sourceHex = "#3A3C4A";
    private string _targetHex = "#3A3C4A";

    private string _origSource = "";
    private string _origTarget = "";
    private bool   _origTuned;
    private string _origTunedVolume = "";
    private string _origBaseVolume = "";

    // Per-field saved-to-disk markers: true after a save, cleared when that field is edited again
    private bool _savedFrom, _savedTo, _savedTuned, _savedTunedVolume, _savedBaseVolume;

    public string Source_Filament {
        get { return _source; }
        set { _source = value; _savedFrom = false; Raise("Source_Filament"); Raise("FromDirty"); Raise("FromSaved"); Raise("IsDirty"); }
    }
    public string Target_Filament {
        get { return _target; }
        set { _target = value; _savedTo = false; Raise("Target_Filament"); Raise("ToDirty"); Raise("ToSaved"); Raise("IsDirty"); }
    }
    public bool Tuned {
        get { return _tuned; }
        set { _tuned = value; _savedTuned = false; Raise("Tuned"); Raise("TunedDirty"); Raise("TunedSaved"); Raise("IsDirty"); Raise("Savings_Pct"); }
    }
    public string Tuned_Volume {
        get { return _tunedVolume; }
        set { _tunedVolume = value; _savedTunedVolume = false; Raise("Tuned_Volume"); Raise("TunedVolumeDirty"); Raise("TunedVolumeSaved"); Raise("IsDirty"); Raise("Savings_Pct"); }
    }
    public string Base_Volume {
        get { return _baseVolume; }
        set { _baseVolume = value; _savedBaseVolume = false; Raise("Base_Volume"); Raise("BaseVolumeDirty"); Raise("BaseVolumeSaved"); Raise("IsDirty"); Raise("Savings_Pct"); }
    }
    public string Source_Hex { get { return _sourceHex; } set { _sourceHex = value; Raise("Source_Hex"); } }
    public string Target_Hex { get { return _targetHex; } set { _targetHex = value; Raise("Target_Hex"); } }

    // Percent of purge volume saved by the tuned value vs. the untuned base - blank for untuned/unparseable rows
    public string Savings_Pct {
        get {
            double baseV, tunedV;
            if (!_tuned) return "";
            if (!double.TryParse(_baseVolume, out baseV) || !double.TryParse(_tunedVolume, out tunedV)) return "";
            if (baseV <= 0) return "";
            double pct = (baseV - tunedV) / baseV * 100.0;
            return pct.ToString("0.0") + "%";
        }
    }
    public double? SavingsValue {
        get {
            double baseV, tunedV;
            if (!_tuned) return null;
            if (!double.TryParse(_baseVolume, out baseV) || !double.TryParse(_tunedVolume, out tunedV)) return null;
            if (baseV <= 0) return null;
            return (baseV - tunedV) / baseV * 100.0;
        }
    }

    public bool FromDirty        { get { return _source != _origSource; } }
    public bool ToDirty          { get { return _target != _origTarget; } }
    public bool TunedDirty       { get { return _tuned != _origTuned; } }
    public bool TunedVolumeDirty { get { return _tunedVolume != _origTunedVolume; } }
    public bool BaseVolumeDirty  { get { return _baseVolume != _origBaseVolume; } }
    public bool IsDirty {
        get { return FromDirty || ToDirty || TunedDirty || TunedVolumeDirty || BaseVolumeDirty; }
    }

    public bool FromSaved        { get { return _savedFrom; } }
    public bool ToSaved          { get { return _savedTo; } }
    public bool TunedSaved       { get { return _savedTuned; } }
    public bool TunedVolumeSaved { get { return _savedTunedVolume; } }
    public bool BaseVolumeSaved  { get { return _savedBaseVolume; } }

    public void Load(string source, string target, bool tuned, string tunedVolume, string baseVolume, string sourceHex, string targetHex) {
        _source = source; _target = target; _tuned = tuned; _tunedVolume = tunedVolume; _baseVolume = baseVolume;
        _origSource = source; _origTarget = target; _origTuned = tuned; _origTunedVolume = tunedVolume; _origBaseVolume = baseVolume;
        _savedFrom = false; _savedTo = false; _savedTuned = false; _savedTunedVolume = false; _savedBaseVolume = false;
        _sourceHex = sourceHex; _targetHex = targetHex;
    }

    public void RevertToBaseline() {
        _source      = _origSource;      _savedFrom        = false;
        _target      = _origTarget;      _savedTo          = false;
        _tuned       = _origTuned;       _savedTuned       = false;
        _tunedVolume = _origTunedVolume; _savedTunedVolume = false;
        _baseVolume  = _origBaseVolume;  _savedBaseVolume  = false;
        Raise("Source_Filament"); Raise("Target_Filament"); Raise("Tuned");
        Raise("Tuned_Volume"); Raise("Base_Volume"); Raise("Savings_Pct");
        Raise("FromDirty");        Raise("ToDirty");          Raise("TunedDirty");
        Raise("TunedVolumeDirty"); Raise("BaseVolumeDirty");  Raise("IsDirty");
        Raise("FromSaved");        Raise("ToSaved");          Raise("TunedSaved");
        Raise("TunedVolumeSaved"); Raise("BaseVolumeSaved");
    }

    public void CommitBaseline() {
        bool chFrom  = _source      != _origSource;
        bool chTo    = _target      != _origTarget;
        bool chTuned = _tuned       != _origTuned;
        bool chTV    = _tunedVolume != _origTunedVolume;
        bool chBV    = _baseVolume  != _origBaseVolume;
        _origSource = _source; _origTarget = _target; _origTuned = _tuned;
        _origTunedVolume = _tunedVolume; _origBaseVolume = _baseVolume;
        if (chFrom)  _savedFrom         = true;
        if (chTo)    _savedTo           = true;
        if (chTuned) _savedTuned        = true;
        if (chTV)    _savedTunedVolume  = true;
        if (chBV)    _savedBaseVolume   = true;
        Raise("FromDirty");        Raise("ToDirty");          Raise("TunedDirty");
        Raise("TunedVolumeDirty"); Raise("BaseVolumeDirty");  Raise("IsDirty");
        Raise("FromSaved");        Raise("ToSaved");          Raise("TunedSaved");
        Raise("TunedVolumeSaved"); Raise("BaseVolumeSaved");
    }
}
"@

$script:LibraryColors = [ordered]@{}
$script:HexToName = @{}
$script:LibraryGradients = [ordered]@{}
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
                $script:LibraryColors[$name] = $hex
                $script:HexToName[$hex] = $name
                $script:HexToName[$hex.Substring(0,7)] = $name
                if ($parts.Count -gt 4) {
                    $gradVals = @()
                    for ($gi = 4; $gi -lt $parts.Count; $gi++) {
                        $gradVals += $parts[$gi].Replace('"','').Trim()
                    }
                    if (($gradVals | Where-Object { $_ -match '^#' }).Count -gt 0) {
                        $script:LibraryGradients[$name] = $gradVals
                    }
                }
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

# SKU validation pattern â€” used everywhere a SKU is read, displayed, or saved.
# Accepts any non-whitespace string of 3 or more characters so numeric SKUs
# (e.g. 00096), dash-format SKUs (e.g. fud-sw-2623), and any future formats
# all pass without code changes.  The old-format TSV guard (Date at col 4)
# is the real safety net against Theme text being mistaken for a SKU.
$script:SkuPattern = '^[^\s]{3,}$'

function Get-WpfColor([string]$hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { $hex = "#808080" }
    if ($hex.Length -eq 9) { $hex = "#" + $hex.Substring(1,6) }
    try { return [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($hex)) }
    catch { return [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Colors]::Gray) }
}

function Remove-StaleFile([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    if ($path -inotmatch '_Data\.tsv$') {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        return
    }
    # _Data.tsv: never delete outright if it carries a SKU - shrink it to a
    # SKU-only stub instead so DataExtract_worker re-seeds the SKU on rebuild.
    $sku = ""
    try {
        $line = Get-Content -LiteralPath $path -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($line) {
            $cols = $line -split "`t"
            $isOldFormat = $cols.Count -gt 4 -and $cols[4] -match '^\d{1,2}/\d{1,2}/\d{4}$'
            if (-not $isOldFormat -and $cols.Count -ge 4 -and $cols[3].Trim() -match '^[^\s]{3,}$') {
                $sku = $cols[3].Trim()
            }
        }
    } catch {}
    if ($sku -ne "") {
        Set-Content -Path $path -Value "`t`t`t$sku" -Encoding UTF8 -NoNewline -Force
    } else {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    }
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

# --- HSV / RGB helpers (used by Libraries filament picker) ---
function Hsv-To-Rgb([double]$h,[double]$s,[double]$v) {
    $h = (($h % 360) + 360) % 360
    if ($s -lt 1e-9) { $c=[int]([Math]::Round($v*255)); return [int[]]($c,$c,$c) }
    $hi=[int]($h/60)%6; $f=$h/60-[int]($h/60)
    $p=$v*(1-$s); $q=$v*(1-$f*$s); $t=$v*(1-(1-$f)*$s)
    $rgb=switch($hi){0{@($v,$t,$p)}1{@($q,$v,$p)}2{@($p,$v,$t)}3{@($p,$q,$v)}4{@($t,$p,$v)}5{@($v,$p,$q)}}
    return [int[]]([Math]::Max(0,[Math]::Min(255,[int][Math]::Round($rgb[0]*255))),
                   [Math]::Max(0,[Math]::Min(255,[int][Math]::Round($rgb[1]*255))),
                   [Math]::Max(0,[Math]::Min(255,[int][Math]::Round($rgb[2]*255))))
}
function Rgb-To-Hsv([int]$r,[int]$g,[int]$b) {
    $rf=$r/255.0;$gf=$g/255.0;$bf=$b/255.0
    $max=[Math]::Max($rf,[Math]::Max($gf,$bf)); $min=[Math]::Min($rf,[Math]::Min($gf,$bf))
    $d=$max-$min; $v=$max; $s=if($max-gt 1e-9){$d/$max}else{0.0}; $h=0.0
    if($d-gt 1e-9){
        if($max-eq $rf){$h=60*(($gf-$bf)/$d%6)}
        elseif($max-eq $gf){$h=60*(($bf-$rf)/$d+2)}
        else{$h=60*(($rf-$gf)/$d+4)}
    }
    if($h-lt 0){$h+=360}
    return [double[]]($h,$s,$v)
}
function HsvPure-HexAt([double]$h) {
    $rgb=Hsv-To-Rgb $h 1.0 1.0
    return "#{0:X2}{1:X2}{2:X2}" -f $rgb[0],$rgb[1],$rgb[2]
}

function Extract-3mfPickImage([string]$mfPath, [string]$outDir, [string]$prefix = "") {
    if (-not (Test-Path -LiteralPath $mfPath)) { return $null }
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($mfPath)
        foreach ($candidate in @("Metadata/top_1.png","Metadata/pick_1.png","Metadata/plate_1.png")) {
            $entry = $zip.Entries | Where-Object { ($_.FullName -replace '\\','/') -eq $candidate } | Select-Object -First 1
            if ($null -ne $entry) {
                $outFile = if ($prefix) { "${prefix}_$([System.IO.Path]::GetFileName($candidate))" } else { [System.IO.Path]::GetFileName($candidate) }
                $outPath = Join-Path $outDir $outFile
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $outPath, $true)
                $zip.Dispose()
                return $outPath
            }
        }
        $zip.Dispose()
    } catch { try { $zip.Dispose() } catch {} }
    return $null
}

# --- Re-Nest axis-picker helpers (inline orientation overlay, no popup) ---
function Read-ZipEntryBytesInline([string]$zipPath, [string]$entryName) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entry = $zip.GetEntry($entryName)
        if ($null -eq $entry) { return $null }
        $ms = New-Object System.IO.MemoryStream
        $entry.Open().CopyTo($ms)
        return $ms.ToArray()
    } finally { $zip.Dispose() }
}
function Read-ZipEntryTextInline([string]$zipPath, [string]$entryName) {
    $bytes = Read-ZipEntryBytesInline $zipPath $entryName
    if ($null -eq $bytes) { return $null }
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

# Locate the centermost (non-wipe-tower) build-item instance on the Nest plate
# and return its row-vector rotation matrix R and plate position.
function Get-CenterMostInstance([string]$nestPath) {
    $bedMM = 256.0
    $bedCenter = $bedMM / 2.0
    $plateJsonText = Read-ZipEntryTextInline $nestPath "Metadata/plate_1.json"
    $bestCx = $bedCenter; $bestCy = $bedCenter

    if ($null -ne $plateJsonText) {
        $plateJson = $plateJsonText | ConvertFrom-Json
        $best = $null; $bestDist = [double]::MaxValue
        foreach ($o in $plateJson.bbox_objects) {
            if ($o.id -eq 1000) { continue }
            $cx = ($o.bbox[0] + $o.bbox[2]) / 2
            $cy = ($o.bbox[1] + $o.bbox[3]) / 2
            $d = [Math]::Sqrt(($cx-$bedCenter)*($cx-$bedCenter) + ($cy-$bedCenter)*($cy-$bedCenter))
            if ($d -lt $bestDist) { $bestDist = $d; $best = $o; $bestCx = $cx; $bestCy = $cy }
        }
        if ($null -eq $best) { return $null }
    }
    # If plate_1.json is absent (e.g. some Full.3mf files), fall back to the
    # build item closest to the bed center directly.

    $modelText = Read-ZipEntryTextInline $nestPath "3D/3dmodel.model"
    if ($null -eq $modelText) { return $null }
    $itemMatches = [regex]::Matches($modelText, '<item objectid="(\d+)"[^>]*transform="([^"]+)"[^>]*/>')
    $bestItem = $null; $bestItemDist = [double]::MaxValue
    foreach ($m in $itemMatches) {
        $vals = $m.Groups[2].Value -split '\s+' | ForEach-Object {[double]$_}
        if ($vals.Count -lt 12) { continue }
        $tx = $vals[9]; $ty = $vals[10]
        $d = [Math]::Sqrt(($tx-$bestCx)*($tx-$bestCx) + ($ty-$bestCy)*($ty-$bestCy))
        if ($d -lt $bestItemDist) { $bestItemDist = $d; $bestItem = $vals }
    }
    if ($null -eq $bestItem) { return $null }
    return @{ R = [double[]]($bestItem[0..8]); tx = $bestItem[9]; ty = $bestItem[10] }
}

# Final.3mf has no Metadata/plate_1.json (it's an isolated single-object file) -
# just take its (sole) build item's transform directly.
function Get-FinalInstance([string]$finalPath) {
    $modelText = Read-ZipEntryTextInline $finalPath "3D/3dmodel.model"
    if ($null -eq $modelText) { return $null }
    $itemMatches = [regex]::Matches($modelText, '<item objectid="(\d+)"[^>]*transform="([^"]+)"[^>]*/>')
    if ($itemMatches.Count -eq 0) { return $null }
    $vals = $itemMatches[0].Groups[2].Value -split '\s+' | ForEach-Object {[double]$_}
    if ($vals.Count -lt 12) { return $null }
    return @{ R = [double[]]($vals[0..8]); tx = $vals[9]; ty = $vals[10] }
}

# --- 3x3 row-vector (v' = v*R) rotation matrix helpers, mirroring RenestFromFinal_worker.ps1 ---
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
# Rotation angle (degrees) represented by a pure rotation matrix, via its trace
function Get-RotationAngleDeg([double[]]$r) {
    $cosA = ($r[0] + $r[4] + $r[8] - 1.0) / 2.0
    if ($cosA -gt 1.0) { $cosA = 1.0 } elseif ($cosA -lt -1.0) { $cosA = -1.0 }
    return [Math]::Acos($cosA) * 180.0 / [Math]::PI
}

# Yaw angle (degrees) of the local +X axis (row 0 of R), projected onto the
# world XY plane, measured from world +X. This picker only ever rotates about
# the world "up" (Z) axis, so this single angle fully describes an instance's
# in-plane orientation for our purposes.
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
# the world up axis - the part Re-Nest must re-apply to the baked Final.
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

# Bounding-box center of the object's silhouette (vs. the corner/background color)
function Get-ContentCenterInline([System.Drawing.Bitmap]$bmp) {
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

# Draw X(red)/Y(green)/Z(blue) axis arrows at a pixel position. Axes that point
# (nearly) straight into/out of the screen have no in-plane arrow to draw and
# are skipped.
function Draw-AxesOverlay([System.Drawing.Graphics]$g, [double]$cxPx, [double]$cyPx, [double[]]$R, [double]$scale, [double]$arrowLen) {
    $axes = @(
        @{ name="X"; color=[System.Drawing.Color]::Red;  px=$R[0]; py=$R[1] },
        @{ name="Y"; color=[System.Drawing.Color]::Lime; px=$R[3]; py=$R[4] },
        @{ name="Z"; color=[System.Drawing.Color]::Blue; px=$R[6]; py=$R[7] }
    )
    foreach ($axis in $axes) {
        $dx = $axis.px * $scale
        $dy = -1 * $axis.py * $scale
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

function Bitmap-To-WpfImageInline([System.Drawing.Bitmap]$bmp) {
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

# Draw the Nest's reference axes (Red=X, Green=Y, Blue=Z) at the centermost
# instance's plate position on the Nest Source preview.
function Update-NestRefAxesOverlay([hashtable]$t) {
    if ([string]::IsNullOrEmpty($t.OvSrcPath) -or -not (Test-Path -LiteralPath $t.OvSrcPath)) { return }
    if ([string]::IsNullOrEmpty($t.SourcePath) -or -not (Test-Path -LiteralPath $t.SourcePath)) { return }
    try {
        $inst = Get-CenterMostInstance $t.SourcePath
        if ($null -eq $inst) { return }
        $t.AxisR = $inst.R
        $scale = $t.AxisScale
        $bmp = [System.Drawing.Bitmap]::FromFile($t.OvSrcPath)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $cxPx = $inst.tx * $scale
        $cyPx = $bmp.Height - ($inst.ty * $scale)
        Draw-AxesOverlay $g $cxPx $cyPx $inst.R $scale 40
        $g.Dispose()
        $t.P.RnImgLeft.Source = Bitmap-To-WpfImageInline $bmp
        $bmp.Dispose()
    } catch {}
}

# Draw confirmation axes on the Final preview. The axis frame is FIXED at the
# Nest Source reference orientation (same frame drawn on the Nest Source);
# the IMAGE spins around the locator anchor by the proposed Re-Nest spin
# (nest angle + Rotate-90 steps). Since the correct spin satisfies
# R_fin * ZRot(spin) = R_src, the spun preview shows the source-reference
# pose exactly when the proposal is right - i.e. the user clicks Rotate 90
# until the fixed axis points at the same part of the model as the axis
# drawn on the Nest Source, which visually confirms the spin Re-Nest applies.
function Update-FinalCurrentAxesOverlay([hashtable]$t) {
    if ([string]::IsNullOrEmpty($t.OvFinPath) -or -not (Test-Path -LiteralPath $t.OvFinPath)) { return }
    if ([string]::IsNullOrEmpty($t.FinalPath) -or -not (Test-Path -LiteralPath $t.FinalPath)) { return }
    try {
        $inst = Get-FinalInstance $t.FinalPath
        if ($null -eq $inst) { return }
        $t.FinalAxisR = $inst.R
        $displayR = $inst.R
        $spinDeg = 0.0
        if ($null -ne $t.SrcRefRotNorm) {
            $displayR = $t.SrcRefRotNorm
            $spinDeg = $t.SrcNestAngleDeg + ($t.AxisYawSteps * 90)
        }
        $rawBmp = [System.Drawing.Bitmap]::FromFile($t.OvFinPath)
        if ($t.OvFinPath -match '(?i)(top|pick)_\d+\.png$') {
            # top_/pick_ previews are plate-mapped top views, so the build
            # item's translation - the object's local-origin "locator" - maps
            # directly to pixels. Anchor (and spin) THERE, not at the content
            # bounding-box center, which drifts as soon as the user adds or
            # moves a detached piece in the Final.
            $pxScale = $rawBmp.Width / 256.0
            $center = @(($inst.tx * $pxScale), ($rawBmp.Height - ($inst.ty * $pxScale)))
        } else {
            # Unknown camera (e.g. plate_1.png perspective render) - fall back
            # to the content bounding-box center.
            $center = Get-ContentCenterInline $rawBmp
        }
        # Spin the preview image about the anchor by the proposed world-Z
        # spin; the axis frame stays fixed. Negative because the preview maps
        # world +Y to screen-up while GDI+ RotateTransform is
        # clockwise-positive in screen space.
        if ([Math]::Abs($spinDeg % 360) -gt 1e-6) {
            $bmp = New-Object System.Drawing.Bitmap $rawBmp.Width, $rawBmp.Height
            $gs = [System.Drawing.Graphics]::FromImage($bmp)
            $gs.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $gs.TranslateTransform($center[0], $center[1])
            $gs.RotateTransform(-$spinDeg)
            $gs.TranslateTransform(-$center[0], -$center[1])
            $gs.DrawImage($rawBmp, 0, 0)
            $gs.Dispose()
            $rawBmp.Dispose()
        } else {
            $bmp = $rawBmp
        }
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $arrowLen = [Math]::Min($bmp.Width, $bmp.Height) / 6.0
        Draw-AxesOverlay $g $center[0] $center[1] $displayR $t.AxisScale $arrowLen
        $g.Dispose()
        $t.P.RnImgRight.Source = Bitmap-To-WpfImageInline $bmp
        $bmp.Dispose()
    } catch {}
}

# Recompute the preview for the reference instance's extra yaw (in 90-degree
# steps about the world "up"/Z axis only) and refresh the status label. Every
# Re-Nest clone keeps the Final's own (baked) orientation; only a world-Z
# spin is applied - the source's non-90 nest angle plus each instance's yaw
# relative to the reference instance, plus the extra Rotate-90 steps.
function Update-AxisPickerPreview([hashtable]$t) {
    $extraYawDeg = $t.AxisYawSteps * 90
    $t.PendingRotCorrection = Get-ZRotationMatrix $extraYawDeg

    $nestAngleTxt = "{0:F2}" -f $t.SrcNestAngleDeg
    $t.LblAxisRef.Text = "Click Rotate 90 to spin the Final preview until its fixed axis points at the SAME part of the model as the Nest Source's axis - that confirms the spin. " + `
        "Re-Nest then spins each clone about the world up axis only: nest angle $nestAngleTxt deg + extra $extraYawDeg deg + each instance's own yaw vs the reference."
}

# Load the orientation state for a Re-Nest card as soon as the card is built -
# no user action needed. Reads the Nest Source reference instance's rotation,
# derives the non-90 "nest angle" (its deviation about world Z from the
# nearest axis-aligned rotation - the part a Bambu trim/cut bake normalizes
# away), and restores any Rotate-90 offset from a legacy
# <stem>_RotCorrection.json if one is still sitting in the folder. The Final
# preview then immediately shows the spin Re-Nest will apply; Rotate 90 only
# adjusts it, and Run Re-Nest hands the on-screen state to the worker via a
# temp file (nothing is written into the design folder). Returns $false
# (silently) when the folder has no usable Final + Nest Source pair.
function Initialize-OrientationState([hashtable]$t) {
    if ([string]::IsNullOrEmpty($t.SourcePath) -or -not (Test-Path -LiteralPath $t.SourcePath)) { return $false }
    if ([string]::IsNullOrEmpty($t.FinalPath) -or -not (Test-Path -LiteralPath $t.FinalPath)) { return $false }
    $srcInst = Get-CenterMostInstance $t.SourcePath
    if ($null -eq $srcInst) { return $false }
    $finInst = Get-FinalInstance $t.FinalPath
    if ($null -eq $finInst) { return $false }
    $t.AxisR = $srcInst.R
    $t.FinalAxisR = $finInst.R
    $t.PendingSourceR = $srcInst.R
    $t.SrcRefRotNorm = Normalize-Rot3x3 $srcInst.R
    $t.SrcRefYawDeg = Get-YawAngleDeg $srcInst.R
    $t.SrcNestAngleDeg = Get-RelativeYawDeg $t.SrcRefRotNorm (Get-AxisSnappedRot $t.SrcRefRotNorm)
    $t.AxisYawSteps = 0
    try {
        $stemO = [System.IO.Path]::GetFileNameWithoutExtension($t.FinalPath) -replace '(?i)_Final$', ''
        $jsonPathO = Join-Path ([System.IO.Path]::GetDirectoryName($t.FinalPath)) ($stemO + "_RotCorrection.json")
        if (Test-Path -LiteralPath $jsonPathO) {
            $jo = Get-Content -LiteralPath $jsonPathO -Raw | ConvertFrom-Json
            if ($null -ne $jo.extraYawDeg) {
                $t.AxisYawSteps = ((([int][Math]::Round([double]$jo.extraYawDeg / 90)) % 4) + 4) % 4
            }
        }
    } catch {}

    Update-AxisPickerPreview $t
    Update-NestRefAxesOverlay $t
    Update-FinalCurrentAxesOverlay $t
    return $true
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

# Scan a string for a token that matches a known GpTheme, skipping leading
# printer-prefix and tag qualifiers.  Returns the canonical theme string from
# $script:GpThemes, or '' if none found.
function Find-ThemeMatch([string]$s) {
    $toks = $s -split '[_\-\s.]+' | Where-Object { $_ -ne '' }
    $i = 0
    while ($i -lt $toks.Count - 1 -and (
           $script:PrinterPrefixes -icontains $toks[$i] -or
           $script:Tags            -icontains $toks[$i])) { $i++ }
    for ($j = $i; $j -lt $toks.Count; $j++) {
        $clean = $toks[$j] -replace '[^a-zA-Z0-9]', ''
        $m = $script:GpThemes | Where-Object { ($_ -replace '[^a-zA-Z0-9]','') -ieq $clean } | Select-Object -First 1
        if ($m) { return $m }
    }
    return ''
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
    # Convention: Prefix_Char_Adj_Theme_Suffix â€” after stripping prefix+suffix, max 3 remain.
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
$script:editQueue = New-Object System.Collections.Queue
$script:editActiveJob = $null
. (Join-Path $PSScriptRoot "..\libraries\NamesLibrary.ps1")
$script:AdjPresets = @('Common','RARE','EPIC','LEGENDARY','Default')

function Find-AnchorFile($folderPath) {
    $files = Get-ChildItem -Path $folderPath -File -ErrorAction SilentlyContinue
    $f = $files | Where-Object { $_.Name -match '(?i)Full\.3mf$'  -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Name -match '(?i)Nest\.3mf$'  -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Name -match '(?i)Final\.3mf$' -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.3mf$'    -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.stl$' }                                               | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.png$'    -and $_.Name -notmatch '(?i)_slicePreview\.png$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    return $f
}

# Returns the best 3MF for color info only: Full > Nest > Final > any .3mf
# Used independently of AnchorFile so that color slots always reflect the
# richest source file even when the anchor was a Final or custom file.
function Find-ColorInfoFile($folderPath) {
    $files = Get-ChildItem -Path $folderPath -File -ErrorAction SilentlyContinue
    $f = $files | Where-Object { $_.Name -match '(?i)Full\.3mf$'  -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Name -match '(?i)Nest\.3mf$'  -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Name -match '(?i)Final\.3mf$' -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f }
    $f = $files | Where-Object { $_.Extension -imatch '\.3mf$'    -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
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
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Background="#1C1D23" Grid.Row="0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="60"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Name="LblGlobalTitle" Text="Loading files into queue..." Foreground="White" FontSize="18" FontWeight="Bold" VerticalAlignment="Center" Margin="15,0,0,0"/>

                <StackPanel Grid.Row="0" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center" Orientation="Vertical">
                    <Button Name="BtnBrowse" Content="Browse Files" Background="#5A78C4" Foreground="White" FontWeight="Bold" Width="140" Height="30" BorderThickness="0" Cursor="Hand"/>
                    <TextBlock Text="Browse or drop files to add" Foreground="#888888" FontSize="10" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                </StackPanel>

                <StackPanel Grid.Row="0" Grid.Column="2" HorizontalAlignment="Right" VerticalAlignment="Center" Orientation="Horizontal" Margin="0,0,15,0">
                    <Button Name="BtnImportSkus" Content="Import SKUs" Background="#7A5C2E" Foreground="White" FontWeight="Bold" Width="110" Height="30" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0"/>
                    <Button Name="BtnStopQueue" Content="Stop Queue" Background="#8B3A3A" Foreground="White" FontWeight="Bold" Width="120" Height="30" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0"/>
                    <Button Name="BtnProcessAll" Content="Process All Tasks" Background="#4CAF72" Foreground="White" FontWeight="Bold" Width="150" Height="30" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>

                <Border Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3" Background="#14151A" BorderBrush="#2A2C35" BorderThickness="0,1,0,0" Padding="12,6,12,6">
                    <StackPanel Name="TopModeBar" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
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

# Hook WPF dispatcher so unhandled exceptions on the UI thread are logged
$window.Dispatcher.add_UnhandledException({
    param($sender, $e)
    $ex = $e.Exception
    Write-Log "DISPATCHER UNHANDLED: $($ex.GetType().FullName): $($ex.Message)" "FATAL"
    if ($ex.InnerException) { Write-Log "  INNER: $($ex.InnerException.Message)" "FATAL" }
    Write-Log "  STACK: $($ex.StackTrace)" "FATAL"
    $e.Handled = $true   # prevent WPF from crashing the process so the log can flush
})
Write-Log "Window created, dispatcher hooked"

$lblGlobalTitle = $window.FindName("LblGlobalTitle")
$btnProcessAll  = $window.FindName("BtnProcessAll")
$btnStopQueue   = $window.FindName("BtnStopQueue")
$btnImportSkus  = $window.FindName("BtnImportSkus")
$btnBrowse      = $window.FindName("BtnBrowse") # <--- ADD THIS LINE BACK
$mainStack      = $window.FindName("MainStack")
$topModeBar     = $window.FindName("TopModeBar")
$scrollViewer   = $mainStack.Parent   # ScrollViewer wrapping MainStack
$script:LibrariesPanel = $null        # built later by Build-LibrariesPanel

# Global workspace mode: "FilePr" | "Editing" | "Review"
$script:GlobalMode = "FilePr"

function Get-PJobHasSelectedTask($pJob) {
    foreach ($name in @('ChkRename','ChkMerge','ChkSlice','ChkExtract','ChkImage','ChkColors','ChkLogs','ChkBOD','ChkPrintQ')) {
        $chk = $pJob.$name
        if ($null -ne $chk -and [bool]$chk.IsChecked) { return $true }
    }
    return $false
}

function Update-ThProcessButtonState($gpJob) {
    if ($null -eq $gpJob -or $null -eq $gpJob.BtnThProcess) { return }
    $hasSelected = $false
    foreach ($p in $gpJob.Parents) {
        if ($p.IsQueued -or $p.IsDone) { continue }
        if (Get-PJobHasSelectedTask $p) { $hasSelected = $true; break }
    }
    if ($hasSelected) {
        $gpJob.BtnThProcess.IsEnabled = $true; $gpJob.BtnThProcess.Background = Get-WpfColor "#4CAF72"; $gpJob.BtnThProcess.Foreground = Get-WpfColor "#FFFFFF"
    } else {
        $gpJob.BtnThProcess.IsEnabled = $false; $gpJob.BtnThProcess.Background = Get-WpfColor "#3A3A3A"; $gpJob.BtnThProcess.Foreground = Get-WpfColor "#666666"
    }
}

function Update-ThRevertButtonState($gpJob) {
    if ($null -eq $gpJob -or $null -eq $gpJob.BtnThRevert) { return }
    $hasRevertible = $false
    foreach ($p in $gpJob.Parents) {
        if ($p.BtnRevertMerge -and $p.BtnRevertMerge.IsEnabled) { $hasRevertible = $true; break }
    }
    if ($hasRevertible) {
        $gpJob.BtnThRevert.IsEnabled = $true; $gpJob.BtnThRevert.Background = Get-WpfColor "#D95F5F"; $gpJob.BtnThRevert.Foreground = Get-WpfColor "#FFFFFF"
        $gpJob.BtnThRevert.ToolTip = $null
    } else {
        $gpJob.BtnThRevert.IsEnabled = $false; $gpJob.BtnThRevert.Background = Get-WpfColor "#3A3A3A"; $gpJob.BtnThRevert.Foreground = Get-WpfColor "#666666"
        $gpJob.BtnThRevert.ToolTip = "No merged files to revert"
    }
}

function Update-GlobalProcessAllStatus {
    $hasAnyIssue = $false
    $hasAnySelected = $false
    foreach ($gp in $script:jobs) {
        foreach ($p in $gp.Parents) {
            if ($p.IsQueued -or $p.IsDone) { continue }
            if (Get-PJobHasSelectedTask $p) { $hasAnySelected = $true }
            if ($p.HasCollision) { $hasAnyIssue = $true; break }
            foreach ($slot in $p.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $hasAnyIssue = $true; break } }
            if ($hasAnyIssue) { break }
        }
        if ($hasAnyIssue) { break }
    }
    if ($hasAnyIssue -or -not $hasAnySelected) {
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
        $pJob.BtnApply.Content = "Name Collision!"; $pJob.BtnApply.Background = Get-WpfColor "#D95F5F"; $pJob.BtnApply.Foreground = Get-WpfColor "#FFFFFF"; $pJob.BtnApply.IsEnabled = $false
    } elseif (-not $colorsSafe) {
        $pJob.BtnApply.Content = "Unmatched Colors"; $pJob.BtnApply.Background = Get-WpfColor "#E8A135"; $pJob.BtnApply.Foreground = Get-WpfColor "#FFFFFF"; $pJob.BtnApply.IsEnabled = $false
    } elseif (-not (Get-PJobHasSelectedTask $pJob)) {
        $pJob.BtnApply.Content = "Add to Queue"; $pJob.BtnApply.Background = Get-WpfColor "#3A3A3A"; $pJob.BtnApply.Foreground = Get-WpfColor "#666666"; $pJob.BtnApply.IsEnabled = $false
    } else {
        $pJob.BtnApply.Content = "Add to Queue"; $pJob.BtnApply.Background = Get-WpfColor "#4CAF72"; $pJob.BtnApply.Foreground = Get-WpfColor "#FFFFFF"; $pJob.BtnApply.IsEnabled = $true
    }
    # Show Rename Only bypass when colors are unmatched (not a collision) and rename is checked
    if ($null -ne $pJob.BtnRenameOnly) {
        $showBypass = (-not $colorsSafe) -and (-not $pJob.HasCollision) -and ([bool]$pJob.ChkRename.IsChecked)
        $pJob.BtnRenameOnly.Visibility = if ($showBypass) { "Visible" } else { "Collapsed" }
    }
    if ($null -ne $pJob._GpJob) { Update-ThProcessButtonState $pJob._GpJob }
    Update-GlobalProcessAllStatus
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
            $runBase = New-Object System.Windows.Documents.Run($basePart); $runBase.Foreground = Get-WpfColor "#D95F5F"
            $runSfx  = New-Object System.Windows.Documents.Run($sfxPart);  $runSfx.Foreground  = Get-WpfColor "#D95F5F"
            $r.NewLbl.Inlines.Add($runBase); $r.NewLbl.Inlines.Add((New-Object System.Windows.Documents.LineBreak)); $r.NewLbl.Inlines.Add($runSfx)
            if ($r.OldLbl) { $r.OldLbl.Foreground = Get-WpfColor "#D95F5F" }
            $hasCollision = $true
        } else {
            $runBase = New-Object System.Windows.Documents.Run($basePart); $runBase.Foreground = Get-WpfColor "#90B8C8"
            $runSfx  = New-Object System.Windows.Documents.Run($sfxPart);  $runSfx.Foreground  = Get-WpfColor $r.BaseColor
            $r.NewLbl.Inlines.Add($runBase); $r.NewLbl.Inlines.Add((New-Object System.Windows.Documents.LineBreak)); $r.NewLbl.Inlines.Add($runSfx)
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

    # Update grandparent folder name preview (Printer_Tag_Theme)
    if ($null -ne $gpJob.LblGpPreview) {
        $gpTg = "Standard"
        if ($null -ne $gpJob.CbTag -and $null -ne $gpJob.CbTag.SelectedItem -and "$($gpJob.CbTag.SelectedItem)" -ne "(none)") {
            $gpTg = "$($gpJob.CbTag.SelectedItem)"
        }
        $gpPreview = ((@($pf, $gpTg, $th) | Where-Object { $_ }) -join '_')
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
    if (-not $newAnchor) {
        Update-GpFileCount $gpJob
        Update-ThRevertButtonState $gpJob
        Update-ThProcessButtonState $gpJob
        Update-GlobalProcessAllStatus
        return
    }

    # Rebuild from scratch â€” re-parses slots, SmartFill, full card UI
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
    Update-ThRevertButtonState $gpJob
    Update-ThProcessButtonState $gpJob
    Update-GlobalProcessAllStatus
}

function Enqueue-PJob($pJob, $gpJob) {
    if ($pJob.IsQueued -or $pJob.IsDone -or $pJob.HasCollision) { return }
    if (-not $pJob.RenameOnlyBypass) {
        foreach ($slot in $pJob.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { return } }
    }

    # First plate queued for this grandparent: confirm rename if needed, then lock UI
    if (-not $gpJob.GpRenameConfirmed) {
        if (-not $gpJob.ChkSkip.IsChecked) {
            $th = ("$($gpJob.TBTheme.SelectedItem)" -replace '[^a-zA-Z0-9]', '')
            $pf = if ($null -ne $gpJob.CbPrefix -and "$($gpJob.CbPrefix.SelectedItem)" -ne "(none)") { "$($gpJob.CbPrefix.SelectedItem)" } else { "" }
            $tg = if ($null -ne $gpJob.CbTag -and "$($gpJob.CbTag.SelectedItem)" -ne "(none)") { "$($gpJob.CbTag.SelectedItem)" } else { "Standard" }
            $newGpFolderName = ((@($pf, $tg, $th) | Where-Object { $_ }) -join '_')
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

    $script:processQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob; SliceOnly = $false })
}

# Apply hex color substitutions directly to every non-gcode .3mf in the folder.
# $SubMap is a hashtable of OldHex7 -> @{ Old7; Old9; New7; New9 }
function Invoke-ColorPatchAllFiles($FolderPath, $SubMap, $SkipPath, $NameByHex) {
    if ($SubMap.Count -eq 0 -and ($null -eq $NameByHex -or $NameByHex.Count -eq 0)) { return }
    $files3mf = Get-ChildItem -Path $FolderPath -Filter "*.3mf" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notmatch '(?i)\.gcode\.3mf$' -and $_.FullName -ne $SkipPath }
    foreach ($f3 in $files3mf) {
        $tmpPath = [System.IO.Path]::GetTempFileName() + ".3mf"
        try {
            $srcStream = [System.IO.File]::OpenRead($f3.FullName)
            $srcZip    = [System.IO.Compression.ZipArchive]::new($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
            $anyChange = $false
            $entryData = @{}
            foreach ($entry in $srcZip.Entries) {
                if ($entry.Name -match '\.(xml|model|config|json)$') {
                    $sr      = [System.IO.StreamReader]::new($entry.Open())
                    $content = $sr.ReadToEnd(); $sr.Dispose()
                    $changed = $false
                    foreach ($sm in $SubMap.Values) {
                        if ($sm.Old9 -ne $sm.New9 -and $content -match "(?i)$([regex]::Escape($sm.Old9))") { $content = $content -ireplace [regex]::Escape($sm.Old9), $sm.New9; $changed = $true }
                        if ($sm.Old7 -ne $sm.New7 -and $content -match "(?i)$([regex]::Escape($sm.Old7))") { $content = $content -ireplace [regex]::Escape($sm.Old7), $sm.New7; $changed = $true }
                    }
                    if ($entry.Name -ieq 'project_settings.config' -and $null -ne $NameByHex -and $NameByHex.Count -gt 0) {
                        $tunedJson = Update-PurgeMatrixInConfigText $content $NameByHex
                        if ($null -ne $tunedJson) { $content = $tunedJson; $changed = $true }
                    }
                    if ($changed) { $anyChange = $true }
                    $entryData[$entry.FullName] = @{ Text = $content; Changed = $changed }
                }
            }
            if ($anyChange) {
                $dstStream = [System.IO.File]::Open($tmpPath, [System.IO.FileMode]::Create)
                $dstZip    = [System.IO.Compression.ZipArchive]::new($dstStream, [System.IO.Compression.ZipArchiveMode]::Create)
                foreach ($entry in $srcZip.Entries) {
                    $dstEntry       = $dstZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                    $dstEntryStream = $dstEntry.Open()
                    if ($entryData.ContainsKey($entry.FullName)) {
                        $bytes = [System.Text.Encoding]::UTF8.GetBytes($entryData[$entry.FullName].Text)
                        $dstEntryStream.Write($bytes, 0, $bytes.Length)
                    } else {
                        $srcEntryStream = $entry.Open(); $srcEntryStream.CopyTo($dstEntryStream); $srcEntryStream.Dispose()
                    }
                    $dstEntryStream.Dispose()
                }
                $dstZip.Dispose(); $dstStream.Dispose()
                $srcZip.Dispose(); $srcStream.Dispose()
                [System.IO.File]::Copy($tmpPath, $f3.FullName, $true)
            } else {
                $srcZip.Dispose(); $srcStream.Dispose()
            }
        } catch {}
        finally { if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force } }
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    }
}

# Build hex(7-char, upper) -> filament name map from the current color selections, for purge tuning lookups
function Get-PurgeNameByHexMap($UISlots) {
    $map = @{}
    foreach ($slot in $UISlots) {
        $selName = $slot.Combo.Text
        if ($script:LibraryColors.Contains($selName)) {
            $hex7 = $script:LibraryColors[$selName].ToUpper().Substring(0,7)
            $map[$hex7] = $selName
        }
    }
    return $map
}

# Look up the tuned purge volume for a (from, to) filament pair - returns $null if untuned/unset
function Get-PurgeTunedVolume([string]$FromName, [string]$ToName) {
    if (-not $FromName -or -not $ToName -or $FromName -eq $ToName -or $null -eq $script:PurgeDict) { return $null }
    foreach ($row in $script:PurgeDict) {
        if ($row.Source_Filament -eq $FromName -and $row.Target_Filament -eq $ToName) {
            $v = $row.Tuned_Volume
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
        }
    }
    return $null
}

# Rewrites flush_volumes_matrix entries in a project_settings.config JSON blob using tuned PurgeDictionary values.
# Returns the updated JSON text, or $null if nothing changed (or the JSON doesn't look like a Bambu project config).
function Update-PurgeMatrixInConfigText([string]$JsonText, [hashtable]$NameByHex) {
    try {
        $data = $JsonText | ConvertFrom-Json -ErrorAction Stop
    } catch { return $null }
    $propNames = @($data.PSObject.Properties.Name)
    if ($propNames -notcontains 'filament_colour' -or $propNames -notcontains 'flush_volumes_matrix') { return $null }

    $colors = @($data.filament_colour)
    $matrix = @($data.flush_volumes_matrix)
    $n = $colors.Count
    if ($n -eq 0 -or $matrix.Count -ne ($n * $n)) { return $null }

    $names = @()
    foreach ($hex in $colors) {
        $clean = ([string]$hex).ToUpper() -replace '^(#[0-9A-F]{6})[0-9A-F]{0,2}$', '$1'
        $clean7 = if ($clean.Length -ge 7) { $clean.Substring(0,7) } else { $clean }
        if ($NameByHex.ContainsKey($clean7)) { $names += $NameByHex[$clean7] }
        elseif ($script:HexToName.ContainsKey($clean7)) { $names += $script:HexToName[$clean7] }
        else { $names += $null }
    }

    $changed = $false
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $n; $j++) {
            if ($i -eq $j) { continue }
            $tunedVol = Get-PurgeTunedVolume $names[$i] $names[$j]
            if ($null -ne $tunedVol) {
                $idx = $i * $n + $j
                if ("$($matrix[$idx])" -ne $tunedVol) { $matrix[$idx] = $tunedVol; $changed = $true }
            }
        }
    }
    if (-not $changed) { return $null }

    $data.flush_volumes_matrix = $matrix
    return ($data | ConvertTo-Json -Depth 20 -Compress:$false)
}

# Color-only enqueue â€” saves current color selections to the .3mf files, no other tasks run
function Enqueue-ColorOnlyJob($pJob, $gpJob) {
    if ($pJob.IsQueued -or $pJob.IsDone) { return }
    $pJob.ColorOnlyBypass    = $true
    $gpJob.GpRenameConfirmed = $true
    $pJob.IsQueued = $true
    $pJob.CardStatusLabel.Text = "[ PREPARING ]"; $pJob.ProcessingOverlay.Visibility = "Visible"
    $pJob.PickStatusLabel.Text = "[ PREPARING ]"; $pJob.PickProcessingOverlay.Visibility = "Visible"
    $script:processQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob; SliceOnly = $false })
}

# Slice-only enqueue â€” bypasses color matching and rename-confirm checks (used by Editing mode)
function Enqueue-SliceOnlyJob($pJob, $gpJob) {
    if ($pJob.IsQueued -or $pJob.IsDone) { return }
    $pJob.SliceOnlyBypass    = $true
    $gpJob.GpRenameConfirmed = $true   # skip rename confirm dialog
    $pJob.IsQueued = $true
    # Show status in editing panel
    if ($null -ne $pJob.RenestStatusLbl) {
        $pJob.RenestStatusLbl.Text = "Export GCode queued..."; $pJob.RenestStatusLbl.Foreground = Get-WpfColor "#E8A135"
    }
    if ($null -ne $pJob.LblEdQueueStatus) {
        $pJob.LblEdQueueStatus.Text = "Export queued"; $pJob.LblEdQueueStatus.Foreground = Get-WpfColor "#E8A135"
    }
    # Show processing overlays on the card image area (visible even in Editing mode)
    $pJob.CardStatusLabel.Text = "[ PREPARING ]"; $pJob.ProcessingOverlay.Visibility = "Visible"
    $pJob.PickStatusLabel.Text = "[ PREPARING ]"; $pJob.PickProcessingOverlay.Visibility = "Visible"
    # Tag SliceOnly=true so the job-finished handler knows to skip KEEP/REVERT and leave card reusable
    $script:processQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob; SliceOnly = $true })
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
    $tg = "Standard"
    if ($null -ne $gpJob.CbTag -and $null -ne $gpJob.CbTag.SelectedItem) {
        $tg = $gpJob.CbTag.SelectedItem.ToString()
        if ($tg -eq "(none)") { $tg = "Standard" }
    }
    $oldGrand = if ($gpJob.DiGrand) { $gpJob.DiGrand.FullName } else { "" }

    # Uncheck merge/image if colors unresolved
    $colorsSafe = $true
    foreach ($slot in $pJob.UISlots) { if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $colorsSafe = $false } }
    if (-not $colorsSafe) { $pJob.ChkMerge.IsChecked = $false; $pJob.ChkImage.IsChecked = $false }

    # Apply color substitutions to extracted temp files
    $allTextFiles = Get-ChildItem -Path $pJob.TempWork -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
    $modifiedFiles = New-Object System.Collections.ArrayList
    $nameByHex = Get-PurgeNameByHexMap $pJob.UISlots

    foreach ($file in $allTextFiles) {
        $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
        $modified = $false
        foreach ($slot in $pJob.UISlots) {
            $selName = $slot.Combo.Text
            if ($script:LibraryColors.Contains($selName)) {
                $newHex = $script:LibraryColors[$selName].ToUpper(); $oldHex = $slot.OldHex.ToUpper()
                $oldHex9 = if ($oldHex.Length -eq 7) { $oldHex + "FF" } else { $oldHex }
                $oldHex7 = $oldHex.Substring(0,7)
                $newHex9 = if ($newHex.Length -eq 7) { $newHex + "FF" } else { $newHex }
                $newHex7 = $newHex.Substring(0,7)
                if ($newHex9 -ne $oldHex9 -and $content -match "(?i)$oldHex9") { $content = $content -ireplace [regex]::Escape($oldHex9), $newHex9; $modified = $true }
                if ($newHex7 -ne $oldHex7 -and $content -match "(?i)$oldHex7") { $content = $content -ireplace [regex]::Escape($oldHex7), $newHex7; $modified = $true }
            }
        }
        if ($file.Name -ieq 'project_settings.config') {
            $tunedJson = Update-PurgeMatrixInConfigText $content $nameByHex
            if ($null -ne $tunedJson) { $content = $tunedJson; $modified = $true }
        }
        if ($modified) {
            [System.IO.File]::WriteAllText($file.FullName, $content, (New-Object System.Text.UTF8Encoding($false)))
            $modifiedFiles.Add($file) | Out-Null
        }
    }

    # RenameOnlyBypass: user confirmed rename despite unmatched colors â€” skip all heavy tasks
    if ($pJob.ColorOnlyBypass) {
        $doRename = $false; $doMerge = $false; $doSlice = $false; $doExtract = $false; $doImage = $false; $doLogs = $false; $doBOD = $false; $doPrintQ = $false
        $pJob.ColorOnlyBypass = $false
        $pJob.ProcessedAnchorPath = $pJob.AnchorFile.FullName
        # Build sub map and patch all other .3mf files in the folder
        $colorSubMap = @{}
        foreach ($slot in $pJob.UISlots) {
            $selName = $slot.Combo.Text
            if ($script:LibraryColors.Contains($selName)) {
                $newHex = $script:LibraryColors[$selName].ToUpper(); $oldHex = $slot.OldHex.ToUpper()
                $old7 = $oldHex.Substring(0,7); $old9 = if ($oldHex.Length -eq 7) { $oldHex + "FF" } else { $oldHex }
                $new7 = $newHex.Substring(0,7); $new9 = if ($newHex.Length -eq 7) { $newHex + "FF" } else { $newHex }
                $colorSubMap[$old7] = @{ Old7 = $old7; Old9 = $old9; New7 = $new7; New9 = $new9 }
            }
        }
        Invoke-ColorPatchAllFiles $pJob.FolderPath $colorSubMap $pJob.AnchorFile.FullName $nameByHex
    } elseif ($pJob.RenameOnlyBypass) {
        $doRename = $true; $doMerge = $false; $doSlice = $false; $doExtract = $false; $doImage = $false; $doLogs = $false; $doBOD = $false; $doPrintQ = $false
        $pJob.RenameOnlyBypass = $false
    } elseif ($pJob.SliceOnlyBypass) {
        # Editing mode slice-only â€” no rename, no merge, just slice using the current anchor
        $doRename = $false; $doMerge = $false; $doSlice = $true; $doExtract = $false; $doImage = $false; $doLogs = $false; $doBOD = $false; $doPrintQ = $false
        $pJob.SliceOnlyBypass = $false
        $pJob.ProcessedAnchorPath = $pJob.AnchorFile.FullName
    } elseif ([bool]$pJob.ChkColors.IsChecked -and -not ([bool]$pJob.ChkRename.IsChecked -or [bool]$pJob.ChkMerge.IsChecked -or [bool]$pJob.ChkSlice.IsChecked -or [bool]$pJob.ChkExtract.IsChecked -or [bool]$pJob.ChkImage.IsChecked -or [bool]$pJob.ChkLogs.IsChecked -or [bool]$pJob.ChkBOD.IsChecked -or [bool]$pJob.ChkPrintQ.IsChecked)) {
        # Save Colors only â€” no heavy tasks
        $doRename = $false; $doMerge = $false; $doSlice = $false; $doExtract = $false; $doImage = $false; $doLogs = $false; $doBOD = $false; $doPrintQ = $false
        $pJob.ProcessedAnchorPath = $pJob.AnchorFile.FullName
        # Build sub map and patch all other .3mf files in the folder
        $colorSubMap = @{}
        foreach ($slot in $pJob.UISlots) {
            $selName = $slot.Combo.Text
            if ($script:LibraryColors.Contains($selName)) {
                $newHex = $script:LibraryColors[$selName].ToUpper(); $oldHex = $slot.OldHex.ToUpper()
                $old7 = $oldHex.Substring(0,7); $old9 = if ($oldHex.Length -eq 7) { $oldHex + "FF" } else { $oldHex }
                $new7 = $newHex.Substring(0,7); $new9 = if ($newHex.Length -eq 7) { $newHex + "FF" } else { $newHex }
                $colorSubMap[$old7] = @{ Old7 = $old7; Old9 = $old9; New7 = $new7; New9 = $new9 }
            }
        }
        Invoke-ColorPatchAllFiles $pJob.FolderPath $colorSubMap $pJob.AnchorFile.FullName $nameByHex
    } else {
        $doRename  = [bool]$pJob.ChkRename.IsChecked
        $doMerge   = [bool]$pJob.ChkMerge.IsChecked
        $doSlice   = [bool]$pJob.ChkSlice.IsChecked
        $doExtract = [bool]$pJob.ChkExtract.IsChecked
        $doImage   = [bool]$pJob.ChkImage.IsChecked
        $doLogs    = [bool]$pJob.ChkLogs.IsChecked
        $doBOD     = [bool]$pJob.ChkBOD.IsChecked
        $doPrintQ  = [bool]$pJob.ChkPrintQ.IsChecked
    }

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
        $newGpFolderName = ((@($pf, $tg, $th) | Where-Object { $_ }) -join '_')
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
    if (-not ($doMerge -or $doSlice -or $doExtract -or $doImage -or $doBOD -or $doPrintQ)) {
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
    $basePrefix = if ($baseName.ToLower().EndsWith("full")) { $baseName.Substring(0, $baseName.Length - 4) }
                  elseif ($baseName -match '(?i)_Final$') { $baseName -replace '(?i)_Final$', '_' }
                  else { $baseName + "_" }

    $anchorPath    = $pJob.ProcessedAnchorPath
    $nestPath      = Join-Path $dir "$($basePrefix)Nest.3mf"
    $finalPath     = Join-Path $dir "$($basePrefix)Final.3mf"
    $tempOut       = Join-Path $dir "$($baseName)_merged_temp.3mf"
    $tempIso       = Join-Path $env:TEMP "iso_$([guid]::NewGuid().ToString().Substring(0,8))"
    $slicedFile    = Join-Path $dir "$($baseName).gcode.3mf"
    $singleFile    = Join-Path $dir "$($basePrefix)Final.gcode.3mf"
    $tsvBaseName   = $baseName -replace '(?i)_Full$', ''
    $tsvFile       = Join-Path $dir "${tsvBaseName}_Data.tsv"
    $bodTempPath   = Join-Path $dir "$($basePrefix)BOD.3mf"
    $bodGcodeTemp  = Join-Path $dir "$($basePrefix)BOD.gcode.3mf"
    $bodQueueBase  = "C:\Users\Owner\SynologyDrive\WIGGLITEERZ\THEKITCHEN\Printing Queue"

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

    if ($doBOD) {
        # Full.3mf path: after a merge it is $baseName.3mf (only meaningful when the anchor is a merged "...Full" design --
        # for Final-only designs $baseName already ends in "_Final" and this path just coincides with $bodFinalPath)
        $bodNestPath  = Join-Path $dir "$($basePrefix)Nest.3mf"
        $bodFullPath  = Join-Path $dir "$baseName.3mf"
        $bodFinalPath = Join-Path $dir "$($basePrefix)Final.3mf"
        $bodIsFullDesign = $baseName.ToLower().EndsWith("full")
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'CREATING BOD...' -Force")
        [void]$sb.AppendLine("`$bodNestPath     = `"$bodNestPath`"")
        [void]$sb.AppendLine("`$bodFullPath     = `"$bodFullPath`"")
        [void]$sb.AppendLine("`$bodFinalPath    = `"$bodFinalPath`"")
        [void]$sb.AppendLine("`$bodIsFullDesign = `$$bodIsFullDesign")
        [void]$sb.AppendLine("`$bodHasNest  = Test-Path `$bodNestPath")
        [void]$sb.AppendLine("`$bodHasFull  = `$bodIsFullDesign -and (Test-Path `$bodFullPath)")
        [void]$sb.AppendLine("`$bodHasFinal = Test-Path `$bodFinalPath")
        [void]$sb.AppendLine("if (-not `$bodHasFull -and -not `$bodHasFinal) { Write-Host '[BOD] Neither Full.3mf nor Final.3mf found - skipping BOD.' -ForegroundColor Yellow }")
        [void]$sb.AppendLine("else {")
        # Mode/count selection:
        #   Nest + Full        -> merged design; keep the 7 pairs closest to plate centre
        #   Full + Final (no Nest) -> Full.3mf was never merged (no Nest = no merge ever happened),
        #                             so its build items are lone objects; keep the 15 closest to centre
        #   Final only         -> single-object design; build a 15-copy grid (Grid mode)
        [void]$sb.AppendLine("    `$bodPairCount = 5")
        [void]$sb.AppendLine("    if (`$bodHasNest -and `$bodHasFull) {")
        [void]$sb.AppendLine("        `$bodMode = 'Full'; `$bodPairCount = 7; `$bodSource = `$bodFullPath")
        [void]$sb.AppendLine("        Write-Host '[BOD] Nest + Full found - merged design, keeping 7 pairs closest to plate centre.' -ForegroundColor Cyan")
        [void]$sb.AppendLine("    } elseif (`$bodHasFull -and `$bodHasFinal) {")
        [void]$sb.AppendLine("        `$bodMode = 'Full'; `$bodPairCount = 15; `$bodSource = `$bodFullPath")
        [void]$sb.AppendLine("        Write-Host '[BOD] Full + Final found, no Nest - unmerged design, keeping 15 lone objects closest to plate centre.' -ForegroundColor Cyan")
        [void]$sb.AppendLine("    } else {")
        [void]$sb.AppendLine("        `$bodMode = 'Grid'; `$bodSource = `$bodFinalPath")
        [void]$sb.AppendLine("        Write-Host '[BOD] Final only - single-object design, building a 15-copy grid.' -ForegroundColor Cyan")
        [void]$sb.AppendLine("    }")
        [void]$sb.AppendLine("    Write-Host `"[BOD] Mode: `$bodMode  Source: `$(Split-Path `$bodSource -Leaf)`" -ForegroundColor Cyan")
        [void]$sb.AppendLine("    `$bodDate = Get-Date -Format 'MMMM d'")
        [void]$sb.AppendLine("    `$bodFolder = Join-Path `"$bodQueueBase`" `$bodDate")
        [void]$sb.AppendLine("    New-Item -ItemType Directory -Path `$bodFolder -Force | Out-Null")
        [void]$sb.AppendLine("    & `"$scriptDir\create_bod_worker.ps1`" -InputPath `$bodSource -OutputPath `"$bodTempPath`" -Mode `$bodMode -PairCount `$bodPairCount")
        [void]$sb.AppendLine("    if (Test-Path `"$bodTempPath`") {")
        [void]$sb.AppendLine("        Set-Content -Path `"$statusFile`" -Value 'SLICING BOD... 0%' -Force")
        [void]$sb.AppendLine("        & `"$scriptDir\Slice_worker.ps1`" -InputPath `"$bodTempPath`" -StatusFile `"$statusFile`"")
        [void]$sb.AppendLine("        if (Test-Path `"$bodGcodeTemp`") {")
        [void]$sb.AppendLine("            `$bodDest = Join-Path `$bodFolder `"$($basePrefix)BOD.gcode.3mf`"")
        [void]$sb.AppendLine("            Move-Item `"$bodGcodeTemp`" `$bodDest -Force")
        [void]$sb.AppendLine("            Write-Host `"[BOD] Exported to: `$bodDest`" -ForegroundColor Green")
        [void]$sb.AppendLine("            Remove-Item `"$bodTempPath`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("        } else { Write-Host '[BOD] Slice produced no gcode output - keeping BOD.3mf for inspection.' -ForegroundColor Yellow }")
        [void]$sb.AppendLine("    } else { Write-Host '[BOD] create_bod_worker produced no output.' -ForegroundColor Yellow }")
        [void]$sb.AppendLine("}")
    }

    if ($doSlice) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'SLICING... 0%' -Force")
        [void]$sb.AppendLine("Start-Sleep -Seconds 3")
        if ($jobWrapper.SliceOnly) {
            # Editing-mode slice: no isolated Final.3mf gcode needed
            [void]$sb.AppendLine("& `"$scriptDir\Slice_worker.ps1`" -InputPath `"$anchorPath`" -StatusFile `"$statusFile`"")
        } else {
            [void]$sb.AppendLine("& `"$scriptDir\Slice_worker.ps1`" -InputPath `"$anchorPath`" -IsolatedPath `"$finalPath`" -StatusFile `"$statusFile`"")
        }
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

    if ($doPrintQ) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value 'COPYING TO PRINT QUEUE...' -Force")
        [void]$sb.AppendLine("if (Test-Path `"$slicedFile`") {")
        [void]$sb.AppendLine("    `$pqDate   = Get-Date -Format 'MMMM d'")
        [void]$sb.AppendLine("    `$pqFolder = Join-Path `"$bodQueueBase`" `$pqDate")
        [void]$sb.AppendLine("    New-Item -ItemType Directory -Path `$pqFolder -Force | Out-Null")
        [void]$sb.AppendLine("    `$pqDest = Join-Path `$pqFolder `"$($baseName).gcode.3mf`"")
        [void]$sb.AppendLine("    Copy-Item -LiteralPath `"$slicedFile`" -Destination `$pqDest -Force")
        [void]$sb.AppendLine("    Write-Host `"[PrintQ] Copied to: `$pqDest`" -ForegroundColor Green")
        [void]$sb.AppendLine("} else { Write-Host '[PrintQ] Gcode not found - run Slice first.' -ForegroundColor Yellow }")
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

# Count objects with type="model" inside a .3mf ZIP's 3dmodel.model
function Count-3mfObjects([string]$path) {
    if (-not (Test-Path $path)) { return 0 }
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        $modelEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -imatch "(?i)3D/3dmodel\.model$" } | Select-Object -First 1
        if ($null -eq $modelEntry) { $zip.Dispose(); return 0 }
        $sr = New-Object System.IO.StreamReader($modelEntry.Open())
        $content = $sr.ReadToEnd(); $sr.Close(); $zip.Dispose()
        return ([regex]::Matches($content, 'type="model"')).Count
    } catch { if ($null -ne $zip) { try { $zip.Dispose() } catch {} }; return 0 }
}

# Read a single named entry's raw bytes out of a .3mf (zip) file, or $null if missing
function Get-ZipEntryBytes([string]$zipPath, [string]$entryName) {
    if (-not (Test-Path -LiteralPath $zipPath)) { return $null }
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        $entry = $zip.Entries | Where-Object { ($_.FullName -replace '\\','/') -eq $entryName } | Select-Object -First 1
        if ($null -eq $entry) { $zip.Dispose(); return $null }
        $ms = New-Object System.IO.MemoryStream
        $entry.Open().CopyTo($ms)
        $zip.Dispose()
        return $ms.ToArray()
    } catch { return $null }
}

# True if the named zip entry exists in both files and its bytes differ
function Compare-ZipEntryImagesDiffer([string]$pathA, [string]$pathB, [string]$entryName) {
    $a = Get-ZipEntryBytes $pathA $entryName
    $b = Get-ZipEntryBytes $pathB $entryName
    if ($null -eq $a -or $null -eq $b) { return $false }
    if ($a.Length -ne $b.Length) { return $true }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hashA = [System.Convert]::ToBase64String($sha.ComputeHash($a))
    $hashB = [System.Convert]::ToBase64String($sha.ComputeHash($b))
    return ($hashA -ne $hashB)
}

# Copy a single design's Full.gcode.3mf to today's Printing Queue folder.
# Used by both the per-card "Printing Queue" button and the theme-wide "Printing Queue (Theme)" button.
function Send-PJobToPrintingQueue($pj) {
    $gcodeFile = Get-ChildItem -Path $pj.FolderPath -Filter "*Full.gcode.3mf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $gcodeFile) {
        $pj.LblReviewPrintQStatus.Text = "No Full.gcode.3mf found - run Slice first."
        $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#D95F5F"
        return
    }
    $pqBase   = "C:\Users\Owner\SynologyDrive\WIGGLITEERZ\THEKITCHEN\Printing Queue"
    $pqDate   = Get-Date -Format 'MMMM d'
    $pqFolder = Join-Path $pqBase $pqDate
    New-Item -ItemType Directory -Path $pqFolder -Force | Out-Null
    $pqDest = Join-Path $pqFolder $gcodeFile.Name
    Copy-Item -LiteralPath $gcodeFile.FullName -Destination $pqDest -Force
    $pj.LblReviewPrintQStatus.Text = "Copied to Printing Queue: $pqDest"
    $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#4CAF72"
}

# Run Deploy_Gcode.bat for a single design and report a friendly summary into $pj.LblReviewPrintQStatus.
# Used by both the per-card "Send to Production" button and the theme-wide "Send to Production (Theme)" button.
# $isFix : $true = "Fixed Designs" deploy ($--fix), $false = "New Designs" deploy.
function Send-PJobToProduction($pj, [bool]$isFix) {
    try {
        $deployScript = "C:\Users\Owner\SynologyDrive\WIGGLITEERZ\04.5 Chippy Queue\Deploy_Gcode.bat"
        if (-not (Test-Path -LiteralPath $deployScript)) {
            $pj.LblReviewPrintQStatus.Text = "Deploy script not found: $deployScript"
            $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#D95F5F"
            return
        }

        # --- Pre-check: warn if files of the same name already exist at the destination ---
        $deployDir    = Split-Path $deployScript -Parent
        $parentFolder = Split-Path $pj.FolderPath -Leaf
        $printerPrefix = ($parentFolder -split '_')[0]
        $printerBase = $null
        if ($printerPrefix -ieq "X1C") { $printerBase = "X1C" }
        elseif ($printerPrefix -ieq "P2S") { $printerBase = "P2S" }

        if ($null -eq $printerBase) {
            $pj.LblReviewPrintQStatus.Text = "Unrecognized printer prefix '$printerPrefix' in folder '$parentFolder' - Deploy_Gcode.bat will skip this folder."
            $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#D95F5F"
            return
        }

        $kindFolder = if ($isFix) { "Fixed Files" } else { "New Files" }
        $targetDir  = Join-Path (Join-Path $deployDir "$printerBase $kindFolder") $parentFolder

        $gcodeFiles = Get-ChildItem -Path $pj.FolderPath -Filter "*Full.gcode.3mf" -File -Recurse -ErrorAction SilentlyContinue
        $existingNames = New-Object System.Collections.Generic.List[string]
        if (Test-Path -LiteralPath $targetDir) {
            foreach ($gf in $gcodeFiles) {
                $names = @($gf.Name, ($gf.Name -replace 'Full\.gcode\.3mf$','Nest.3mf'), ($gf.Name -replace 'Full\.gcode\.3mf$','Data.tsv'))
                foreach ($n in $names) {
                    if (Test-Path -LiteralPath (Join-Path $targetDir $n)) { $existingNames.Add($n) }
                }
            }
        }

        if ($existingNames.Count -gt 0) {
            $msg = "The following file(s) already exist in:`n$targetDir`n`n" + (($existingNames | ForEach-Object { "  - $_" }) -join "`n") + "`n`nThey will be overwritten. Continue?"
            $overwrite = [System.Windows.MessageBox]::Show($msg, "Send to Production - Files Exist ($parentFolder)",
                [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($overwrite -ne [System.Windows.MessageBoxResult]::Yes) {
                $pj.LblReviewPrintQStatus.Text = "Send to Production cancelled - existing files in $targetDir were not overwritten."
                $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#E8A135"
                return
            }
        }

        $deployArgs = @("`"$($pj.FolderPath)`"", "--fork")
        if ($isFix) { $deployArgs += "--fix" }

        $pj.LblReviewPrintQStatus.Text = "Sending to production..."
        $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#888A9A"

        $outFile = [System.IO.Path]::GetTempFileName()
        $errFile = [System.IO.Path]::GetTempFileName()
        try {
            Start-Process -FilePath $deployScript -ArgumentList $deployArgs -NoNewWindow -Wait `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $output = (Get-Content -Path $outFile -Raw -ErrorAction SilentlyContinue)
            $errOutput = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue)
        } finally {
            Remove-Item -Path $outFile, $errFile -ErrorAction SilentlyContinue
        }

        # --- Build a friendly step-by-step summary instead of dumping raw bat output ---
        $summary = New-Object System.Collections.Generic.List[string]
        $allOk = $true

        if ($output -match "Unrecognized prefix") {
            $summary.Add("[SKIPPED] Unrecognized printer prefix - nothing was deployed.")
            $allOk = $false
        } elseif ($gcodeFiles.Count -eq 0 -or $output -notmatch "Deployment Complete") {
            $summary.Add("[FAIL] No Full.gcode.3mf package was deployed.")
            $allOk = $false
        } else {
            foreach ($gf in $gcodeFiles) {
                $nestName = $gf.Name -replace 'Full\.gcode\.3mf$','Nest.3mf'
                $dataName = $gf.Name -replace 'Full\.gcode\.3mf$','Data.tsv'
                $summary.Add("Package: $($gf.Name)")

                $gcodeOk = Test-Path -LiteralPath (Join-Path $targetDir $gf.Name)
                $summary.Add("  [$(if ($gcodeOk) {'OK'} else {'FAIL'})] $($gf.Name) copied")
                if (-not $gcodeOk) { $allOk = $false }

                if (Test-Path -LiteralPath (Join-Path $gf.DirectoryName $nestName)) {
                    $nestOk = Test-Path -LiteralPath (Join-Path $targetDir $nestName)
                    $summary.Add("  [$(if ($nestOk) {'OK'} else {'FAIL'})] $nestName copied")
                    if (-not $nestOk) { $allOk = $false }
                } else {
                    $summary.Add("  [MISSING] $nestName not found in design folder - not copied")
                    $allOk = $false
                }

                if (Test-Path -LiteralPath (Join-Path $gf.DirectoryName $dataName)) {
                    $dataOk = Test-Path -LiteralPath (Join-Path $targetDir $dataName)
                    $summary.Add("  [$(if ($dataOk) {'OK'} else {'FAIL'})] $dataName copied")
                    if (-not $dataOk) { $allOk = $false }

                    if ($output -match "SHEETS ERROR:(.+)") {
                        $summary.Add("  [FAIL] Google Sheet row - $($Matches[1].Trim())")
                        $allOk = $false
                    } elseif ($output -match "Uploading .*to Google Sheets") {
                        $summary.Add("  [OK] Google Sheet row added")
                    } else {
                        $summary.Add("  [FAIL] Google Sheet row not attempted")
                        $allOk = $false
                    }
                } else {
                    $summary.Add("  [MISSING] $dataName not found in design folder - not copied, Google Sheet not updated")
                    $allOk = $false
                }
            }
            $summary.Add("")
            $summary.Add("Target: $targetDir")
        }

        if (-not [string]::IsNullOrWhiteSpace($errOutput)) {
            $summary.Add("")
            $summary.Add("Errors: $errOutput")
            $allOk = $false
        }

        $summaryText = $summary -join "`n"
        if ($allOk) {
            $pj.LblReviewPrintQStatus.Text = "Sent to production:`n$summaryText"
            $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#4CAF72"
        } else {
            $pj.LblReviewPrintQStatus.Text = "Send to Production completed with issues:`n$summaryText"
            $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#D95F5F"
        }
    } catch {
        $pj.LblReviewPrintQStatus.Text = "Send to Production error: $($_.Exception.Message)"
        $pj.LblReviewPrintQStatus.Foreground = Get-WpfColor "#D95F5F"
        Write-Log "Send to Production error ($($pj.FolderPath)): $($_.Exception.Message)" "ERROR"
    }
}

# Build the read-only review content into $pJob.ReviewStack (called lazily on first toggle)
function Build-ReviewContent($pJob) {
    $sp = $pJob.ReviewStack

    # --- Print time and pre-merge object count from *_Data.tsv ---
    # TSV layout: Printer(0), FileType(1), FileName(2), SKU(3), Theme(4), Date(5), H(6), M(7), ... ObjCount(25)
    $printTimeStr = "n/a"
    $printHours = 0.0
    $objCount = 0
    $tsvFile = Get-ChildItem -Path $pJob.FolderPath -Filter "*_Data.tsv" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($tsvFile) {
        try {
            $tsvLine = Get-Content $tsvFile.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ($tsvLine) {
                $cols = $tsvLine -split "`t"
                if ($cols.Count -ge 8) {
                    $th = 0; $tm = 0
                    [int]::TryParse($cols[6], [ref]$th) | Out-Null
                    [int]::TryParse($cols[7], [ref]$tm) | Out-Null
                    $printTimeStr = if ($th -gt 0) { "${th}h ${tm}m" } else { "${tm}m" }
                    $printHours = $th + ($tm / 60.0)
                }
                if ($cols.Count -ge 26) {
                    [int]::TryParse($cols[25], [ref]$objCount) | Out-Null
                }
            }
        } catch {}
    }

    # --- Locate the base files this design depends on ---
    $allFiles  = Get-ChildItem -Path $pJob.FolderPath -File -ErrorAction SilentlyContinue
    $finalFile = $allFiles | Where-Object { $_.Name -imatch '(?i)[._ ]Final\.3mf$' }       | Select-Object -First 1
    $nestFile  = $allFiles | Where-Object { $_.Name -imatch '(?i)[._ ]Nest\.3mf$' }        | Select-Object -First 1
    $fullFile  = $allFiles | Where-Object { $_.Name -imatch 'Full\.3mf$' -and $_.Name -notmatch '(?i)gcode' } | Select-Object -First 1
    $gcodeFile = $allFiles | Where-Object { $_.Name -imatch '(?i)Full\.gcode\.3mf$' }      | Select-Object -First 1
    $pngFile   = $allFiles | Where-Object { $_.Name -imatch '\.png$' }                    | Select-Object -First 1

    # --- Pre / Post merge object counts ---
    $preCount = "n/a"; $postCount = "n/a"
    if ($nestFile) { $preCount  = (Count-3mfObjects $nestFile.FullName).ToString() }
    if ($fullFile) { $postCount = (Count-3mfObjects $fullFile.FullName).ToString() }

    # --- Production-readiness checklist ---
    $checkSku    = -not [string]::IsNullOrWhiteSpace($pJob.TxtSKU.Text)
    $checkMerged = ($null -ne $nestFile)
    $checkImage  = ($null -ne $fullFile) -and ($null -ne $gcodeFile) -and
                   (Compare-ZipEntryImagesDiffer $fullFile.FullName $gcodeFile.FullName "Metadata/plate_1.png")
    $checkFiles  = ($null -ne $finalFile) -and ($null -ne $nestFile) -and ($null -ne $fullFile) -and
                   ($null -ne $gcodeFile) -and ($null -ne $pngFile) -and ($null -ne $tsvFile)
    $isProdReady = $checkSku -and $checkMerged -and $checkImage -and $checkFiles
    $pJob.IsProdReady = $isProdReady

    if ($isProdReady) {
        $pJob.RowPanel.BorderBrush     = Get-WpfColor "#4CAF72"
        $pJob.RowPanel.BorderThickness = New-Object System.Windows.Thickness(8)
        $hdrReady = Create-TextBlock "PRODUCTION READY" "#4CAF72" 13 "Bold"
        $hdrReady.Margin = New-Object System.Windows.Thickness(0,0,0,8)
        $sp.Children.Add($hdrReady) | Out-Null
    } else {
        $hdrNotReady = Create-TextBlock "NOT READY" "#D95F5F" 13 "Bold"
        $hdrNotReady.Margin = New-Object System.Windows.Thickness(0,0,0,8)
        $sp.Children.Add($hdrNotReady) | Out-Null
    }

    # --- Checklist section ---
    $hdrChecklist = Create-TextBlock "Production Checklist" "#A0A0A0" 11 "Bold"
    $hdrChecklist.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $sp.Children.Add($hdrChecklist) | Out-Null
    foreach ($item in @(
        @{ Label = "SKU exists";           Ok = $checkSku }
        @{ Label = "File merged";          Ok = $checkMerged }
        @{ Label = "Gcode image replaced"; Ok = $checkImage }
        @{ Label = "Base files present";   Ok = $checkFiles }
    )) {
        $row = New-Object System.Windows.Controls.StackPanel; $row.Orientation = "Horizontal"
        $row.Margin = New-Object System.Windows.Thickness(0,1,0,1)
        $mark  = if ($item.Ok) { "[OK]  " } else { "[ ]  " }
        $color = if ($item.Ok) { "#4CAF72" } else { "#D95F5F" }
        $row.Children.Add((Create-TextBlock $mark $color 11 "Bold")) | Out-Null
        $row.Children.Add((Create-TextBlock $item.Label "#CCCCCC" 11 "Normal")) | Out-Null
        $sp.Children.Add($row) | Out-Null
    }

    # --- Print time ---
    $sep1 = New-Object System.Windows.Controls.Separator
    $sep1.Margin = New-Object System.Windows.Thickness(0,6,0,6)
    $sep1.Background = Get-WpfColor "#2A2C35"; $sp.Children.Add($sep1) | Out-Null
    $ptRow = New-Object System.Windows.Controls.StackPanel; $ptRow.Orientation = "Horizontal"
    $ptRow.Children.Add((Create-TextBlock "Print Time:  " "#A0A0A0" 12 "Normal")) | Out-Null
    $ptRow.Children.Add((Create-TextBlock $printTimeStr "#FFFFFF" 12 "Bold")) | Out-Null
    $sp.Children.Add($ptRow) | Out-Null

    # --- Merge counts ---
    $sep3 = New-Object System.Windows.Controls.Separator
    $sep3.Margin = New-Object System.Windows.Thickness(0,6,0,6)
    $sep3.Background = Get-WpfColor "#2A2C35"; $sp.Children.Add($sep3) | Out-Null
    $mRow = New-Object System.Windows.Controls.StackPanel; $mRow.Orientation = "Horizontal"
    $mRow.Children.Add((Create-TextBlock "Pre-merge:  "  "#A0A0A0" 11 "Normal")) | Out-Null
    $mRow.Children.Add((Create-TextBlock "$preCount   "  "#FFD700" 11 "Bold"))   | Out-Null
    $mRow.Children.Add((Create-TextBlock "Post-merge:  " "#A0A0A0" 11 "Normal")) | Out-Null
    $mRow.Children.Add((Create-TextBlock $postCount      "#4CAF72" 11 "Bold"))   | Out-Null
    $sp.Children.Add($mRow) | Out-Null

    # --- Wigglitz per day = (pre-merge object count / print hours) * 24 ---
    # Also rolled up into the theme-level average shown at the top of the group header.
    if ($objCount -gt 0 -and $printHours -gt 0) {
        $pJob.WigglitzValue = ($objCount / $printHours) * 24
    }
    $wRow = New-Object System.Windows.Controls.StackPanel; $wRow.Orientation = "Horizontal"
    $wRow.Children.Add((Create-TextBlock "Wigglitz/Day:  " "#A0A0A0" 12 "Normal")) | Out-Null
    $wVal = if ($null -ne $pJob.WigglitzValue) { "{0:N2}" -f $pJob.WigglitzValue } else { "n/a" }
    $wRow.Children.Add((Create-TextBlock $wVal "#FFD700" 12 "Bold")) | Out-Null
    $sp.Children.Add($wRow) | Out-Null

    # --- Action buttons: Printing Queue / Send to Production ---
    # Bigger, anchored in the footer row that spans the full width at the
    # bottom of the card (Review mode only).
    $footerStack = New-Object System.Windows.Controls.StackPanel
    $pJob.ReviewFooter.Child = $footerStack

    $reviewActionRow = New-Object System.Windows.Controls.StackPanel
    $reviewActionRow.Orientation = "Horizontal"
    $reviewActionRow.HorizontalAlignment = "Center"

    $btnReviewPrintQ = New-Object System.Windows.Controls.Button
    $btnReviewPrintQ.Content = "Printing Queue"; $btnReviewPrintQ.Height = 48; $btnReviewPrintQ.Width = 200
    $btnReviewPrintQ.FontWeight = [System.Windows.FontWeights]::Bold; $btnReviewPrintQ.FontSize = 16
    $btnReviewPrintQ.Background = Get-WpfColor "#3A5080"; $btnReviewPrintQ.Foreground = Get-WpfColor "#FFFFFF"
    $btnReviewPrintQ.BorderThickness = 0; $btnReviewPrintQ.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnReviewPrintQ.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $btnReviewPrintQ.ToolTip = "Copies the existing Full.gcode.3mf to the Printing Queue folder with today's date"
    $reviewActionRow.Children.Add($btnReviewPrintQ) | Out-Null

    $btnSendToProduction = New-Object System.Windows.Controls.Button
    $btnSendToProduction.Content = "Send to Production"; $btnSendToProduction.Height = 48; $btnSendToProduction.Width = 220
    $btnSendToProduction.FontWeight = [System.Windows.FontWeights]::Bold; $btnSendToProduction.FontSize = 16
    $btnSendToProduction.Background = Get-WpfColor "#4CAF72"; $btnSendToProduction.Foreground = Get-WpfColor "#FFFFFF"
    $btnSendToProduction.BorderThickness = 0; $btnSendToProduction.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnSendToProduction.ToolTip = "Runs Deploy_Gcode.bat to copy the gcode/Nest/Data files into the production libraries and upload the Data.tsv row to Google Sheets"
    $reviewActionRow.Children.Add($btnSendToProduction) | Out-Null

    $footerStack.Children.Add($reviewActionRow) | Out-Null

    $lblReviewPrintQStatus = New-Object System.Windows.Controls.TextBlock
    $lblReviewPrintQStatus.FontSize = 12; $lblReviewPrintQStatus.TextWrapping = "Wrap"
    $lblReviewPrintQStatus.TextAlignment = "Center"
    $lblReviewPrintQStatus.Margin = New-Object System.Windows.Thickness(0,8,0,0)
    $lblReviewPrintQStatus.Foreground = Get-WpfColor "#888A9A"
    $footerStack.Children.Add($lblReviewPrintQStatus) | Out-Null
    $pJob.LblReviewPrintQStatus = $lblReviewPrintQStatus

    # Lock "Send to Production" until this design passes the Production Checklist
    $btnSendToProduction.IsEnabled = $isProdReady
    if (-not $isProdReady) {
        $btnSendToProduction.Background = Get-WpfColor "#3A3D4A"
        $btnSendToProduction.ToolTip = "Locked until this design passes the Production Checklist above"
    }

    $btnReviewPrintQ.Tag = @{ P = $pJob }
    $btnReviewPrintQ.Add_Click({
        $t = $this.Tag
        Send-PJobToPrintingQueue $t.P
    })

    $btnSendToProduction.Tag = @{ P = $pJob }
    $btnSendToProduction.Add_Click({
        $t = $this.Tag; $pj = $t.P
        $choice = [System.Windows.MessageBox]::Show(
            "Is this a NEW design (not previously sent to production)?`n`nYes = New Designs`nNo = Fixed Designs",
            "Send to Production",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question)
        if ($choice -eq [System.Windows.MessageBoxResult]::Cancel) { return }
        Send-PJobToProduction $pj ($choice -eq [System.Windows.MessageBoxResult]::No)
    })
}

# Put a single parent into Review mode
function Set-PJobReviewMode($pJob) {
    if (-not $pJob.ReviewBuilt) {
        Build-ReviewContent $pJob
        $pJob.ReviewBuilt = $true
    }
    # Re-apply the production-ready border each time we enter Review mode -
    # Set-PJobEditMode resets it on the way out, and Build-ReviewContent only
    # runs once (gated by ReviewBuilt), so it won't set it again on its own.
    if ($pJob.IsProdReady) {
        $pJob.RowPanel.BorderBrush     = Get-WpfColor "#4CAF72"
        $pJob.RowPanel.BorderThickness = New-Object System.Windows.Thickness(8)
    }
    # Review shows only the gcode plate_1.png (via the full-card overlay image) -
    # hide the color-edit overlay and [CURRENT] thumbnail, but keep the
    # randomized Pick.png panel visible. Also hide the Re-Nest editor's
    # preview/axis-picker images.
    if ($pJob.GcodeImgPath -and (Test-Path $pJob.GcodeImgPath)) {
        $pJob.ReviewCardOverlay.Source     = Load-WpfImage $pJob.GcodeImgPath
        $pJob.ReviewCardOverlay.Visibility = "Visible"
    } else {
        $pJob.ReviewCardOverlay.Visibility = "Collapsed"
    }
    if ($null -ne $pJob.RnOvLeft)  { $pJob.RnOvLeft.Visibility  = "Collapsed" }
    if ($null -ne $pJob.RnOvRight) { $pJob.RnOvRight.Visibility = "Collapsed" }
    if ($null -ne $pJob.ColorsOverlayStack) { $pJob.ColorsOverlayStack.Visibility = "Collapsed" }
    if ($null -ne $pJob.CurrentThumb) {
        $pJob._SavedCurrentThumbVis = $pJob.CurrentThumb.Visibility
        $pJob.CurrentThumb.Visibility = "Collapsed"
    }
    $pJob.ReviewPanel.Visibility       = "Visible"
    if ($null -ne $pJob.ReviewFooter) { $pJob.ReviewFooter.Visibility = "Visible" }
    $pJob.TasksBox.Visibility          = "Collapsed"
    $pJob.EditBox.Visibility           = "Collapsed"
    $pJob.PnlFiles.Visibility          = "Collapsed"
    $pJob.ApplyRow.Visibility          = "Collapsed"
    $pJob.BtnRefresh.Visibility        = "Collapsed"
    $pJob.BtnRemoveP.Visibility        = "Visible"
    # Hide both tab panels so no empty space shows
    if ($null -ne $pJob.FilePrepPanel) { $pJob.FilePrepPanel.Visibility = "Collapsed" }
    if ($null -ne $pJob.EditingPanel)  { $pJob.EditingPanel.Visibility  = "Collapsed" }
}

# Put a single parent back into Edit mode
function Set-PJobEditMode($pJob) {
    $pJob.ReviewCardOverlay.Visibility = "Collapsed"
    $pJob.ReviewPanel.Visibility       = "Collapsed"
    if ($null -ne $pJob.ReviewFooter) { $pJob.ReviewFooter.Visibility = "Collapsed" }
    if ($null -ne $pJob.ColorsOverlayStack) { $pJob.ColorsOverlayStack.Visibility = "Visible" }
    if ($null -ne $pJob.CurrentThumb -and $pJob._SavedCurrentThumbVis) {
        $pJob.CurrentThumb.Visibility = $pJob._SavedCurrentThumbVis
    }
    $pJob.TasksBox.Visibility          = "Visible"
    $pJob.EditBox.Visibility           = "Visible"
    $pJob.PnlFiles.Visibility          = "Visible"
    $pJob.ApplyRow.Visibility          = "Visible"
    $pJob.BtnRefresh.Visibility        = "Visible"
    $pJob.BtnRemoveP.Visibility        = "Visible"
    $pJob.BtnKeepReview.Visibility     = "Collapsed"
    $pJob.BtnRevertReview.Visibility   = "Collapsed"
    # Restore the correct tab panel based on the current global mode
    $mode = if ([string]::IsNullOrEmpty($script:GlobalMode)) { "FilePr" } else { $script:GlobalMode }
    if ($null -ne $pJob.FilePrepPanel) { $pJob.FilePrepPanel.Visibility = if ($mode -eq "Editing") { "Collapsed" } else { "Visible" } }
    if ($null -ne $pJob.EditingPanel)  { $pJob.EditingPanel.Visibility  = if ($mode -eq "Editing") { "Visible"   } else { "Collapsed" } }
    # Reset review border and status label
    $pJob.RowPanel.BorderBrush     = Get-WpfColor "#2A2C35"
    $pJob.RowPanel.BorderThickness = New-Object System.Windows.Thickness(1)
    if ($null -ne $pJob.ReviewStatusLabel) {
        $pJob.ReviewStatusLabel.Text       = ""
        $pJob.ReviewStatusLabel.Visibility = "Collapsed"
    }
}

# Apply a review-mode state to a single group (direction: $true = enter, $false = exit)
function Apply-GpReviewMode($gpJob, [bool]$enter) {
    if ($enter) {
        $gpJob.ReviewMode = $true
        $gpJob.HeaderGrid.Background   = Get-WpfColor "#1A1C22"
        $gpJob.CbPrefix.IsEnabled      = $false
        $gpJob.TBTheme.IsEnabled       = $false
        $gpJob.ChkSkip.IsEnabled       = $false
        $gpJob.LblGpPreview.Visibility = "Collapsed"
        if ($null -ne $gpJob.ThemeBar) { $gpJob.ThemeBar.Visibility = "Collapsed" }
        if ($null -ne $gpJob.BtnCombineGp) { $gpJob.BtnCombineGp.Visibility = "Collapsed" }
        # The Printer/Tag/Theme/"Don't rename folder" controls aren't needed once in Review -
        # the folder name readout at the left of the header bar is enough.
        if ($null -ne $gpJob.RenameGroup) { $gpJob.RenameGroup.Visibility = "Collapsed" }
        foreach ($pj in $gpJob.Parents) { Set-PJobReviewMode $pj }
        if ($null -ne $gpJob.LblWigglitzAvg) {
            $wigglitzVals = @($gpJob.Parents | Where-Object { $null -ne $_.WigglitzValue } | ForEach-Object { $_.WigglitzValue })
            if ($wigglitzVals.Count -gt 0) {
                $avgVal = ($wigglitzVals | Measure-Object -Average).Average
                $gpJob.LblWigglitzAvg.Text = "Wigglitz/Day avg: {0:N2} ({1} designs)" -f $avgVal, $wigglitzVals.Count
            } else {
                $gpJob.LblWigglitzAvg.Text = "Wigglitz/Day avg: n/a"
            }
            $gpJob.LblWigglitzAvg.Visibility = "Visible"
        }
        # Theme-wide "X/Y designs ready" readout + theme-wide action buttons
        $totalCount = $gpJob.Parents.Count
        $readyCount = @($gpJob.Parents | Where-Object { $_.IsProdReady }).Count
        if ($null -ne $gpJob.LblReadyCount) {
            $gpJob.LblReadyCount.Text = "Ready: $readyCount/$totalCount"
            $gpJob.LblReadyCount.Foreground = if ($totalCount -gt 0 -and $readyCount -eq $totalCount) { Get-WpfColor "#4CAF72" } else { Get-WpfColor "#E8A135" }
            $gpJob.LblReadyCount.Visibility = "Visible"
        }
        $allReady = ($totalCount -gt 0 -and $readyCount -eq $totalCount)
        if ($null -ne $gpJob.BtnThPrintQReview) { $gpJob.BtnThPrintQReview.Visibility = "Visible" }
        if ($null -ne $gpJob.BtnThSendProdReview) {
            $gpJob.BtnThSendProdReview.Visibility = "Visible"
            $gpJob.BtnThSendProdReview.IsEnabled  = $allReady
            $gpJob.BtnThSendProdReview.Background = if ($allReady) { Get-WpfColor "#4CAF72" } else { Get-WpfColor "#3A3D4A" }
            $gpJob.BtnThSendProdReview.ToolTip = if ($allReady) {
                "Runs Send to Production for every design in this theme."
            } else {
                "Locked: $readyCount/$totalCount designs pass the Production Checklist. All designs must be ready first."
            }
        }
    } else {
        $gpJob.ReviewMode = $false
        $gpJob.HeaderGrid.Background   = Get-WpfColor "#2A2C35"
        $gpJob.CbPrefix.IsEnabled      = $true
        $gpJob.TBTheme.IsEnabled       = $true
        $gpJob.ChkSkip.IsEnabled       = $true
        $gpJob.LblGpPreview.Visibility = "Visible"
        if ($null -ne $gpJob.ThemeBar) { $gpJob.ThemeBar.Visibility = "Visible" }
        if ($null -ne $gpJob.BtnCombineGp) { $gpJob.BtnCombineGp.Visibility = "Visible" }
        if ($null -ne $gpJob.LblWigglitzAvg) { $gpJob.LblWigglitzAvg.Visibility = "Collapsed" }
        if ($null -ne $gpJob.LblReadyCount) { $gpJob.LblReadyCount.Visibility = "Collapsed" }
        if ($null -ne $gpJob.BtnThPrintQReview) { $gpJob.BtnThPrintQReview.Visibility = "Collapsed" }
        if ($null -ne $gpJob.BtnThSendProdReview) { $gpJob.BtnThSendProdReview.Visibility = "Collapsed" }
        foreach ($pj in $gpJob.Parents) { Set-PJobEditMode $pj }
    }
}

# Switch the global workspace mode and update every loaded group/card.
# $mode : "FilePr" | "Editing" | "Review"
function Set-GlobalMode([string]$mode) {
    $script:GlobalMode = $mode
    # Libraries mode: swap to the libraries panel and skip all card iteration
    $isLibraries = ($mode -eq "Libraries")
    if ($null -ne $scrollViewer)              { $scrollViewer.Visibility              = if ($isLibraries) { "Collapsed" } else { "Visible" } }
    if ($null -ne $script:LibrariesPanel)     { $script:LibrariesPanel.Visibility     = if ($isLibraries) { "Visible"   } else { "Collapsed" } }
    if ($isLibraries) {
        # Update top-bar button styles then return â€” no card panels to iterate
        if ($null -ne $script:BtnModeFilePr) {
            $activeStyle   = @{ bg = "#3A5080"; fg = "#FFFFFF"; border = "#5A78C4" }
            $inactiveStyle = @{ bg = "#252630"; fg = "#7A7D90"; border = "#3A3D50" }
            $libStyle      = @{ bg = "#5A3A80"; fg = "#FFFFFF"; border = "#8A5AC4" }
            foreach ($pair in @(
                @{ Btn=$script:BtnModeFilePr;   Active=($false); Style=$inactiveStyle },
                @{ Btn=$script:BtnModeEditing;  Active=($false); Style=$inactiveStyle },
                @{ Btn=$script:BtnModeReview;   Active=($false); Style=$inactiveStyle },
                @{ Btn=$script:BtnModeLibraries;Active=($true);  Style=$libStyle }
            )) {
                $pair.Btn.Background  = Get-WpfColor $pair.Style.bg
                $pair.Btn.Foreground  = Get-WpfColor $pair.Style.fg
                $pair.Btn.BorderBrush = Get-WpfColor $pair.Style.border
                $pair.Btn.IsEnabled   = -not $pair.Active
            }
        }
        return
    }
    $enterReview = ($mode -eq "Review")
    foreach ($gpJob in $script:jobs) {
        if ($enterReview -and -not $gpJob.ReviewMode) {
            Apply-GpReviewMode $gpJob $true
        } elseif (-not $enterReview -and $gpJob.ReviewMode) {
            Apply-GpReviewMode $gpJob $false
        }
        # For non-review modes, flip the tab panels on every card and the group task bars
        if (-not $enterReview) {
            if ($null -ne $gpJob.ThemeBar)        { $gpJob.ThemeBar.Visibility        = if ($mode -eq "Editing") { "Collapsed" } else { "Visible"   } }
            if ($null -ne $gpJob.EditingThemeBar) { $gpJob.EditingThemeBar.Visibility = if ($mode -eq "Editing") { "Visible"   } else { "Collapsed" } }
            if ($null -ne $gpJob.RenameGroup)     { $gpJob.RenameGroup.Visibility     = if ($mode -eq "Editing") { "Collapsed" } else { "Visible"   } }
            foreach ($pj in $gpJob.Parents) {
                if ($null -ne $pj.FilePrepPanel) {
                    $pj.FilePrepPanel.Visibility = if ($mode -eq "Editing") { "Collapsed" } else { "Visible" }
                }
                if ($null -ne $pj.EditingPanel) {
                    $pj.EditingPanel.Visibility  = if ($mode -eq "Editing") { "Visible" } else { "Collapsed" }
                }
                # Edit overlays: show in Editing mode, hide in FilePr/Review
                if ($null -ne $pj.RnOvLeft)  { $pj.RnOvLeft.Visibility  = if ($mode -eq "Editing") { "Visible" } else { "Collapsed" } }
                if ($null -ne $pj.RnOvRight) { $pj.RnOvRight.Visibility = if ($mode -eq "Editing") { "Visible" } else { "Collapsed" } }
            }
        }
    }
    # Update top-bar button styles
    if ($null -ne $script:BtnModeFilePr -and $null -ne $script:BtnModeEditing -and $null -ne $script:BtnModeReview) {
        $activeStyle   = @{ bg = "#3A5080"; fg = "#FFFFFF"; border = "#5A78C4" }
        $inactiveStyle = @{ bg = "#252630"; fg = "#7A7D90"; border = "#3A3D50" }
        $reviewActive  = @{ bg = "#7A5A2A"; fg = "#FFFFFF"; border = "#C4943A" }
        $libActive     = @{ bg = "#5A3A80"; fg = "#FFFFFF"; border = "#8A5AC4" }
        foreach ($pair in @(
            @{ Btn=$script:BtnModeFilePr;   Active=($mode -eq "FilePr");   Style=$activeStyle  },
            @{ Btn=$script:BtnModeEditing;  Active=($mode -eq "Editing");  Style=$activeStyle  },
            @{ Btn=$script:BtnModeReview;   Active=($mode -eq "Review");   Style=$reviewActive },
            @{ Btn=$script:BtnModeLibraries;Active=($mode -eq "Libraries");Style=$libActive    }
        )) {
            if ($null -eq $pair.Btn) { continue }
            $s = if ($pair.Active) { $pair.Style } else { $inactiveStyle }
            $pair.Btn.Background   = Get-WpfColor $s.bg
            $pair.Btn.Foreground   = Get-WpfColor $s.fg
            $pair.Btn.BorderBrush  = Get-WpfColor $s.border
            $pair.Btn.IsEnabled    = -not $pair.Active
        }
    }
}

# Saves a SKU value to the TSV for a given card folder.
# Called from both the per-card Save button and the queue-level CSV import.
function Save-SkuToTsv([string]$skuVal, [string]$folderPath, $anchorFile) {
    $af         = $anchorFile
    $anchorBase = if ($null -ne $af) { [System.IO.Path]::GetFileNameWithoutExtension($af.FullName) } else { "" }
    $tsvBase    = $anchorBase -replace '(?i)_?(Full|Nest|Final)$', ''
    $tsvPath    = Join-Path $folderPath "${tsvBase}_Data.tsv"
    if (Test-Path $tsvPath) {
        $lines = @(Get-Content $tsvPath)
        if ($lines.Count -gt 0) {
            $cols     = $lines[-1] -split "`t"
            $datePat  = '^\d{1,2}/\d{1,2}/\d{4}$'
            $isOldFmt = $cols.Count -gt 4 -and $cols[4] -match $datePat
            if ($isOldFmt) {
                $newCols   = $cols[0..2] + @($skuVal) + $cols[3..($cols.Count - 1)]
                $lines[-1] = $newCols -join "`t"
            } else {
                while ($cols.Count -lt 4) { $cols += "" }
                $cols[3]   = $skuVal
                $lines[-1] = $cols -join "`t"
            }
            Set-Content -Path $tsvPath -Value $lines -Encoding UTF8
        }
    } else {
        Set-Content -Path $tsvPath -Value "`t`t`t$skuVal" -Encoding UTF8
    }
}

# Searches sibling printer folders under the same theme root for a Data TSV
# that shares the same design name and already has a SKU assigned.
# Returns a hashtable { SKU; Printer; FilePath } or $null if nothing found.
# Theme root is 3 levels up from the design folder:
#   ThemeRoot / {Printer}_{Theme} / {Printer}_{Type}_{Theme} / {Printer}_{Design}_{Theme}
function Find-SiblingSkU([string]$folderPath, [string]$designName) {
    if ([string]::IsNullOrWhiteSpace($designName)) { return $null }
    $themeRoot = Split-Path (Split-Path (Split-Path $folderPath -Parent) -Parent) -Parent
    if ([string]::IsNullOrWhiteSpace($themeRoot) -or -not (Test-Path $themeRoot)) { return $null }

    $datePat = '^\d{1,2}/\d{1,2}/\d{4}$'
    $currentTsvItem = Get-ChildItem $folderPath -Filter "*_Data.tsv" -ErrorAction SilentlyContinue | Select-Object -First 1
    $currentTsv = if ($null -ne $currentTsvItem) { $currentTsvItem.FullName } else { '' }

    foreach ($tsv in (Get-ChildItem $themeRoot -Recurse -Filter "*_Data.tsv" -ErrorAction SilentlyContinue)) {
        if ($tsv.FullName -eq $currentTsv) { continue }
        try {
            $line = Get-Content $tsv.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if (-not $line) { continue }
            $cols = $line -split "`t"
            if ($cols.Count -lt 4 -or $cols[2].Trim() -ne $designName) { continue }
            $isOld = $cols.Count -gt 4 -and $cols[4] -match $datePat
            $sku   = if (-not $isOld -and $cols[3].Trim() -match $script:SkuPattern) { $cols[3].Trim() } else { '' }
            if ($sku -eq '') { continue }
            return @{ SKU = $sku; Printer = $cols[0].Trim(); FilePath = $tsv.FullName }
        } catch { continue }
    }
    return $null
}

function Build-PJob($parentPath, $anchorFile, $gpJob) {
    $tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -ItemType Directory -Path $tempWork | Out-Null

    $isZip3mf = $anchorFile.Extension -imatch '\.3mf$' -and $anchorFile.Name -notmatch '(?i)\.gcode\.3mf$'
    if ($isZip3mf) {
        try { [System.IO.Compression.ZipFile]::ExtractToDirectory($anchorFile.FullName, $tempWork) } catch {}
    }

    # Find the best 3MF for color info (Full > Nest > Final > any), independent of anchor.
    # If it differs from the anchor, extract it to a short-lived temp dir for reading only.
    $colorFile     = Find-ColorInfoFile $parentPath
    if ($null -eq $colorFile) { $colorFile = $anchorFile }
    $colorIsZip3mf = $colorFile.Extension -imatch '\.3mf$' -and $colorFile.Name -notmatch '(?i)\.gcode\.3mf$'
    $colorTempWork = $null
    $colorWorkDir  = $tempWork
    if ($colorIsZip3mf -and $colorFile.FullName -ne $anchorFile.FullName) {
        $colorTempWork = Join-Path $env:TEMP ("ColorInfo_" + [guid]::NewGuid().ToString().Substring(0,8))
        New-Item -ItemType Directory -Path $colorTempWork | Out-Null
        try { [System.IO.Compression.ZipFile]::ExtractToDirectory($colorFile.FullName, $colorTempWork) } catch {}
        $colorWorkDir = $colorTempWork
    }

    $activeSlots = New-Object System.Collections.ArrayList

    if ($colorIsZip3mf) {
        $projPath   = Join-Path $colorWorkDir "Metadata\project_settings.config"
        $modSetPath = Join-Path $colorWorkDir "Metadata\model_settings.config"

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
        # (filament_ids is 0-based). Use this when available â€” it's more precise than
        # scanning the whole model, and correctly excludes slots only used on other plates.
        $plateSlots = New-Object System.Collections.Generic.HashSet[string]
        $metaDir = Join-Path $colorWorkDir "Metadata"
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
            # Plate data found â€” use it directly; it already represents only what's on the plates
            $UsedSlots = $plateSlots
        } else {
            # No plate JSON â€” use build item positions to find on-plate objects, then read their extruders.
            # Bambu Studio stores off-plate objects at negative X; on-plate objects have positive X and Y
            # within reasonable bed bounds (up to ~450 mm covers all current Bambu models).
            $onPlateIds = New-Object System.Collections.Generic.HashSet[string]
            $modelFile = Get-ChildItem -Path $colorWorkDir -Filter '3dmodel.model' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($modelFile -and (Test-Path $modelFile.FullName)) {
                try {
                    $modelContent = [System.IO.File]::ReadAllText($modelFile.FullName, [System.Text.Encoding]::UTF8)
                    $itemRx = [regex]::Matches($modelContent, '<item\s[^>]*objectid="(\d+)"[^>]*transform="([^"]+)"')
                    foreach ($m in $itemRx) {
                        $tfParts = $m.Groups[2].Value -split '\s+'
                        if ($tfParts.Count -ge 12) {
                            $tx = 0.0; $ty = 0.0
                            if ([double]::TryParse($tfParts[9],  [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$tx) -and
                                [double]::TryParse($tfParts[10], [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ty)) {
                                if ($tx -ge 0 -and $ty -ge 0 -and $tx -le 450 -and $ty -le 450) {
                                    $onPlateIds.Add($m.Groups[1].Value) | Out-Null
                                }
                            }
                        }
                    }
                    # If no items parsed (old single-mesh format), also collect any materialid refs
                    if ($onPlateIds.Count -eq 0) {
                        $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
                        foreach ($mm in $matMatches) { $UsedSlots.Add($mm.Groups[1].Value) | Out-Null }
                    }
                } catch {}
            }

            if (Test-Path $modSetPath) {
                try {
                    [xml]$modXml = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
                    foreach ($obj in $modXml.config.object) {
                        # If we identified on-plate objects, skip any that aren't on the plate
                        if ($onPlateIds.Count -gt 0 -and -not $onPlateIds.Contains($obj.id)) { continue }
                        foreach ($node in $obj.SelectNodes('.//metadata[contains(@key,"extruder")]')) {
                            $val = $node.GetAttribute('value')
                            if (-not [string]::IsNullOrWhiteSpace($val)) { $UsedSlots.Add($val) | Out-Null }
                        }
                    }
                } catch {}
            }
        }

        foreach ($hex in $SlotMap.Keys) {
            $slotId = $SlotMap[$hex]
            if ($UsedSlots.Contains($slotId)) {
                $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
                $matchedName = if ($script:HexToName.Contains($checkHex)) { $script:HexToName[$checkHex] } else { "" }
                $activeSlots.Add([PSCustomObject]@{ OldHex = $checkHex; Name = $matchedName }) | Out-Null
            }
        }
        if ($activeSlots.Count -gt 8) { $activeSlots = $activeSlots[0..7] }
    }

    # Clean up the color-only temp dir (kept separate from $tempWork which is still needed for thumbnails)
    if ($null -ne $colorTempWork -and (Test-Path $colorTempWork)) {
        try { Remove-Item -Path $colorTempWork -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }

    $pJob = @{
        FolderPath = $parentPath; AnchorFile = $anchorFile; TempWork = $tempWork
        ProcessedAnchorPath = ""; CustomImagePath = $null
        UISlots = New-Object System.Collections.ArrayList
        FileRows = New-Object System.Collections.ArrayList
        IsDone = $false; IsQueued = $false; HasCollision = $false
        RenameOnlyBypass = $false; SliceOnlyBypass = $false; ColorOnlyBypass = $false
        ReviewBuilt = $false; WigglitzValue = $null; IsProdReady = $false
        GcodeImgPath = ""; ReviewCardOverlay = $null; ReviewPanel = $null; ReviewStack = $null
        ColorsOverlayStack = $null; CurrentThumb = $null; _SavedCurrentThumbVis = $null; ReviewFooter = $null
        TasksBox = $null; EditBox = $null; ApplyRow = $null
        BtnApply = $null; BtnRenameOnly = $null; BtnColorOnly = $null
        BtnRefresh = $null; BtnRemoveP = $null; BtnKeepReview = $null; BtnRevertReview = $null
        ReviewStatusLabel = $null
        BtnRunRenest = $null; RenestStatusLbl = $null
        FilePrepPanel = $null; EditingPanel = $null
        MainTabFilePrepBtn = $null; MainTabEditingBtn = $null
        RnOvLeft = $null; RnOvRight = $null; RnImgLeft = $null; RnImgRight = $null
        RnLblLeft = $null; RnLblRight = $null
        ChkEdReNest = $null
        BtnEdQueue = $null; LblEdQueueStatus = $null
        EdIsQueued = $false
        LblReviewPrintQStatus = $null
        _GpJob = $gpJob
    }

    # â”€â”€ Outer row border (the RowPanel equivalent) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
    # Row 0 holds the card/pick/controls content; Row 1 is the Review-mode action footer
    $pGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $pGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height=[System.Windows.GridLength]::Auto}))
    $pBorder.Child = $pGrid

    # Review-mode footer: bigger Printing Queue / Send to Production buttons,
    # anchored across the full width at the bottom of the card. Populated
    # lazily by Build-ReviewContent; hidden outside of Review mode.
    $reviewFooter = New-Object System.Windows.Controls.Border
    $reviewFooter.Background = Get-WpfColor "#1C1D23"; $reviewFooter.BorderBrush = Get-WpfColor "#2A2C35"
    $reviewFooter.BorderThickness = New-Object System.Windows.Thickness(0,1,0,0)
    $reviewFooter.Margin = New-Object System.Windows.Thickness(0,10,0,0); $reviewFooter.Padding = New-Object System.Windows.Thickness(0,10,0,0)
    $reviewFooter.Visibility = "Collapsed"
    [System.Windows.Controls.Grid]::SetRow($reviewFooter, 1)
    [System.Windows.Controls.Grid]::SetColumnSpan($reviewFooter, 2)
    $pGrid.Children.Add($reviewFooter) | Out-Null
    $pJob.ReviewFooter = $reviewFooter

# â”€â”€ LEFT COLUMN: Card panel + Pick panel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $leftGrid = New-Object System.Windows.Controls.Grid
    $leftGrid.VerticalAlignment = "Top"
    $leftGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))
    $leftGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))
    $leftGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height=[System.Windows.GridLength]::Auto}))
    $leftGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height=[System.Windows.GridLength]::Auto}))

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
        $pJob.CustomImagePath = $null  # STL or unknown â€” no preview available
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
                $pJob.GcodeImgPath = $gcodeImgPath
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

    # Scale swatch size and spacing to fit 5â€“8 color slots inside the 438px card
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
        $combo.IsTextSearchEnabled = $false  # We handle filtering ourselves
        $combo.MinWidth = $comboMinW; $combo.MaxWidth = 210 # Safe limits so it doesn't hit the image
        $combo.Width = [System.Double]::NaN # Tells WPF to Auto-size dynamically!
        $allColorKeys = @($script:LibraryColors.Keys | Sort-Object)
        $colorCollection = New-Object System.Collections.ObjectModel.ObservableCollection[string]
        foreach ($k in $allColorKeys) { $colorCollection.Add($k) | Out-Null }
        $combo.ItemsSource = $colorCollection
        # CollectionView filter â€” never touch the underlying collection, just swap predicates
        $comboView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($colorCollection)
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

        $combo.Tag = @{ StatusLbl = $lblStatus; OrigName = $slotData.Name; P = $pJob; Swatch = $swatchBorder; LblNum = $lblNum; AllKeys = $allColorKeys; ComboView = $comboView; Filtering = $false; NeedsFilter = $false; Confirmed = $true; TypedText = $(if ($slotData.Name) { $slotData.Name } else { "" }) }

        # Flag re-filter on typing/deletion; select-all when starting fresh after a confirmed value
        $combo.Add_PreviewTextInput({
            param($s, $e)
            $data = $s.Tag
            if ($data.Confirmed) {
                $data.Confirmed = $false
                $tb = $s.Template.FindName("PART_EditableTextBox", $s)
                if ($tb -ne $null) { $tb.SelectAll() }  # Replace confirmed value cleanly on first keystroke
            }
            $data.NeedsFilter = $true
        })

        # Arrow keys navigate filtered suggestions; Tab/Enter accepts; Escape restores typed text
        $combo.Add_PreviewKeyDown({
            param($s, $e)
            $data = $s.Tag
            if ($e.Key -eq [System.Windows.Input.Key]::Back -or $e.Key -eq [System.Windows.Input.Key]::Delete) {
                if ($data.Confirmed) {
                    $data.Confirmed = $false
                    $tb = $s.Template.FindName("PART_EditableTextBox", $s)
                    if ($tb -ne $null) { $tb.SelectAll() }  # Select all so delete clears the whole value
                }
                $data.NeedsFilter = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Down -or $e.Key -eq [System.Windows.Input.Key]::Up) {
                if (-not $s.IsDropDownOpen -and $s.Items.Count -gt 0) { $s.IsDropDownOpen = $true; $e.Handled = $true }
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Tab -or $e.Key -eq [System.Windows.Input.Key]::Enter) {
                if ($s.IsDropDownOpen -and $s.Items.Count -gt 0) {
                    $accepted = if ($s.SelectedItem -ne $null) { $s.SelectedItem.ToString() } else { $s.Items[0].ToString() }
                    $data.Filtering = $true
                    $s.IsDropDownOpen = $false
                    $data.ComboView.Filter = $null   # Clear filter â€” full list visible next open
                    $data.TypedText = $accepted
                    $data.NeedsFilter = $false
                    $data.Confirmed = $true           # Mark as confirmed so next keypress selects-all
                    $data.Filtering = $false          # Release BEFORE setting Text so TextChanged runs
                    $s.Text = $accepted               # TextChanged fires â†’ updates status/swatch/validate
                    $tb = $s.Template.FindName("PART_EditableTextBox", $s)
                    if ($tb -ne $null) { $tb.SelectAll() }  # Select all so next search replaces cleanly
                    $e.Handled = $true
                }
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                if ($s.IsDropDownOpen) {
                    $data.Filtering = $true
                    $s.IsDropDownOpen = $false
                    $data.ComboView.Filter = $null   # Clear filter
                    $data.NeedsFilter = $false
                    $data.Confirmed = $true
                    $data.Filtering = $false          # Release BEFORE restoring text
                    $s.Text = $data.TypedText         # TextChanged fires â†’ updates status
                    $tb = $s.Template.FindName("PART_EditableTextBox", $s)
                    if ($tb -ne $null) { $tb.SelectAll() }
                    $e.Handled = $true
                }
            }
        })

        # On mouse-click selection: clear filter so full list is ready next open; commit TypedText
        $combo.Add_DropDownClosed({
            param($s, $e)
            $data = $s.Tag
            if ($data.Filtering) { return }
            $data.Filtering = $true
            $data.ComboView.Filter = $null   # Just clear the predicate â€” collection itself never changes
            if ($script:LibraryColors.Contains($s.Text)) { $data.TypedText = $s.Text; $data.Confirmed = $true }
            $data.Filtering = $false
            # TextChanged already updated status when the item text was set
        })

        $combo.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, [System.Windows.Controls.TextChangedEventHandler]{
            param($s, $e)
            $data = $s.Tag
            if ($data.Filtering) { return }

            $typed = $s.Text

            # Re-filter only when user typed/deleted (NeedsFilter set by PreviewTextInput/PreviewKeyDown)
            # Arrow-key navigation leaves NeedsFilter = false so the filtered view is preserved
            if ($data.NeedsFilter) {
                $data.NeedsFilter = $false
                $data.Filtering = $true
                $data.TypedText = $typed
                $cv = $data.ComboView
                if ([string]::IsNullOrWhiteSpace($typed) -or $typed -eq "Select Color...") {
                    $cv.Filter = $null
                    $s.IsDropDownOpen = $false
                } else {
                    $typedLower = $typed.ToLower()
                    $cv.Filter = [Predicate[object]]({ param($item) $item.ToString().ToLower().Contains($typedLower) }.GetNewClosure())
                    $s.IsDropDownOpen = (-not $cv.IsEmpty)
                }
                $data.Filtering = $false
            }

            # Status + swatch always update (typing, navigation preview, Tab/Enter accept, mouse click)
            if ($script:LibraryColors.Contains($typed)) {
                $newHex = $script:LibraryColors[$typed]
                $data.Swatch.Background = Get-WpfColor $newHex
                try {
                    $rv = [Convert]::ToInt32($newHex.Substring(1,2), 16)
                    $gv = [Convert]::ToInt32($newHex.Substring(3,2), 16)
                    $bv = [Convert]::ToInt32($newHex.Substring(5,2), 16)
                    $nc = if ((0.299*$rv + 0.587*$gv + 0.114*$bv) -gt 128) { "#000000" } else { "#FFFFFF" }
                    $data.LblNum.Foreground = Get-WpfColor $nc
                } catch {}
            }
            if ($typed -eq $data.OrigName -and $data.OrigName) {
                $data.StatusLbl.Text = "[MATCHED]"; $data.StatusLbl.Foreground = Get-WpfColor "#4CAF72"
            } elseif ($script:LibraryColors.Contains($typed)) {
                $data.StatusLbl.Text = "[CHANGED]"; $data.StatusLbl.Foreground = Get-WpfColor "#E8A135"
            } else {
                $data.StatusLbl.Text = "[UNMATCHED]"; $data.StatusLbl.Foreground = Get-WpfColor "#D95F5F"
            }
            Validate-PJob $data.P
        })
        $slotIdx++
    }
    $cardGrid.Children.Add($colorsOverlayStack) | Out-Null
    $pJob.ColorsOverlayStack = $colorsOverlayStack

    # Processing Overlay (border + bottom label + progress bar)
    $cardBorderOverlay = New-Object System.Windows.Controls.Border
    $cardBorderOverlay.BorderThickness = New-Object System.Windows.Thickness(6)
    $cardBorderOverlay.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(220,232,161,53))
    $cardBorderOverlay.Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(30,232,161,53))
    $cardBorderOverlay.Visibility = "Collapsed"
    $cardStatusLbl = New-Object System.Windows.Controls.TextBlock
    $cardStatusLbl.Text = "[ PROCESSING ]"
    $cardStatusLbl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255,232,161,53))
    $cardStatusLbl.FontSize = 13; $cardStatusLbl.FontWeight = [System.Windows.FontWeights]::Bold
    $cardStatusLbl.TextAlignment = "Center"; $cardStatusLbl.HorizontalAlignment = "Stretch"
    $cardStatusLbl.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(180,0,0,0))
    $cardStatusLbl.Padding = New-Object System.Windows.Thickness(5,4,5,6)
    $cardProgressBar = New-Object System.Windows.Controls.ProgressBar
    $cardProgressBar.Height = 7; $cardProgressBar.Minimum = 0; $cardProgressBar.Maximum = 100; $cardProgressBar.Value = 0
    $cardProgressBar.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255,232,161,53))
    $cardProgressBar.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(180,0,0,0))
    $cardProgressBar.BorderThickness = New-Object System.Windows.Thickness(0)
    $cardProgressBar.HorizontalAlignment = "Stretch"; $cardProgressBar.Visibility = "Collapsed"
    $cardOverlayPanel = New-Object System.Windows.Controls.StackPanel
    $cardOverlayPanel.VerticalAlignment = "Bottom"; $cardOverlayPanel.HorizontalAlignment = "Stretch"
    $cardOverlayPanel.Children.Add($cardProgressBar) | Out-Null
    $cardOverlayPanel.Children.Add($cardStatusLbl) | Out-Null
    $cardBorderOverlay.Child = $cardOverlayPanel
    $cardGrid.Children.Add($cardBorderOverlay) | Out-Null
    $pJob.ProcessingOverlay  = $cardBorderOverlay
    $pJob.CardStatusLabel    = $cardStatusLbl
    $pJob.CardProgressBar    = $cardProgressBar

    # Review-mode plate overlay â€” covers the full card area; shown only in Review mode
    $reviewCardOverlay = New-Object System.Windows.Controls.Image
    $reviewCardOverlay.Stretch = "Uniform"; $reviewCardOverlay.Visibility = "Collapsed"
    $reviewCardOverlay.Cursor = [System.Windows.Input.Cursors]::Hand
    $reviewCardOverlay.Add_MouseLeftButtonDown({
        if ($_.ClickCount -ge 2 -and $null -ne $this.Source) {
            $viewer = New-Object System.Windows.Window
            $viewer.Title = "Plate Preview (gcode)"; $viewer.Background = Get-WpfColor "#0D0E10"
            $viewer.SizeToContent = "WidthAndHeight"; $viewer.WindowStartupLocation = "CenterScreen"
            $viewer.ResizeMode = "CanResizeWithGrip"
            $imgV = New-Object System.Windows.Controls.Image
            $imgV.Source = $this.Source; $imgV.MaxWidth = 900; $imgV.MaxHeight = 900; $imgV.Stretch = "Uniform"
            $imgV.Margin = New-Object System.Windows.Thickness(10)
            $viewer.Content = $imgV; $viewer.ShowDialog() | Out-Null
        }
    })
    $cardGrid.Children.Add($reviewCardOverlay) | Out-Null
    $pJob.ReviewCardOverlay = $reviewCardOverlay

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
        if ($files -and $files.Count -gt 0 -and [System.IO.Path]::GetExtension($files[0]) -imatch '\.(png|jpg|jpeg)$') {
            $job = $this.Tag
            $srcPath = $files[0]
            $destName = [System.IO.Path]::GetFileName($srcPath)
            $destPath = Join-Path $job.FolderPath $destName

            # Remove any existing non-slicePreview PNGs in the folder first
            Get-ChildItem -Path $job.FolderPath -Filter "*.png" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" } |
                ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }

            # Copy the dropped file
            if ($srcPath -ne $destPath) { Copy-Item -Path $srcPath -Destination $destPath -Force }
            $job.CustomImagePath = $destPath
            $job.PbPlate.Source  = Load-WpfImage $destPath
            if ($job.BtnBrowseImg) { $job.BtnBrowseImg.Background = Get-WpfColor "#4CAF72" }

            # Refresh the file-list panel
            $job.FileRows.Clear()
            $job.PnlFiles.Children.Clear()
            $gj = $job._GpJob
            $folderFiles = Get-ChildItem -Path $job.FolderPath -File -ErrorAction SilentlyContinue |
                Sort-Object { switch -Regex ($_.Name) { 'Final\.3mf$' {0} 'Nest\.3mf$' {1} 'Full\.3mf$' {2} 'Full\.gcode\.3mf$' {3} default {4} } }, Name
            foreach ($fi2 in $folderFiles) { Add-FileRow $job $gj $fi2 }
            Update-ParentPreview $job $gj
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
                $base     = (Split-Path $t.Original3mf -Leaf) -replace '(?i)_?(Full\.gcode\.3mf|Full\.3mf|Final\.3mf|Nest\.3mf)$', ''
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
    $pickStatusLbl.TextAlignment = "Center"; $pickStatusLbl.HorizontalAlignment = "Stretch"
    $pickStatusLbl.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(180,0,0,0))
    $pickStatusLbl.Padding = New-Object System.Windows.Thickness(5,4,5,6)
    $pickProgressBar = New-Object System.Windows.Controls.ProgressBar
    $pickProgressBar.Height = 7; $pickProgressBar.Minimum = 0; $pickProgressBar.Maximum = 100; $pickProgressBar.Value = 0
    $pickProgressBar.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255,232,161,53))
    $pickProgressBar.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(180,0,0,0))
    $pickProgressBar.BorderThickness = New-Object System.Windows.Thickness(0)
    $pickProgressBar.HorizontalAlignment = "Stretch"; $pickProgressBar.Visibility = "Collapsed"
    $pickOverlayPanel = New-Object System.Windows.Controls.StackPanel
    $pickOverlayPanel.VerticalAlignment = "Bottom"; $pickOverlayPanel.HorizontalAlignment = "Stretch"
    $pickOverlayPanel.Children.Add($pickProgressBar) | Out-Null
    $pickOverlayPanel.Children.Add($pickStatusLbl) | Out-Null
    $pickBorderOverlay.Child = $pickOverlayPanel
    $pickGrid.Children.Add($pickBorderOverlay) | Out-Null
    $pJob.PickProcessingOverlay = $pickBorderOverlay
    $pJob.PickStatusLabel       = $pickStatusLbl
    $pJob.PickProgressBar       = $pickProgressBar

    # Merge detected banner (top of pick image)
    $nestExists = Get-ChildItem -Path $parentPath -Filter "*Nest.3mf" -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '(?i)[._ ]Nest\.3mf$' } | Select-Object -First 1
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

    # Finished image overlay â€” spans both card and pick columns, shown after processing completes
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

    # â”€â”€ RIGHT: Controls column â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $rightStack = New-Object System.Windows.Controls.StackPanel
    $rightStack.Orientation = "Vertical"; $rightStack.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($rightStack, 1); $pGrid.Children.Add($rightStack) | Out-Null

    # Header: folder name on its own line, then the action buttons below
    $headerStack = New-Object System.Windows.Controls.StackPanel; $headerStack.Orientation = "Vertical"

    $lblFolder = Create-TextBlock "Folder: $(Split-Path $parentPath -Leaf)" "#FFFFFF" 14 "Bold"
    $lblFolder.Margin = New-Object System.Windows.Thickness(0,0,0,6)
    $headerStack.Children.Add($lblFolder) | Out-Null
    $pJob.LblFolder = $lblFolder

    $btnHdrStack = New-Object System.Windows.Controls.StackPanel; $btnHdrStack.Orientation = "Horizontal"

    $btnRefresh = New-Object System.Windows.Controls.Button
    $btnRefresh.Content = "Refresh"; $btnRefresh.Background = Get-WpfColor "#5A78C4"; $btnRefresh.Foreground = Get-WpfColor "#FFFFFF"
    $btnRefresh.Width = 100; $btnRefresh.Height = 25; $btnRefresh.BorderThickness = 0
    $btnRefresh.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnRefresh.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRefresh.Tag = @{ P = $pJob; G = $gpJob }
    $btnRefresh.Add_Click({ $t = $this.Tag; Refresh-PJob $t.P $t.G })
    $btnHdrStack.Children.Add($btnRefresh) | Out-Null
    $pJob.BtnRefresh = $btnRefresh

    $btnOpenFolder = New-Object System.Windows.Controls.Button
    $btnOpenFolder.Content = "Open Folder"; $btnOpenFolder.Background = Get-WpfColor "#2E7D8A"; $btnOpenFolder.Foreground = Get-WpfColor "#FFFFFF"
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
    $pJob.BtnRemoveP = $btnRemoveP

    # Keep button (Review mode only â€” marks design as checked, green border)
    $btnKeepReview = New-Object System.Windows.Controls.Button
    $btnKeepReview.Content = "Keep"; $btnKeepReview.Background = Get-WpfColor "#2E7D32"; $btnKeepReview.Foreground = Get-WpfColor "#FFFFFF"
    $btnKeepReview.Width = 70; $btnKeepReview.Height = 25; $btnKeepReview.BorderThickness = 0
    $btnKeepReview.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnKeepReview.Visibility = "Collapsed"; $btnKeepReview.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnKeepReview.Tag = $pJob
    $btnKeepReview.Add_Click({
        $pj = $this.Tag
        $pj.RowPanel.BorderBrush     = Get-WpfColor "#4CAF72"
        $pj.RowPanel.BorderThickness = New-Object System.Windows.Thickness(15)
        $pj.ReviewStatusLabel.Text       = "CHECKED"
        $pj.ReviewStatusLabel.Foreground = Get-WpfColor "#4CAF72"
        $pj.ReviewStatusLabel.Visibility = "Visible"
    })
    $btnHdrStack.Children.Add($btnKeepReview) | Out-Null
    $pJob.BtnKeepReview = $btnKeepReview

    # Revert button (Review mode only â€” runs the revert bat, marks red)
    $btnRevertReview = New-Object System.Windows.Controls.Button
    $btnRevertReview.Content = "Revert"; $btnRevertReview.Background = Get-WpfColor "#D95F5F"; $btnRevertReview.Foreground = Get-WpfColor "#FFFFFF"
    $btnRevertReview.Width = 70; $btnRevertReview.Height = 25; $btnRevertReview.BorderThickness = 0
    $btnRevertReview.Margin = New-Object System.Windows.Thickness(0,0,0,0); $btnRevertReview.Visibility = "Collapsed"; $btnRevertReview.Cursor = [System.Windows.Input.Cursors]::Hand
    if (-not $nestExists) { $btnRevertReview.IsEnabled = $false; $btnRevertReview.Background = Get-WpfColor "#555555" }
    $btnRevertReview.Tag = @{ P = $pJob; G = $gpJob }
    $btnRevertReview.Add_Click({
        $t = $this.Tag; $pj = $t.P
        $batPath = Join-Path $scriptDir "..\callers\RevertMerge.bat"
        if (-not (Test-Path $batPath)) {
            [System.Windows.MessageBox]::Show("RevertMerge.bat not found.", "Error") | Out-Null; return
        }
        $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }
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
        $pj.RowPanel.BorderBrush     = Get-WpfColor "#D95F5F"
        $pj.RowPanel.BorderThickness = New-Object System.Windows.Thickness(3)
        $pj.ReviewStatusLabel.Text       = "REVERTED"
        $pj.ReviewStatusLabel.Foreground = Get-WpfColor "#D95F5F"
        $pj.ReviewStatusLabel.Visibility = "Visible"
    })
    $btnHdrStack.Children.Add($btnRevertReview) | Out-Null
    $pJob.BtnRevertReview = $btnRevertReview

    $headerStack.Children.Add($btnHdrStack) | Out-Null
    $rightStack.Children.Add($headerStack) | Out-Null

    # â”€â”€ Tab content panels (visibility controlled globally from the top-bar buttons)
    $filePrepPanel = New-Object System.Windows.Controls.StackPanel
    $filePrepPanel.Visibility = if ($script:GlobalMode -eq "Editing") { "Collapsed" } else { "Visible" }
    $rightStack.Children.Add($filePrepPanel) | Out-Null

    $editingPanel = New-Object System.Windows.Controls.StackPanel
    $editingPanel.Visibility = if ($script:GlobalMode -eq "Editing") { "Visible" } else { "Collapsed" }
    $rightStack.Children.Add($editingPanel) | Out-Null

    $pJob.FilePrepPanel = $filePrepPanel; $pJob.EditingPanel = $editingPanel
    $pJob.MainTabFilePrepBtn = $null; $pJob.MainTabEditingBtn = $null

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
    $chkColors  = New-Object System.Windows.Controls.CheckBox; $chkColors.Content  = "Save Colors";          $chkColors.IsChecked  = $false; $chkColors.Foreground  = Get-WpfColor "#FFFFFF"; $chkColors.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkColors.ToolTip = "Write current color selections to the .3mf files only, skipping all other tasks"
    $chkLogs    = New-Object System.Windows.Controls.CheckBox; $chkLogs.Content    = "Create Logs";          $chkLogs.IsChecked    = $false; $chkLogs.Foreground    = Get-WpfColor "#FFFFFF"; $chkLogs.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkBOD     = New-Object System.Windows.Controls.CheckBox; $chkBOD.Content     = "Create BOD";           $chkBOD.IsChecked     = $false; $chkBOD.Foreground     = Get-WpfColor "#FFFFFF"; $chkBOD.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkBOD.ToolTip = "Reduces the merged Full.3mf to the 5 pairs closest to centre and exports a BOD.gcode.3mf to the Printing Queue"
    $chkPrintQ  = New-Object System.Windows.Controls.CheckBox; $chkPrintQ.Content  = "Printing Queue";       $chkPrintQ.IsChecked  = $false; $chkPrintQ.Foreground  = Get-WpfColor "#FFFFFF"
    $chkPrintQ.ToolTip = "Copies the existing Full.gcode.3mf to the Printing Queue folder with today's date"
    $tasksRow1.Children.Add($chkRename) | Out-Null; $tasksRow1.Children.Add($chkMerge) | Out-Null; $tasksRow1.Children.Add($chkSlice) | Out-Null
    $tasksRow1.Children.Add($chkExtract) | Out-Null; $tasksRow1.Children.Add($chkImage) | Out-Null; $tasksRow1.Children.Add($chkColors) | Out-Null
    $tasksRow1.Children.Add($chkLogs) | Out-Null; $tasksRow1.Children.Add($chkBOD) | Out-Null; $tasksRow1.Children.Add($chkPrintQ) | Out-Null
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

    # SKU row â€” locked by default; Edit button unlocks for manual entry
    $tasksRow3 = New-Object System.Windows.Controls.StackPanel
    $tasksRow3.Orientation = "Horizontal"; $tasksRow3.Margin = New-Object System.Windows.Thickness(0,8,0,0)

    $lblSku = New-Object System.Windows.Controls.TextBlock
    $lblSku.Text = "SKU:"; $lblSku.Foreground = Get-WpfColor "#A0A0A0"; $lblSku.FontSize = 12
    $lblSku.VerticalAlignment = "Center"; $lblSku.Margin = New-Object System.Windows.Thickness(0,0,8,0)

    $txtSku = New-Object System.Windows.Controls.TextBox
    $txtSku.Width = 140; $txtSku.Height = 26; $txtSku.FontSize = 12
    $txtSku.Background = Get-WpfColor "#1A1C23"; $txtSku.Foreground = Get-WpfColor "#888888"
    $txtSku.BorderBrush = Get-WpfColor "#333333"; $txtSku.Padding = New-Object System.Windows.Thickness(4,2,4,2)
    $txtSku.Margin = New-Object System.Windows.Thickness(0,0,8,0); $txtSku.VerticalContentAlignment = "Center"
    $txtSku.IsReadOnly = $true

    # Pre-populate SKU from existing TSV; also capture design name for sibling lookup.
    # Handles three TSV formats:
    #   New format    : Printer | FileType | DesignName | SKU | Theme | Date | ...
    #   Old format    : Printer | FileType | DesignName | Theme | Date | ...  (Date at col 4)
    #   Very old fmt  : FullFolderName | Theme | Date | H | M | ...           (Date at col 2)
    # For the very old format col[2] is a date, so we fall back to deriving
    # the design name from the folder leaf (strip leading Printer prefix token).
    $skuDesignName      = ''
    $existingSkuTsvFile = Get-ChildItem -Path $pJob.FolderPath -Filter "*_Data.tsv" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existingSkuTsvFile) {
        try {
            $skuSeedLine = Get-Content $existingSkuTsvFile.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ($skuSeedLine) {
                $skuSeedCols   = $skuSeedLine -split "`t"
                $skuDatePat    = '^\d{1,2}/\d{1,2}/\d{4}$'
                $skuOldFmt     = $skuSeedCols.Count -gt 4 -and $skuSeedCols[4] -match $skuDatePat
                # Design name at col[2] only when col[2] is not a date (very old format guard)
                if ($skuSeedCols.Count -ge 3 -and $skuSeedCols[2].Trim() -notmatch $skuDatePat) {
                    $skuDesignName = $skuSeedCols[2].Trim()
                } else {
                    # Very old format: derive design name from folder leaf by stripping printer prefix
                    $folderLeaf    = Split-Path $pJob.FolderPath -Leaf
                    $prefixEnd     = $folderLeaf.IndexOf('_')
                    $skuDesignName = if ($prefixEnd -ge 0) { $folderLeaf.Substring($prefixEnd + 1) } else { $folderLeaf }
                }
                if (-not $skuOldFmt -and $skuSeedCols.Count -ge 4 -and $skuSeedCols[3].Trim() -match $script:SkuPattern) {
                    $txtSku.Text = $skuSeedCols[3].Trim()
                }
            }
        } catch {}
    }

    # If no SKU yet, look for one in a sibling printer variant under the same theme root
    $siblingSkuInfo = $null
    if ($txtSku.Text -eq '' -and $skuDesignName -ne '') {
        $siblingSkuInfo = Find-SiblingSkU $pJob.FolderPath $skuDesignName
    }

    # Sync SKU button â€” visible only when a sibling SKU is available and this card has none
    $btnSyncSku = New-Object System.Windows.Controls.Button
    $btnSyncSku.Content    = if ($null -ne $siblingSkuInfo) { "Sync from $($siblingSkuInfo.Printer)" } else { "Sync SKU" }
    $btnSyncSku.Height     = 26; $btnSyncSku.Padding = New-Object System.Windows.Thickness(8,0,8,0)
    $btnSyncSku.Background = Get-WpfColor "#2A5A7A"; $btnSyncSku.Foreground = Get-WpfColor "#FFFFFF"
    $btnSyncSku.BorderThickness = 0; $btnSyncSku.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnSyncSku.Margin     = New-Object System.Windows.Thickness(0,0,4,0)
    $btnSyncSku.Visibility = if ($null -ne $siblingSkuInfo) { "Visible" } else { "Collapsed" }

    # Edit button â€” shown when locked; unlocks the box for manual entry
    $btnEditSku = New-Object System.Windows.Controls.Button
    $btnEditSku.Content = "Edit"; $btnEditSku.Height = 26; $btnEditSku.Width = 46
    $btnEditSku.Background = Get-WpfColor "#3A3C45"; $btnEditSku.Foreground = Get-WpfColor "#CCCCCC"
    $btnEditSku.BorderThickness = 0; $btnEditSku.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnEditSku.Margin = New-Object System.Windows.Thickness(0,0,4,0)

    # Save button â€” hidden until editing; commits value and re-locks
    $btnSaveSku = New-Object System.Windows.Controls.Button
    $btnSaveSku.Content = "Save"; $btnSaveSku.Height = 26; $btnSaveSku.Width = 46
    $btnSaveSku.Background = Get-WpfColor "#2E5A42"; $btnSaveSku.Foreground = Get-WpfColor "#FFFFFF"
    $btnSaveSku.BorderThickness = 0; $btnSaveSku.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnSaveSku.Visibility = "Collapsed"; $btnSaveSku.Margin = New-Object System.Windows.Thickness(0,0,4,0)

    # Cancel button â€” hidden until editing; restores original value and re-locks
    $btnCancelSku = New-Object System.Windows.Controls.Button
    $btnCancelSku.Content = "Cancel"; $btnCancelSku.Height = 26; $btnCancelSku.Width = 56
    $btnCancelSku.Background = Get-WpfColor "#3A2020"; $btnCancelSku.Foreground = Get-WpfColor "#CCCCCC"
    $btnCancelSku.BorderThickness = 0; $btnCancelSku.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnCancelSku.Visibility = "Collapsed"; $btnCancelSku.Margin = New-Object System.Windows.Thickness(0,0,4,0)

    $lblSkuStatus = New-Object System.Windows.Controls.TextBlock
    $lblSkuStatus.FontSize = 11; $lblSkuStatus.Margin = New-Object System.Windows.Thickness(8,0,0,0)
    $lblSkuStatus.VerticalAlignment = "Center"; $lblSkuStatus.Text = ""

    $skuTag = @{
        TxtSku      = $txtSku
        BtnSync     = $btnSyncSku
        BtnEdit     = $btnEditSku
        BtnSave     = $btnSaveSku
        BtnCancel   = $btnCancelSku
        LblStatus   = $lblSkuStatus
        FolderPath  = $pJob.FolderPath
        AnchorFile  = $pJob.AnchorFile
        SiblingInfo = $siblingSkuInfo
        OriginalVal = ""
    }

    $btnSyncSku.Tag = $skuTag
    $btnSyncSku.Add_Click({
        $t = $this.Tag
        if ($null -eq $t.SiblingInfo) { return }
        try {
            Save-SkuToTsv $t.SiblingInfo.SKU $t.FolderPath $t.AnchorFile
            $t.TxtSku.Text = $t.SiblingInfo.SKU
            $t.LblStatus.Text = "Synced from $($t.SiblingInfo.Printer)"
            $t.LblStatus.Foreground = Get-WpfColor "#4CAF72"
            $this.Visibility = "Collapsed"
            $t.BtnEdit.Visibility = "Visible"
        } catch {
            $t.LblStatus.Text = "Sync failed: $($_.Exception.Message)"
            $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"
        }
    }.GetNewClosure())

    $btnEditSku.Tag = $skuTag
    $btnEditSku.Add_Click({
        $t = $this.Tag
        $t.OriginalVal = $t.TxtSku.Text
        $t.TxtSku.IsReadOnly = $false
        $t.TxtSku.Background = Get-WpfColor "#2A2C35"
        $t.TxtSku.Foreground = Get-WpfColor "#FFFFFF"
        $t.TxtSku.BorderBrush = Get-WpfColor "#5A78C4"
        $this.Visibility = "Collapsed"
        $t.BtnSync.Visibility = "Collapsed"
        $t.BtnSave.Visibility = "Visible"
        $t.BtnCancel.Visibility = "Visible"
        $t.LblStatus.Text = ""
        $t.TxtSku.Focus() | Out-Null
        $t.TxtSku.SelectAll()
    }.GetNewClosure())

    $btnSaveSku.Tag = $skuTag
    $btnSaveSku.Add_Click({
        $t = $this.Tag
        try {
            $skuVal = $t.TxtSku.Text.Trim()
            if ($skuVal -notmatch $script:SkuPattern) {
                $t.LblStatus.Text = "SKU must be 3+ chars, no spaces"
                $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"
                return
            }
            Save-SkuToTsv $skuVal $t.FolderPath $t.AnchorFile
            $t.TxtSku.IsReadOnly = $true
            $t.TxtSku.Background = Get-WpfColor "#1A1C23"
            $t.TxtSku.Foreground = Get-WpfColor "#888888"
            $t.TxtSku.BorderBrush = Get-WpfColor "#333333"
            $t.BtnEdit.Visibility = "Visible"
            $this.Visibility = "Collapsed"
            $t.BtnCancel.Visibility = "Collapsed"
            $t.LblStatus.Text = "Saved"; $t.LblStatus.Foreground = Get-WpfColor "#4CAF72"
        } catch {
            $t.LblStatus.Text = "Failed: $($_.Exception.Message)"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"
        }
    }.GetNewClosure())

    $btnCancelSku.Tag = $skuTag
    $btnCancelSku.Add_Click({
        $t = $this.Tag
        $t.TxtSku.Text = $t.OriginalVal
        $t.TxtSku.IsReadOnly = $true
        $t.TxtSku.Background = Get-WpfColor "#1A1C23"
        $t.TxtSku.Foreground = Get-WpfColor "#888888"
        $t.TxtSku.BorderBrush = Get-WpfColor "#333333"
        $t.BtnEdit.Visibility = "Visible"
        # Restore Sync button if SKU is still empty and sibling info is available
        if ($null -ne $t.SiblingInfo -and $t.TxtSku.Text -eq '') {
            $t.BtnSync.Visibility = "Visible"
        }
        $t.BtnSave.Visibility = "Collapsed"
        $this.Visibility = "Collapsed"
        $t.LblStatus.Text = ""
    }.GetNewClosure())

    $tasksRow3.Children.Add($lblSku) | Out-Null
    $tasksRow3.Children.Add($txtSku) | Out-Null
    $tasksRow3.Children.Add($btnSyncSku) | Out-Null
    $tasksRow3.Children.Add($btnEditSku) | Out-Null
    $tasksRow3.Children.Add($btnSaveSku) | Out-Null
    $tasksRow3.Children.Add($btnCancelSku) | Out-Null
    $tasksRow3.Children.Add($lblSkuStatus) | Out-Null
    $tasksOuter.Children.Add($tasksRow3) | Out-Null

    $tasksBox.Child = $tasksOuter
    $filePrepPanel.Children.Add($tasksBox) | Out-Null
    $pJob.TasksBox = $tasksBox

    $pJob.ChkRename = $chkRename; $pJob.ChkMerge = $chkMerge; $pJob.ChkSlice = $chkSlice; $pJob.ChkExtract = $chkExtract; $pJob.ChkImage = $chkImage; $pJob.ChkColors = $chkColors; $pJob.ChkLogs = $chkLogs; $pJob.ChkBOD = $chkBOD; $pJob.ChkPrintQ = $chkPrintQ
    $pJob.TxtSKU     = $txtSku
    $pJob.DesignName = $skuDesignName
    $pJob.BtnSyncSku = $btnSyncSku
    $pJob.SkuTag     = $skuTag
    if ($nestExists) {
        $chkMerge.IsChecked = $false; $chkMerge.IsEnabled = $false; $chkMerge.Foreground = Get-WpfColor "#555555"
        $chkMerge.ToolTip = "Remove Nest.3mf or Revert Merge before merging again"
    }

    # Re-evaluate Rename Only button visibility whenever the Rename checkbox is toggled
    $chkRename.Tag = $pJob
    $chkRename.Add_Checked({   Validate-PJob $this.Tag })
    $chkRename.Add_Unchecked({ Validate-PJob $this.Tag })

    # Checkbox interdependencies
    $tasksData = @{ Rename = $chkRename; Merge = $chkMerge; Slice = $chkSlice; Extract = $chkExtract; Image = $chkImage; Colors = $chkColors; Logs = $chkLogs; BOD = $chkBOD; PrintQ = $chkPrintQ; PJob = $pJob; GpJob = $gpJob }

    # Re-validate Add to Queue / Process buttons whenever any task checkbox changes
    foreach ($chkName in @('ChkMerge','ChkSlice','ChkExtract','ChkImage','ChkColors','ChkLogs','ChkBOD','ChkPrintQ')) {
        $chk = $pJob.$chkName
        $chk.Add_Checked({ Validate-PJob $pJob }.GetNewClosure())
        $chk.Add_Unchecked({ Validate-PJob $pJob }.GetNewClosure())
    }

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
        $t.Rename.IsChecked = $true; $t.Slice.IsChecked = $true; $t.Extract.IsChecked = $true; $t.Image.IsChecked = $true; $t.Colors.IsChecked = $true
        if ($t.Merge.IsEnabled) { $t.Merge.IsChecked = $true }
        # BOD is intentionally left at its current value â€” it defaults off and is a special-purpose task
    })

    $btnDeselAll.Tag = $tasksData
    $btnDeselAll.Add_Click({
        $t = $this.Tag
        $t.Rename.IsChecked = $false; $t.Slice.IsChecked = $false; $t.Extract.IsChecked = $false; $t.Image.IsChecked = $false; $t.Colors.IsChecked = $false; $t.Logs.IsChecked = $false; $t.BOD.IsChecked = $false; $t.PrintQ.IsChecked = $false
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
    $editBox.Child = $editStack; $filePrepPanel.Children.Add($editBox) | Out-Null
    $pJob.EditBox = $editBox

    # Review panel â€” shown only in Review mode (lazy-populated on first toggle)
    $reviewPanelBorder = New-Object System.Windows.Controls.Border
    $reviewPanelBorder.Background = Get-WpfColor "#1C1D23"; $reviewPanelBorder.BorderBrush = Get-WpfColor "#2A2C35"
    $reviewPanelBorder.BorderThickness = New-Object System.Windows.Thickness(1)
    $reviewPanelBorder.Margin = New-Object System.Windows.Thickness(0,10,0,0); $reviewPanelBorder.Padding = New-Object System.Windows.Thickness(10)
    $reviewPanelBorder.Visibility = "Collapsed"
    $reviewStack = New-Object System.Windows.Controls.StackPanel
    $reviewPanelBorder.Child = $reviewStack
    $rightStack.Children.Add($reviewPanelBorder) | Out-Null
    $pJob.ReviewPanel = $reviewPanelBorder; $pJob.ReviewStack = $reviewStack

    # Files list
    $pnlFiles = New-Object System.Windows.Controls.StackPanel; $pnlFiles.Margin = New-Object System.Windows.Thickness(0,10,0,0)
    $filePrepPanel.Children.Add($pnlFiles) | Out-Null; $pJob.PnlFiles = $pnlFiles
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

    $btnColorOnly = New-Object System.Windows.Controls.Button
    $btnColorOnly.Content = "Save Colors"; $btnColorOnly.Background = Get-WpfColor "#4A7A9B"; $btnColorOnly.Foreground = Get-WpfColor "#FFFFFF"
    $btnColorOnly.FontWeight = [System.Windows.FontWeights]::Bold; $btnColorOnly.Width = 110; $btnColorOnly.Height = 35; $btnColorOnly.BorderThickness = 0; $btnColorOnly.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnColorOnly.Margin = New-Object System.Windows.Thickness(0,0,10,0)
    $btnColorOnly.ToolTip = "Write current color selections to the .3mf files only, skipping all other tasks"
    $applyRow.Children.Add($btnColorOnly) | Out-Null; $pJob.BtnColorOnly = $btnColorOnly

    $btnRenameOnly = New-Object System.Windows.Controls.Button
    $btnRenameOnly.Content = "Rename Only"; $btnRenameOnly.Background = Get-WpfColor "#5A6A8A"; $btnRenameOnly.Foreground = Get-WpfColor "#FFFFFF"
    $btnRenameOnly.FontWeight = [System.Windows.FontWeights]::Bold; $btnRenameOnly.Width = 110; $btnRenameOnly.Height = 35; $btnRenameOnly.BorderThickness = 0; $btnRenameOnly.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRenameOnly.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnRenameOnly.Visibility = "Collapsed"
    $btnRenameOnly.ToolTip = "Rename files only, skipping color validation and heavy tasks"
    $applyRow.Children.Add($btnRenameOnly) | Out-Null; $pJob.BtnRenameOnly = $btnRenameOnly

    $btnApply = New-Object System.Windows.Controls.Button
    $btnApply.Content = "Add to Queue"; $btnApply.Background = Get-WpfColor "#4CAF72"; $btnApply.Foreground = Get-WpfColor "#FFFFFF"
    $btnApply.FontWeight = [System.Windows.FontWeights]::Bold; $btnApply.Width = 150; $btnApply.Height = 35; $btnApply.BorderThickness = 0; $btnApply.Cursor = [System.Windows.Input.Cursors]::Hand
    $applyRow.Children.Add($btnApply) | Out-Null; $pJob.BtnApply = $btnApply

    $btnRevertDone = New-Object System.Windows.Controls.Button
    $btnRevertDone.Content = "REVERT"; $btnRevertDone.Background = Get-WpfColor "#D95F5F"; $btnRevertDone.Foreground = Get-WpfColor "#FFFFFF"
    $btnRevertDone.FontWeight = [System.Windows.FontWeights]::Bold; $btnRevertDone.Width = 75; $btnRevertDone.Height = 35; $btnRevertDone.BorderThickness = 0
    $btnRevertDone.Margin = New-Object System.Windows.Thickness(10,0,0,0); $btnRevertDone.Visibility = "Collapsed"; $btnRevertDone.Cursor = [System.Windows.Input.Cursors]::Hand
    $applyRow.Children.Add($btnRevertDone) | Out-Null; $pJob.BtnRevertDone = $btnRevertDone
    $filePrepPanel.Children.Add($applyRow) | Out-Null
    $pJob.ApplyRow = $applyRow

    # Review status label â€” shown at bottom in Review mode after Keep/Revert
    $reviewStatusLabel = New-Object System.Windows.Controls.TextBlock
    $reviewStatusLabel.FontSize = 16; $reviewStatusLabel.FontWeight = [System.Windows.FontWeights]::Bold
    $reviewStatusLabel.HorizontalAlignment = "Center"; $reviewStatusLabel.TextAlignment = "Center"
    $reviewStatusLabel.Margin = New-Object System.Windows.Thickness(0,10,0,0)
    $reviewStatusLabel.Visibility = "Collapsed"
    $rightStack.Children.Add($reviewStatusLabel) | Out-Null
    $pJob.ReviewStatusLabel = $reviewStatusLabel

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # EDITING tab content â€” sub-tab strip for file-edit operations.
    # Add new sub-tab buttons/panels here as more editing tools are introduced.
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    # â”€â”€ Per-card editing queue row â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $edQueueRow = New-Object System.Windows.Controls.Border
    $edQueueRow.Background = Get-WpfColor "#1A1C24"
    $edQueueRow.BorderBrush = Get-WpfColor "#2A2C38"; $edQueueRow.BorderThickness = New-Object System.Windows.Thickness(0,0,0,1)
    $edQueueRow.Padding = New-Object System.Windows.Thickness(0,6,0,6)
    $edQueueRow.Margin = New-Object System.Windows.Thickness(0,4,0,4)
    $edQRowStack = New-Object System.Windows.Controls.StackPanel; $edQRowStack.Orientation = "Horizontal"

    $chkEdReNest = New-Object System.Windows.Controls.CheckBox; $chkEdReNest.Content = "Re-Nest"
    $chkEdReNest.IsChecked = $true; $chkEdReNest.Foreground = Get-WpfColor "#C8CFDD"
    $chkEdReNest.VerticalAlignment = "Center"; $chkEdReNest.Margin = New-Object System.Windows.Thickness(0,0,16,0)

    $btnEdQueue = New-Object System.Windows.Controls.Button; $btnEdQueue.Content = "Queue"
    $btnEdQueue.Height = 26; $btnEdQueue.Width = 70; $btnEdQueue.FontSize = 11
    $btnEdQueue.FontWeight = [System.Windows.FontWeights]::Bold
    $btnEdQueue.Background = Get-WpfColor "#3A5080"; $btnEdQueue.Foreground = Get-WpfColor "#FFFFFF"
    $btnEdQueue.BorderThickness = 0; $btnEdQueue.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnEdQueue.Margin = New-Object System.Windows.Thickness(0,0,10,0)

    $lblEdQueueStatus = New-Object System.Windows.Controls.TextBlock
    $lblEdQueueStatus.FontSize = 11; $lblEdQueueStatus.VerticalAlignment = "Center"
    $lblEdQueueStatus.Foreground = Get-WpfColor "#555868"

    $edQRowStack.Children.Add($chkEdReNest)     | Out-Null
    $edQRowStack.Children.Add($btnEdQueue)       | Out-Null
    $edQRowStack.Children.Add($lblEdQueueStatus) | Out-Null
    $edQueueRow.Child = $edQRowStack
    $editingPanel.Children.Add($edQueueRow) | Out-Null

    $pJob.ChkEdReNest      = $chkEdReNest
    $pJob.BtnEdQueue       = $btnEdQueue
    $pJob.LblEdQueueStatus = $lblEdQueueStatus

    # Sub-tab strip inside the Editing panel
    $feTabStrip = New-Object System.Windows.Controls.StackPanel
    $feTabStrip.Orientation = "Horizontal"
    $feTabStrip.Margin = New-Object System.Windows.Thickness(0,4,0,0)
    $editingPanel.Children.Add($feTabStrip) | Out-Null

    # Content area (sub-panels swap in/out here as sub-tabs are selected)
    $feContent = New-Object System.Windows.Controls.Border
    $feContent.Background = Get-WpfColor "#1E1F27"
    $feContent.CornerRadius = New-Object System.Windows.CornerRadius(0,4,4,4)
    $feContent.Padding = New-Object System.Windows.Thickness(12,10,12,12)
    $feContent.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $editingPanel.Children.Add($feContent) | Out-Null

    # â”€â”€ Re-Nest sub-tab button â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $feTabRenest = New-Object System.Windows.Controls.Button
    $feTabRenest.Content = "Re-Nest"; $feTabRenest.FontSize = 11
    $feTabRenest.FontWeight = [System.Windows.FontWeights]::Bold
    $feTabRenest.Background = Get-WpfColor "#1E1F27"   # matches content bg = "active"
    $feTabRenest.Foreground = Get-WpfColor "#C8CFDD"
    $feTabRenest.BorderBrush = Get-WpfColor "#2A2C38"; $feTabRenest.BorderThickness = New-Object System.Windows.Thickness(1,1,1,0)
    $feTabRenest.Padding = New-Object System.Windows.Thickness(14,5,14,6)
    $feTabRenest.Margin = New-Object System.Windows.Thickness(0,0,2,0)
    $feTabRenest.Cursor = [System.Windows.Input.Cursors]::Hand; $feTabRenest.IsEnabled = $false
    $feTabStrip.Children.Add($feTabRenest) | Out-Null

    # â”€â”€ Re-Nest panel content â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $panelRenest = New-Object System.Windows.Controls.StackPanel

    # â”€â”€ Info grid (Final / Source / Output) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $infoGrid = New-Object System.Windows.Controls.Grid
    $igC1 = New-Object System.Windows.Controls.ColumnDefinition; $igC1.Width = "Auto"
    $igC2 = New-Object System.Windows.Controls.ColumnDefinition; $igC2.Width = "*"
    $infoGrid.ColumnDefinitions.Add($igC1) | Out-Null; $infoGrid.ColumnDefinitions.Add($igC2) | Out-Null
    $infoGrid.Margin = New-Object System.Windows.Thickness(0,0,0,10)
    foreach ($rIdx in 0..2) {
        $rd = New-Object System.Windows.Controls.RowDefinition; $rd.Height = "Auto"
        $infoGrid.RowDefinitions.Add($rd) | Out-Null
    }
    $infoLabelTxts = @("Final:","Source:","Output:")
    $infoTbFinal = $null; $infoTbSource = $null; $infoTbOutput = $null
    for ($ri = 0; $ri -lt 3; $ri++) {
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $infoLabelTxts[$ri]; $lbl.Foreground = Get-WpfColor "#888A9A"
        $lbl.FontSize = 11; $lbl.Margin = New-Object System.Windows.Thickness(0,1,8,1)
        [System.Windows.Controls.Grid]::SetRow($lbl, $ri); [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
        $infoGrid.Children.Add($lbl) | Out-Null
        $val = New-Object System.Windows.Controls.TextBlock
        $val.FontSize = 11; $val.Foreground = Get-WpfColor "#7AAABB"
        $val.TextWrapping = "Wrap"; $val.Margin = New-Object System.Windows.Thickness(0,1,0,1)
        [System.Windows.Controls.Grid]::SetRow($val, $ri); [System.Windows.Controls.Grid]::SetColumn($val, 1)
        $infoGrid.Children.Add($val) | Out-Null
        if ($ri -eq 0) { $infoTbFinal = $val } elseif ($ri -eq 1) { $infoTbSource = $val } else { $infoTbOutput = $val }
    }
    $panelRenest.Children.Add($infoGrid) | Out-Null

    # â”€â”€ Run button row â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # WrapPanel so the orientation buttons (Rotate 90 / Save Orientation) wrap
    # onto a new line instead of being clipped when the right panel is narrow.
    $renestRow = New-Object System.Windows.Controls.WrapPanel
    $renestRow.Orientation = "Horizontal"
    $renestRow.Margin = New-Object System.Windows.Thickness(0,0,0,8)

    $btnRunRenest = New-Object System.Windows.Controls.Button
    $btnRunRenest.Content = "Run Re-Nest"; $btnRunRenest.Height = 32; $btnRunRenest.Width = 115
    $btnRunRenest.FontWeight = [System.Windows.FontWeights]::Bold; $btnRunRenest.FontSize = 12
    $btnRunRenest.BorderThickness = 0; $btnRunRenest.Cursor = [System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnRunRenest, 0)
    $renestRow.Children.Add($btnRunRenest) | Out-Null

    $lblRenestStatus = New-Object System.Windows.Controls.TextBlock
    $lblRenestStatus.FontSize = 11; $lblRenestStatus.VerticalAlignment = "Center"
    $lblRenestStatus.Margin = New-Object System.Windows.Thickness(10,0,0,0); $lblRenestStatus.TextWrapping = "Wrap"
    $lblRenestStatus.MaxWidth = 400
    $renestRow.Children.Add($lblRenestStatus) | Out-Null
    $panelRenest.Children.Add($renestRow) | Out-Null

    # Off-tree placeholder so Update-AxisPickerPreview can still set .Text safely.
    $lblAxisRef = New-Object System.Windows.Controls.TextBlock

    # Rotate 90 button - always visible in the right-side Re-Nest panel
    # (renestRow). The orientation state loads automatically with the card,
    # Rotate 90 adjusts the previewed spin, and Run Re-Nest hands the
    # on-screen state to the worker via a temp file - nothing to save.
    $btnAxisRotate = New-Object System.Windows.Controls.Button
    $btnAxisRotate.Content = "Rotate 90"; $btnAxisRotate.Height = 32; $btnAxisRotate.Width = 90
    $btnAxisRotate.FontSize = 11; $btnAxisRotate.Background = Get-WpfColor "#2A2C38"
    $btnAxisRotate.Foreground = Get-WpfColor "#C8CFDD"; $btnAxisRotate.BorderThickness = 0
    $btnAxisRotate.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnAxisRotate.Margin = New-Object System.Windows.Thickness(8,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($btnAxisRotate, 1)
    $renestRow.Children.Add($btnAxisRotate) | Out-Null

    # â”€â”€ Review panel (hidden until success) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $borderReview = New-Object System.Windows.Controls.Border
    $borderReview.Background = Get-WpfColor "#13151C"
    $borderReview.BorderBrush = Get-WpfColor "#3A3C4A"; $borderReview.BorderThickness = 1
    $borderReview.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $borderReview.Visibility = "Collapsed"
    $panelReviewInner = New-Object System.Windows.Controls.StackPanel
    $panelReviewInner.Margin = New-Object System.Windows.Thickness(8)

    $reviewBtnRow = New-Object System.Windows.Controls.StackPanel
    $reviewBtnRow.Orientation = "Horizontal"
    $btnReplaceSource = New-Object System.Windows.Controls.Button
    $btnReplaceSource.Content = "Confirm and Replace"; $btnReplaceSource.Height = 30; $btnReplaceSource.Width = 160
    $btnReplaceSource.FontWeight = [System.Windows.FontWeights]::Bold; $btnReplaceSource.FontSize = 11
    $btnReplaceSource.Background = Get-WpfColor "#2E5A42"; $btnReplaceSource.Foreground = Get-WpfColor "#FFFFFF"
    $btnReplaceSource.BorderThickness = 0; $btnReplaceSource.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnReplaceSource.Margin = New-Object System.Windows.Thickness(0,0,8,0)
    $reviewBtnRow.Children.Add($btnReplaceSource) | Out-Null

    $btnOpenRenest = New-Object System.Windows.Controls.Button
    $btnOpenRenest.Content = "Open in Bambu"; $btnOpenRenest.Height = 30; $btnOpenRenest.Width = 110
    $btnOpenRenest.FontSize = 11; $btnOpenRenest.Background = Get-WpfColor "#2A2C38"
    $btnOpenRenest.Foreground = Get-WpfColor "#7AAABB"; $btnOpenRenest.BorderThickness = 0
    $btnOpenRenest.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnOpenRenest.Margin = New-Object System.Windows.Thickness(0,0,8,0)
    $btnOpenRenest.ToolTip = "Open the re-nested .3mf in Bambu Studio without replacing the source file"
    $reviewBtnRow.Children.Add($btnOpenRenest) | Out-Null

    $btnDiscardRenest = New-Object System.Windows.Controls.Button
    $btnDiscardRenest.Content = "Discard"; $btnDiscardRenest.Height = 30; $btnDiscardRenest.Width = 80
    $btnDiscardRenest.FontSize = 11; $btnDiscardRenest.Background = Get-WpfColor "#3A2020"
    $btnDiscardRenest.Foreground = Get-WpfColor "#CC8888"; $btnDiscardRenest.BorderThickness = 0
    $btnDiscardRenest.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnDiscardRenest.Margin = New-Object System.Windows.Thickness(0,0,8,0)
    $reviewBtnRow.Children.Add($btnDiscardRenest) | Out-Null

    $btnOpenDebug = New-Object System.Windows.Controls.Button
    $btnOpenDebug.Content = "Open Debug"; $btnOpenDebug.Height = 30; $btnOpenDebug.Width = 90
    $btnOpenDebug.FontSize = 11; $btnOpenDebug.Background = Get-WpfColor "#2A2C38"
    $btnOpenDebug.Foreground = Get-WpfColor "#7AAABB"; $btnOpenDebug.BorderThickness = 0
    $btnOpenDebug.Cursor = [System.Windows.Input.Cursors]::Hand
    $reviewBtnRow.Children.Add($btnOpenDebug) | Out-Null

    $btnSaveDebug = New-Object System.Windows.Controls.Button
    $btnSaveDebug.Content = "Save Debug"; $btnSaveDebug.Height = 30; $btnSaveDebug.Width = 85
    $btnSaveDebug.FontSize = 11; $btnSaveDebug.Background = Get-WpfColor "#2A2C38"
    $btnSaveDebug.Foreground = Get-WpfColor "#888A9A"; $btnSaveDebug.BorderThickness = 0
    $btnSaveDebug.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnSaveDebug.Margin = New-Object System.Windows.Thickness(6,0,0,0)
    $reviewBtnRow.Children.Add($btnSaveDebug) | Out-Null

    $panelReviewInner.Children.Add($reviewBtnRow) | Out-Null
    $borderReview.Child = $panelReviewInner
    $panelRenest.Children.Add($borderReview) | Out-Null

    $feContent.Child = $panelRenest

    # â”€â”€ Detect Final.3mf and sibling Source at build time â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $feRenestFinal  = Get-ChildItem -Path $parentPath -Filter "*_Final.3mf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    $feRenestSource = $null
    if ($null -ne $feRenestFinal) {
        $reStem = $feRenestFinal.BaseName -replace '(?i)_Final$', ''
        $feRenestSource = Get-ChildItem -Path $parentPath -Filter "${reStem}_Nest.3mf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -eq $feRenestSource) {
            $feRenestSource = Get-ChildItem -Path $parentPath -Filter "${reStem}_Full.3mf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        }
    }

    # Populate info rows
    $infoTbFinal.Text  = if ($null -ne $feRenestFinal)  { $feRenestFinal.Name }  else { "(not found)" }
    $infoTbFinal.Foreground  = Get-WpfColor $(if ($null -ne $feRenestFinal)  { "#7AAABB" } else { "#D95F5F" })
    $infoTbSource.Text = if ($null -ne $feRenestSource) { $feRenestSource.Name } else { "(not found)" }
    $infoTbSource.Foreground = Get-WpfColor $(if ($null -ne $feRenestSource) { "#7AAABB" } else { "#D95F5F" })
    $outputName = if ($null -ne $feRenestFinal) { ($feRenestFinal.BaseName -replace '(?i)_Final$','') + "_Renest.3mf" } else { "(n/a)" }
    $infoTbOutput.Text = $outputName; $infoTbOutput.Foreground = Get-WpfColor "#888A9A"

    if ($null -eq $feRenestFinal) {
        $btnRunRenest.IsEnabled = $false
        $btnRunRenest.Background = Get-WpfColor "#2A2C38"; $btnRunRenest.Foreground = Get-WpfColor "#555868"
        $lblRenestStatus.Text = "No *_Final.3mf found"; $lblRenestStatus.Foreground = Get-WpfColor "#444658"
        $btnAxisRotate.IsEnabled = $false
        $btnAxisRotate.Background = Get-WpfColor "#2A2C38"; $btnAxisRotate.Foreground = Get-WpfColor "#555868"
    } else {
        $btnRunRenest.IsEnabled = $true
        $btnRunRenest.Background = Get-WpfColor "#2E5A42"; $btnRunRenest.Foreground = Get-WpfColor "#FFFFFF"
        $lblRenestStatus.Text = "Ready"; $lblRenestStatus.Foreground = Get-WpfColor "#666878"
    }

    # Shared state for all three buttons
    $renestTag = @{
        P                 = $pJob
        WorkerPath        = (Join-Path $scriptDir "..\workers\RenestFromFinal_worker.ps1")
        FinalPath         = if ($null -ne $feRenestFinal)  { $feRenestFinal.FullName }  else { "" }
        SourcePath        = if ($null -ne $feRenestSource) { $feRenestSource.FullName } else { "" }
        RenestPath        = ""
        BtnRun            = $btnRunRenest
        LblStatus         = $lblRenestStatus
        BorderReview      = $borderReview
        LogOut            = ""
        LogErr            = ""
        Proc              = $null
        Timer             = $null
        DebugTempPath     = ""
        OvSrcPath         = $null
        OvFinPath         = $null
        AxisR             = $null
        FinalAxisR        = $null
        AxisScale         = 2.0
        PendingRotCorrection = $null
        PendingSourceR    = $null
        SrcRefRotNorm     = $null
        SrcRefYawDeg      = 0.0
        SrcNestAngleDeg   = 0.0
        AxisYawSteps      = 0
        LblAxisRef        = $lblAxisRef
        BtnAxisRotate     = $btnAxisRotate
        RotCorrTempPath   = ""
    }
    $btnRunRenest.Tag        = $renestTag
    $btnReplaceSource.Tag    = $renestTag
    $btnOpenRenest.Tag       = $renestTag
    $btnDiscardRenest.Tag    = $renestTag
    $btnOpenDebug.Tag        = $renestTag
    $btnSaveDebug.Tag        = $renestTag
    $btnAxisRotate.Tag  = $renestTag

    $btnRunRenest.Add_Click({
        $t = $this.Tag
        if (-not (Test-Path -LiteralPath $t.WorkerPath)) {
            $t.LblStatus.Text = "Worker script not found"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        if ([string]::IsNullOrEmpty($t.FinalPath) -or -not (Test-Path -LiteralPath $t.FinalPath)) {
            $t.LblStatus.Text = "No *_Final.3mf found"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        # Reset UI
        $t.BorderReview.Visibility = "Collapsed"
        $t.BtnRun.IsEnabled = $false; $t.BtnRun.Content = "Running..."
        $t.BtnRun.Background = Get-WpfColor "#333333"; $t.BtnRun.Foreground = Get-WpfColor "#888888"
        $t.LblStatus.Text = "Re-nesting..."; $t.LblStatus.Foreground = Get-WpfColor "#F0A030"
        # Temp log files
        $tmpBase = Join-Path $env:TEMP ("renest_" + [System.Guid]::NewGuid().ToString("N"))
        $t.LogOut = $tmpBase + "_out.txt"; $t.LogErr = $tmpBase + "_err.txt"
        [System.IO.File]::WriteAllText($t.LogOut, ""); [System.IO.File]::WriteAllText($t.LogErr, "")
        $t.RenestTempOut = Join-Path $env:TEMP ("renest_" + [System.Guid]::NewGuid().ToString("N") + ".3mf")
        $psArgs = @("-NoProfile","-ExecutionPolicy","Bypass","-File","`"$($t.WorkerPath)`"","-FinalPath","`"$($t.FinalPath)`"","-OutputPath","`"$($t.RenestTempOut)`"","-NoConfirm")
        # Hand the on-screen orientation state to the worker via a TEMP json -
        # nothing is written into the design folder. refSpinDeg is the exact
        # total world-Z spin currently previewed (nest angle + Rotate-90
        # steps), applied verbatim by the worker, so what the user sees is
        # literally what runs.
        $t.RotCorrTempPath = ""
        if ($null -ne $t.SrcRefRotNorm) {
            try {
                $rcPath = $tmpBase + "_rotcorr.json"
                ([ordered]@{
                    srcRefRot    = $t.SrcRefRotNorm
                    srcRefYawDeg = $t.SrcRefYawDeg
                    nestAngleDeg = $t.SrcNestAngleDeg
                    extraYawDeg  = ($t.AxisYawSteps * 90)
                    refSpinDeg   = ($t.SrcNestAngleDeg + $t.AxisYawSteps * 90)
                } | ConvertTo-Json -Depth 5) | Set-Content -Path $rcPath -Encoding UTF8
                $t.RotCorrTempPath = $rcPath
                $psArgs += @("-RotCorrectionPath","`"$rcPath`"")
            } catch {}
        }
        $t.Proc = Start-Process -FilePath "powershell.exe" -ArgumentList $psArgs `
            -NoNewWindow -RedirectStandardOutput $t.LogOut -RedirectStandardError $t.LogErr -PassThru
        $script:_RenestActive = $t
        $renestTimer = New-Object System.Windows.Threading.DispatcherTimer
        $renestTimer.Interval = [TimeSpan]::FromMilliseconds(250)
        $t.Timer = $renestTimer
        $renestTimer.Add_Tick($script:_RenestTickSB)
        $renestTimer.Start()
    })

    $btnAxisRotate.Add_Click({
        $t = $this.Tag
        if ($null -eq $t.SrcRefRotNorm) {
            $t.LblStatus.Text = "Orientation not loaded (need Final + Nest Source)"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        $t.AxisYawSteps = ($t.AxisYawSteps + 1) % 4
        Update-AxisPickerPreview $t
        Update-FinalCurrentAxesOverlay $t
    })

    $btnReplaceSource.Add_Click({
        $t = $this.Tag
        if ([string]::IsNullOrEmpty($t.RenestPath) -or -not (Test-Path -LiteralPath $t.RenestPath)) {
            $t.LblStatus.Text = "Renest file not found"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        if ([string]::IsNullOrEmpty($t.SourcePath)) {
            $t.LblStatus.Text = "Source path unknown"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        try {
            # 1. Replace the source file with the renested file first
            Move-Item -LiteralPath $t.RenestPath -Destination $t.SourcePath -Force
            $t.BorderReview.Visibility = "Collapsed"
            $t.RenestPath = ""

            # 2. Auto-revert merge if applicable
            $srcDir  = Split-Path $t.SourcePath -Parent
            $srcBase = [System.IO.Path]::GetFileNameWithoutExtension($t.SourcePath)

            $mergeReverted = $false

            if ($srcBase -imatch '_Nest$') {
                # â”€â”€ Merged folder: source was Nest.3mf â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                # feRenestSource resolves to Nest.3mf when the folder is merged because
                # it carries the individual-object transforms. The renested file now
                # lives at Nest.3mf; we need to delete the stale Full.3mf and rename
                # Nest â†’ Full to restore a clean pre-merge state.
                $corePrefix    = $srcBase -replace '(?i)_Nest$', ''  # e.g. X1C_Arthropleura
                $staleFullPath = Join-Path $srcDir ($corePrefix + '_Full.3mf')
                $staleFiles    = @(
                    $staleFullPath,
                    (Join-Path $srcDir ($corePrefix + '_Final.gcode.3mf')),
                    (Join-Path $srcDir ($corePrefix + '_Full.gcode.3mf')),
                    (Join-Path $srcDir ($corePrefix + '_Data.tsv'))
                )
                foreach ($s in $staleFiles) {
                    Remove-StaleFile $s
                }
                # Rename the renested Nest.3mf â†’ Full.3mf
                Rename-Item -LiteralPath $t.SourcePath -NewName ($corePrefix + '_Full.3mf') -Force
                $mergeReverted = $true

            } elseif ($srcBase -imatch '_Full$') {
                # â”€â”€ Unmerged folder: source was Full.3mf â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                # A Nest.3mf shouldn't normally be here but handle it in case.
                $corePrefix = $srcBase -replace '(?i)_Full$', ''
                $nestFile   = Join-Path $srcDir ($corePrefix + '_Nest.3mf')
                if (Test-Path -LiteralPath $nestFile) {
                    $staleFiles = @(
                        $nestFile,
                        (Join-Path $srcDir ($corePrefix + '_Final.gcode.3mf')),
                        (Join-Path $srcDir ($srcBase    + '.gcode.3mf')),
                        (Join-Path $srcDir ($corePrefix + '_Data.tsv'))
                    )
                    foreach ($s in $staleFiles) {
                        if (Test-Path -LiteralPath $s) { Remove-Item -LiteralPath $s -Force -ErrorAction SilentlyContinue }
                    }
                    $mergeReverted = $true
                }
            }

            if ($mergeReverted) {
                # Update the card UI to reflect that the merge has been undone
                if ($null -ne $t.P.ChkMerge) {
                    $t.P.ChkMerge.IsEnabled  = $true
                    $t.P.ChkMerge.Foreground = Get-WpfColor "#FFFFFF"
                    $t.P.ChkMerge.ToolTip    = $null
                }
                if ($null -ne $t.P.BtnRevertMerge) {
                    $t.P.BtnRevertMerge.IsEnabled  = $false
                    $t.P.BtnRevertMerge.Background = Get-WpfColor "#3A3A3A"
                    $t.P.BtnRevertMerge.Foreground = Get-WpfColor "#666666"
                    $t.P.BtnRevertMerge.ToolTip    = "No merged file detected"
                }
                if ($null -ne $t.P.MergeBanner) { $t.P.MergeBanner.Visibility = "Collapsed" }
                $t.LblStatus.Text = "Source replaced + merge reverted."; $t.LblStatus.Foreground = Get-WpfColor "#4CAF72"
            } else {
                $t.LblStatus.Text = "Source replaced."; $t.LblStatus.Foreground = Get-WpfColor "#4CAF72"
            }
        } catch {
            $t.LblStatus.Text = "Replace failed: $($_.Exception.Message)"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"
        }
    })

    $btnOpenRenest.Add_Click({
        $t = $this.Tag
        if ([string]::IsNullOrEmpty($t.RenestPath) -or -not (Test-Path -LiteralPath $t.RenestPath)) {
            $t.LblStatus.Text = "Renest file not found"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        $bambuPath = "C:\Program Files\Bambu Studio\bambu-studio.exe"
        if (-not (Test-Path -LiteralPath $bambuPath)) {
            $t.LblStatus.Text = "Bambu Studio not found: $bambuPath"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"; return
        }
        try {
            Start-Process -FilePath $bambuPath -ArgumentList "`"$($t.RenestPath)`""
            $t.LblStatus.Text = "Opened re-nested file in Bambu Studio for review."
            $t.LblStatus.Foreground = Get-WpfColor "#888A9A"
        } catch {
            $t.LblStatus.Text = "Failed to open: $($_.Exception.Message)"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"
        }
    })

    $btnDiscardRenest.Add_Click({
        $t = $this.Tag
        if (-not [string]::IsNullOrEmpty($t.RenestPath) -and (Test-Path -LiteralPath $t.RenestPath)) {
            try { Remove-Item -LiteralPath $t.RenestPath -Force -ErrorAction SilentlyContinue } catch {}
        }
        # Restore right overlay to the original Final image (revert label back to "Final")
        try {
            if ($null -ne $t.P.RnLblRight) { $t.P.RnLblRight.Text = "Final" }
            if (-not [string]::IsNullOrEmpty($t.FinalPath) -and (Test-Path -LiteralPath $t.FinalPath)) {
                $discardTmpDir = Join-Path $env:TEMP ("discard_imgs_" + [System.Guid]::NewGuid().ToString("N"))
                New-Item -ItemType Directory -Path $discardTmpDir -Force | Out-Null
                $discardFinPath = Extract-3mfPickImage $t.FinalPath $discardTmpDir "dfin"
                if ($null -ne $discardFinPath -and (Test-Path $discardFinPath)) {
                    $t.P.RnImgRight.Source = Load-WpfImage $discardFinPath
                }
            }
            # Re-apply the orientation spin + axes to the restored Final preview
            Update-FinalCurrentAxesOverlay $t
        } catch {}
        $t.BorderReview.Visibility = "Collapsed"
        $t.LblStatus.Text = "Discarded."; $t.LblStatus.Foreground = Get-WpfColor "#888A9A"
        $t.RenestPath = ""
    })

    $btnOpenDebug.Add_Click({
        $t = $this.Tag
        $dbPath = if ($null -ne $t.DebugTempPath) { $t.DebugTempPath } else { "" }
        if ([string]::IsNullOrEmpty($dbPath) -or -not (Test-Path -LiteralPath $dbPath)) {
            $t.LblStatus.Text = "No debug file available."; $t.LblStatus.Foreground = Get-WpfColor "#888A9A"
            return
        }
        try { Start-Process "notepad.exe" -ArgumentList $dbPath } catch {}
    })

    $btnSaveDebug.Add_Click({
        $t = $this.Tag
        $dbPath = if ($null -ne $t.DebugTempPath) { $t.DebugTempPath } else { "" }
        if ([string]::IsNullOrEmpty($dbPath) -or -not (Test-Path -LiteralPath $dbPath)) {
            $t.LblStatus.Text = "No debug file to save."; $t.LblStatus.Foreground = Get-WpfColor "#888A9A"
            return
        }
        $stemR2  = [System.IO.Path]::GetFileNameWithoutExtension($t.FinalPath) -replace '(?i)_Final$', ''
        $outDir2 = [System.IO.Path]::GetDirectoryName($t.FinalPath)
        $savePath = Join-Path $outDir2 ($stemR2 + "_Renest_debug.txt")
        try {
            Copy-Item -LiteralPath $dbPath -Destination $savePath -Force
            $t.LblStatus.Text = "Debug saved: $([System.IO.Path]::GetFileName($savePath))"
            $t.LblStatus.Foreground = Get-WpfColor "#4CAF72"
        } catch {
            $t.LblStatus.Text = "Save failed: $($_.Exception.Message)"; $t.LblStatus.Foreground = Get-WpfColor "#D95F5F"
        }
    })

    $pJob.BtnRunRenest = $btnRunRenest; $pJob.RenestStatusLbl = $lblRenestStatus

    # â”€â”€ Per-card Queue button handler â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $btnEdQueue.Tag = @{ P = $pJob; G = $gpJob }
    $btnEdQueue.Add_Click({
        $t = $this.Tag
        Enqueue-EditJob $t.P $t.G
        if ($null -eq $script:editActiveJob) { Start-NextEditJob }
    }.GetNewClosure())

    # â”€â”€ Edit-mode overlays â€” always visible in Editing mode, show top_1.png from source/final â”€â”€
    # Col 0 of leftGrid = "Nest Source" top_1.png  (static â€” never changes)
    # Col 1 of leftGrid = "Final" top_1.png        (updated to "Re-Nest" after renesting)
    $ovInitVis = if ($script:GlobalMode -eq "Editing") { "Visible" } else { "Collapsed" }
    foreach ($ovSide in @(0, 1)) {
        $ovBorder = New-Object System.Windows.Controls.Border
        $ovBorder.Background = Get-WpfColor "#0A0B0F"; $ovBorder.Visibility = $ovInitVis
        $ovBorder.VerticalAlignment = "Stretch"; $ovBorder.HorizontalAlignment = "Stretch"
        [System.Windows.Controls.Grid]::SetColumn($ovBorder, $ovSide)
        $ovLbl = New-Object System.Windows.Controls.TextBlock
        $ovLbl.Text = if ($ovSide -eq 0) { "Nest Source" } else { "Final" }
        $ovLbl.Foreground = Get-WpfColor "#7AAABB"; $ovLbl.FontSize = 12
        $ovLbl.FontWeight = [System.Windows.FontWeights]::Bold
        $ovLbl.HorizontalAlignment = "Center"; $ovLbl.Margin = New-Object System.Windows.Thickness(0,0,0,8)
        $ovImg = New-Object System.Windows.Controls.Image
        $ovImg.Stretch = "Uniform"; $ovImg.MaxHeight = 420
        $ovImg.Cursor = [System.Windows.Input.Cursors]::Hand
        $ovImg.Add_MouseLeftButtonDown({
            if ($_.ClickCount -ge 2 -and $null -ne $this.Source) {
                $viewer = New-Object System.Windows.Window
                $viewer.Title = if ($this.Tag -eq 0) { "Nest Source (top_1.png)" } else { "Final / Re-Nest (top_1.png)" }
                $viewer.Background = Get-WpfColor "#0D0E10"
                $viewer.SizeToContent = "WidthAndHeight"; $viewer.WindowStartupLocation = "CenterScreen"
                $viewer.ResizeMode = "CanResizeWithGrip"
                $imgView = New-Object System.Windows.Controls.Image
                $imgView.Source = $this.Source; $imgView.MaxWidth = 900; $imgView.MaxHeight = 900; $imgView.Stretch = "Uniform"
                $imgView.Margin = New-Object System.Windows.Thickness(10)
                $viewer.Content = $imgView
                $viewer.ShowDialog() | Out-Null
            }
        })
        $ovImg.Tag = $ovSide

        $ovStack = New-Object System.Windows.Controls.StackPanel
        $ovStack.VerticalAlignment = "Center"; $ovStack.HorizontalAlignment = "Center"
        $ovStack.Children.Add($ovLbl) | Out-Null
        $ovStack.Children.Add($ovImg)  | Out-Null
        $ovBorder.Child = $ovStack

        $leftGrid.Children.Add($ovBorder) | Out-Null
        if ($ovSide -eq 0) { $pJob.RnOvLeft  = $ovBorder; $pJob.RnImgLeft  = $ovImg; $pJob.RnLblLeft  = $ovLbl }
        else               { $pJob.RnOvRight = $ovBorder; $pJob.RnImgRight = $ovImg; $pJob.RnLblRight = $ovLbl }
    }

    # Pre-load top_1.png images from source and final 3MF files into the overlays
    try {
        $ovTmpDir = Join-Path $env:TEMP ("ov_imgs_" + [System.Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $ovTmpDir -Force | Out-Null
        $ovSrcPath = if ($null -ne $feRenestSource -and (Test-Path -LiteralPath $feRenestSource.FullName)) {
            Extract-3mfPickImage $feRenestSource.FullName $ovTmpDir "ovsrc"
        } else { $null }
        $ovFinPath = if ($null -ne $feRenestFinal -and (Test-Path -LiteralPath $feRenestFinal.FullName)) {
            Extract-3mfPickImage $feRenestFinal.FullName $ovTmpDir "ovfin"
        } else { $null }
        if ($null -ne $ovSrcPath -and (Test-Path $ovSrcPath)) { $pJob.RnImgLeft.Source  = Load-WpfImage $ovSrcPath }
        if ($null -ne $ovFinPath -and (Test-Path $ovFinPath)) { $pJob.RnImgRight.Source = Load-WpfImage $ovFinPath }
        $renestTag.OvSrcPath = $ovSrcPath
        $renestTag.OvFinPath = $ovFinPath
        # Load the orientation state immediately so the Final preview shows
        # the spin Re-Nest will apply (nest angle + any saved Rotate-90
        # offset) with the reference axes, before any user interaction.
        if (-not (Initialize-OrientationState $renestTag)) {
            Update-NestRefAxesOverlay $renestTag
            Update-FinalCurrentAxesOverlay $renestTag
        }
    } catch {}

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

    $btnColorOnly.Tag = @{ P = $pJob; G = $gpJob }
    $btnColorOnly.Add_Click({
        $t = $this.Tag
        Enqueue-ColorOnlyJob $t.P $t.G
    })

    $btnRenameOnly.Tag = @{ P = $pJob; G = $gpJob }
    $btnRenameOnly.Add_Click({
        $t = $this.Tag
        $result = [System.Windows.MessageBox]::Show(
            "Colors are unmatched. This will rename files only`n(merge, slice, extract, and image will be skipped).`n`nContinue?",
            "Rename Only",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
        $t.P.RenameOnlyBypass = $true
        Enqueue-PJob $t.P $t.G
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

    # â”€â”€ Detect printer prefix â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Peel leading qualifier tokens (printer prefix + tags) from the GP folder
    # name; also search anchor file stems, and walk up to the great-grandparent.
    $gpDetectedPrefix = ""; $gpDetectedTag = ""; $gpNameForTheme = $gpName
    if ($gpName -ne "(No Parent Folder)") {
        $gpTokens = [System.Collections.Generic.List[string]]($gpName -split '_' | Where-Object { $_ -ne '' })
        while ($gpTokens.Count -gt 1) {
            $head = $gpTokens[0]
            if ($script:PrinterPrefixes -icontains $head) {
                if ($gpDetectedPrefix -eq '') { $gpDetectedPrefix = $head }
                $gpTokens.RemoveAt(0)
            } elseif ($script:Tags -icontains $head) {
                if ($gpDetectedTag -eq '') { $gpDetectedTag = $head }
                $gpTokens.RemoveAt(0)
            } else { break }
        }
        $gpNameForTheme = $gpTokens -join '_'

        # Prefix fallback 1 â€“ anchor file stems
        if ($gpDetectedPrefix -eq "") {
            foreach ($pKey in $parentDict.Keys) {
                $afParts = (($parentDict[$pKey]).BaseName -replace '(?i)_(Full|Final|Nest)$','') -split '_'
                if ($afParts.Count -gt 0 -and $script:PrinterPrefixes -icontains $afParts[0]) {
                    $gpDetectedPrefix = $afParts[0]; break
                }
            }
        }
        # Prefix fallback 2 â€“ great-grandparent folder tokens
        if ($gpDetectedPrefix -eq "" -and $diGrand -and $diGrand.Parent -and $diGrand.Parent.Parent) {
            foreach ($tok in ($diGrand.Parent.Name -split '_' | Where-Object { $_ -ne '' })) {
                if ($script:PrinterPrefixes -icontains $tok) { $gpDetectedPrefix = $tok; break }
            }
        }

        # â”€â”€ Detect theme name â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Try each source in priority order; stop at the first recognised match.
        # 1. GP folder (already peeled above)
        $detectedTheme = Find-ThemeMatch $gpNameForTheme
        # 2. Anchor file stems
        if (-not $detectedTheme) {
            foreach ($pKey in $parentDict.Keys) {
                $stemTest = ($parentDict[$pKey]).BaseName -replace '(?i)_(Full|Final|Nest)$', ''
                $detectedTheme = Find-ThemeMatch $stemTest
                if ($detectedTheme) { break }
            }
        }
        # 3. Parent (design) folder names
        if (-not $detectedTheme) {
            foreach ($pKey in $parentDict.Keys) {
                $detectedTheme = Find-ThemeMatch (Split-Path $pKey -Leaf)
                if ($detectedTheme) { break }
            }
        }
        # 4. Great-grandparent folder name
        if (-not $detectedTheme -and $diGrand -and $diGrand.Parent -and $diGrand.Parent.Parent) {
            $detectedTheme = Find-ThemeMatch $diGrand.Parent.Name
        }

        if ($detectedTheme) { $gpNameForTheme = $detectedTheme }
    }

    $gpJob = @{ GpPath = $gpPath; DiGrand = $diGrand; Parents = New-Object System.Collections.ArrayList; CbPrefix = $null; CbTag = $null; GpRenameConfirmed = $false; ReviewMode = $false; HeaderGrid = $null; ThemeBar = $null; EditingThemeBar = $null; RenameGroup = $null; BtnThProcess = $null; BtnThRevert = $null; LblWigglitzAvg = $null; LblReadyCount = $null; BtnThPrintQReview = $null; BtnThSendProdReview = $null; BtnCombineGp = $null }
    $script:jobs.Add($gpJob) | Out-Null

    $container = New-Object System.Windows.Controls.Border
    $container.Background = Get-WpfColor "#1C1D23"; $container.BorderBrush = Get-WpfColor "#2A2C35"
    $container.BorderThickness = New-Object System.Windows.Thickness(1)
    $container.Margin = New-Object System.Windows.Thickness(0,0,0,20); $container.CornerRadius = New-Object System.Windows.CornerRadius(6)
    $gpJob.Container = $container

    $gpStack = New-Object System.Windows.Controls.StackPanel; $container.Child = $gpStack

    $headerGrid = New-Object System.Windows.Controls.WrapPanel
    $headerGrid.Background = Get-WpfColor "#2A2C35"
    $gpJob.HeaderGrid = $headerGrid

    $headerStack = New-Object System.Windows.Controls.StackPanel; $headerStack.Orientation = "Horizontal"
    $headerStack.Margin = New-Object System.Windows.Thickness(0,13,0,13)

    # Current folder name (far left)
    $lblCurrentName = Create-TextBlock $gpName "#CCCCCC" 13 "Bold"
    $lblCurrentName.VerticalAlignment = "Center"; $lblCurrentName.Margin = New-Object System.Windows.Thickness(15,0,20,0)
    $headerStack.Children.Add($lblCurrentName) | Out-Null

    # Renaming controls â€” hidden in Editing mode
    $renameGroup = New-Object System.Windows.Controls.StackPanel; $renameGroup.Orientation = "Horizontal"
    $renameGroup.VerticalAlignment = "Center"
    $renameGroup.Visibility = if ($script:GlobalMode -eq "Editing") { "Collapsed" } else { "Visible" }
    $headerStack.Children.Add($renameGroup) | Out-Null
    $gpJob.RenameGroup = $renameGroup

    # Printer prefix dropdown
    $lblPrefix = Create-TextBlock "Printer: " "#E8A135" 14 "Bold"
    $lblPrefix.Margin = New-Object System.Windows.Thickness(0,0,0,0); $renameGroup.Children.Add($lblPrefix) | Out-Null
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
    $renameGroup.Children.Add($cbPrefix) | Out-Null; $gpJob.CbPrefix = $cbPrefix

    # Grandparent tag dropdown
    $lblGpTag = Create-TextBlock "Tag: " "#E8A135" 14 "Bold"
    $lblGpTag.Margin = New-Object System.Windows.Thickness(0,0,0,0); $renameGroup.Children.Add($lblGpTag) | Out-Null
    $cbGpTag = New-Object System.Windows.Controls.ComboBox; $cbGpTag.Width = 80
    $cbGpTag.Background = Get-WpfColor "#2A2C35"; $cbGpTag.Foreground = Get-WpfColor "#FFFFFF"
    $cbGpTag.BorderBrush = Get-WpfColor "#5A78C4"; $cbGpTag.BorderThickness = New-Object System.Windows.Thickness(1)
    $cbGpTag.VerticalAlignment = "Center"; $cbGpTag.Margin = New-Object System.Windows.Thickness(5,0,20,0)
    $cbGpTag.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, [System.Windows.Controls.ToolBar]::ComboBoxStyleKey)
    $cbGpTag.Resources[[System.Windows.SystemColors]::WindowBrushKey]          = Get-WpfColor "#2A2C35"
    $cbGpTag.Resources[[System.Windows.SystemColors]::WindowTextBrushKey]      = Get-WpfColor "#FFFFFF"
    $cbGpTag.Resources[[System.Windows.SystemColors]::HighlightBrushKey]       = Get-WpfColor "#5A78C4"
    $cbGpTag.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey]   = Get-WpfColor "#FFFFFF"
    $cbGpTagItemStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
    $cbGpTagItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, (Get-WpfColor "#2A2C35"))))
    $cbGpTagItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, (Get-WpfColor "#FFFFFF"))))
    $cbGpTag.ItemContainerStyle = $cbGpTagItemStyle
    [void]$cbGpTag.Items.Add("(none)")
    foreach ($tag in $script:Tags) { [void]$cbGpTag.Items.Add($tag) }
    if ($gpDetectedTag -ne "" -and $script:Tags -icontains $gpDetectedTag) {
        $cbGpTag.SelectedItem = $gpDetectedTag
    } else { $cbGpTag.SelectedIndex = 0 }
    $renameGroup.Children.Add($cbGpTag) | Out-Null; $gpJob.CbTag = $cbGpTag

    # Grandparent theme label + dropdown
    $lblGP = Create-TextBlock "Theme: " "#E8A135" 14 "Bold"
    $lblGP.Margin = New-Object System.Windows.Thickness(0,0,0,0); $renameGroup.Children.Add($lblGP) | Out-Null

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
    $renameGroup.Children.Add($cbTheme) | Out-Null; $gpJob.TBTheme = $cbTheme

    $chkSkip = New-Object System.Windows.Controls.CheckBox; $chkSkip.Content = "Don't rename folder"
    $chkSkip.Foreground = Get-WpfColor "#FFFFFF"; $chkSkip.VerticalAlignment = "Center"
    $chkSkip.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    $renameGroup.Children.Add($chkSkip) | Out-Null; $gpJob.ChkSkip = $chkSkip

    # Live preview of the full grandparent folder name (Prefix_Theme or just Theme)
    $lblGpPreview = Create-TextBlock "" "#6B9FD4" 14 "Bold"
    $lblGpPreview.VerticalAlignment = "Center"; $lblGpPreview.Margin = New-Object System.Windows.Thickness(20,0,0,0)
    $initGpTag = if ($gpDetectedTag) { $gpDetectedTag } else { "Standard" }
    $initGpPreview = ((@($gpDetectedPrefix, $initGpTag, $gpNameForTheme) | Where-Object { $_ }) -join '_')
    $lblGpPreview.Text = if ($initGpPreview) { [char]0x2192 + " $initGpPreview" } else { "" }
    $renameGroup.Children.Add($lblGpPreview) | Out-Null; $gpJob.LblGpPreview = $lblGpPreview

    # Plate/file count - lives directly on the header bar so it stays visible
    # even when the rename controls (RenameGroup) are hidden in Review mode.
    $lblFileCount = Create-TextBlock "" "#888888" 11 "Normal"
    $lblFileCount.VerticalAlignment = "Center"; $lblFileCount.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    $headerStack.Children.Add($lblFileCount) | Out-Null; $gpJob.LblFileCount = $lblFileCount

    # Review-mode only: theme-wide Wigglitz/Day average
    $lblWigglitzAvg = Create-TextBlock "" "#4CAF72" 13 "Bold"
    $lblWigglitzAvg.VerticalAlignment = "Center"; $lblWigglitzAvg.Margin = New-Object System.Windows.Thickness(20,0,0,0)
    $lblWigglitzAvg.Visibility = "Collapsed"
    $headerStack.Children.Add($lblWigglitzAvg) | Out-Null; $gpJob.LblWigglitzAvg = $lblWigglitzAvg

    # Review-mode only: how many designs in this theme pass the production-ready checklist
    $lblReadyCount = Create-TextBlock "" "#4CAF72" 13 "Bold"
    $lblReadyCount.VerticalAlignment = "Center"; $lblReadyCount.Margin = New-Object System.Windows.Thickness(20,0,0,0)
    $lblReadyCount.Visibility = "Collapsed"
    $headerStack.Children.Add($lblReadyCount) | Out-Null; $gpJob.LblReadyCount = $lblReadyCount

    $headerGrid.Children.Add($headerStack) | Out-Null

    $gpRightBtnStack = New-Object System.Windows.Controls.StackPanel
    $gpRightBtnStack.Orientation = "Horizontal"; $gpRightBtnStack.HorizontalAlignment = "Right"
    $gpRightBtnStack.VerticalAlignment = "Center"; $gpRightBtnStack.Margin = New-Object System.Windows.Thickness(0,13,15,13)

    # (Review Mode is now a global top-bar button â€” no per-group toggle needed)

    # Theme-wide Review actions - only shown while this group is in Review mode
    $btnThPrintQReview = New-Object System.Windows.Controls.Button
    $btnThPrintQReview.Content = "Printing Queue (Theme)"; $btnThPrintQReview.Height = 30; $btnThPrintQReview.Width = 160
    $btnThPrintQReview.FontWeight = [System.Windows.FontWeights]::Bold; $btnThPrintQReview.FontSize = 11
    $btnThPrintQReview.Background = Get-WpfColor "#3A5080"; $btnThPrintQReview.Foreground = Get-WpfColor "#FFFFFF"
    $btnThPrintQReview.BorderThickness = 0; $btnThPrintQReview.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnThPrintQReview.Margin = New-Object System.Windows.Thickness(0,0,10,0)
    $btnThPrintQReview.ToolTip = "Copies every design's Full.gcode.3mf in this theme to the Printing Queue folder with today's date"
    $btnThPrintQReview.Visibility = "Collapsed"
    $gpRightBtnStack.Children.Add($btnThPrintQReview) | Out-Null
    $gpJob.BtnThPrintQReview = $btnThPrintQReview

    $btnThSendProdReview = New-Object System.Windows.Controls.Button
    $btnThSendProdReview.Content = "Send to Production (Theme)"; $btnThSendProdReview.Height = 30; $btnThSendProdReview.Width = 180
    $btnThSendProdReview.FontWeight = [System.Windows.FontWeights]::Bold; $btnThSendProdReview.FontSize = 11
    $btnThSendProdReview.Background = Get-WpfColor "#4CAF72"; $btnThSendProdReview.Foreground = Get-WpfColor "#FFFFFF"
    $btnThSendProdReview.BorderThickness = 0; $btnThSendProdReview.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnThSendProdReview.Margin = New-Object System.Windows.Thickness(0,0,10,0)
    $btnThSendProdReview.ToolTip = "Runs Send to Production for every design in this theme. Disabled until all designs pass the Production Checklist."
    $btnThSendProdReview.Visibility = "Collapsed"
    $gpRightBtnStack.Children.Add($btnThSendProdReview) | Out-Null
    $gpJob.BtnThSendProdReview = $btnThSendProdReview

    $btnThPrintQReview.Tag = $gpJob
    $btnThPrintQReview.Add_Click({
        $gp = $this.Tag
        foreach ($pj in $gp.Parents) {
            if ($null -ne $pj.LblReviewPrintQStatus) { Send-PJobToPrintingQueue $pj }
        }
    })

    $btnThSendProdReview.Tag = $gpJob
    $btnThSendProdReview.Add_Click({
        $gp = $this.Tag
        $choice = [System.Windows.MessageBox]::Show(
            "Is this a NEW theme (not previously sent to production)?`n`nYes = New Designs`nNo = Fixed Designs",
            "Send to Production - $($gp.GpPath | Split-Path -Leaf)",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question)
        if ($choice -eq [System.Windows.MessageBoxResult]::Cancel) { return }
        $isFix = ($choice -eq [System.Windows.MessageBoxResult]::No)
        foreach ($pj in $gp.Parents) {
            if ($pj.IsProdReady -and $null -ne $pj.LblReviewPrintQStatus) { Send-PJobToProduction $pj $isFix }
        }
    })

    $btnCombineGp = New-Object System.Windows.Controls.Button
    $btnCombineGp.Content = "Copy TSV Data"; $btnCombineGp.Background = Get-WpfColor "#7B4FBF"; $btnCombineGp.Foreground = Get-WpfColor "#FFFFFF"
    $btnCombineGp.FontWeight = [System.Windows.FontWeights]::Bold; $btnCombineGp.Width = 130; $btnCombineGp.Height = 30; $btnCombineGp.BorderThickness = 0
    $btnCombineGp.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnCombineGp.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnCombineGp.Tag = $gpJob
    $btnCombineGp.Add_Click({
        $gp = $this.Tag
        $targetDir = $gp.GpPath
        if ([string]::IsNullOrWhiteSpace($targetDir) -or -not (Test-Path $targetDir)) {
            [System.Windows.MessageBox]::Show("Group folder path is not valid.", "Copy TSV Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $tsvFiles = Get-ChildItem -Path $targetDir -Filter "*_Data.tsv" -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notmatch "(?i)^.*_Design_Data\.tsv$" }
        if ($tsvFiles.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No TSV data files found in:`n$targetDir", "Nothing to Copy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $combined = [ordered]@{}
        foreach ($tsv in $tsvFiles) {
            $line = Get-Content $tsv.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $key = ($line -split "`t")[2]   # DesignName as dedup key (col 2)
            $combined[$key] = $line
        }
        if ($combined.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No TSV data rows found to copy.", "Nothing to Copy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        [System.Windows.Clipboard]::SetText($combined.Values -join "`r`n")
        [System.Windows.MessageBox]::Show(
            "Copied $($combined.Count) row(s) to clipboard.",
            "Copy TSV Data",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    }.GetNewClosure())
    $gpRightBtnStack.Children.Add($btnCombineGp) | Out-Null
    $gpJob.BtnCombineGp = $btnCombineGp

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

    $themeBarStack = New-Object System.Windows.Controls.WrapPanel

    $chkThRename  = New-Object System.Windows.Controls.CheckBox; $chkThRename.Content  = "Rename";  $chkThRename.IsChecked  = $false; $chkThRename.Foreground  = Get-WpfColor "#CCCCCC"; $chkThRename.VerticalAlignment  = "Center"; $chkThRename.Margin  = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThMerge   = New-Object System.Windows.Controls.CheckBox; $chkThMerge.Content   = "Merge";   $chkThMerge.IsChecked   = $false; $chkThMerge.Foreground   = Get-WpfColor "#CCCCCC"; $chkThMerge.VerticalAlignment   = "Center"; $chkThMerge.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThSlice   = New-Object System.Windows.Controls.CheckBox; $chkThSlice.Content   = "Slice";   $chkThSlice.IsChecked   = $false; $chkThSlice.Foreground   = Get-WpfColor "#CCCCCC"; $chkThSlice.VerticalAlignment   = "Center"; $chkThSlice.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThExtract = New-Object System.Windows.Controls.CheckBox; $chkThExtract.Content = "Extract"; $chkThExtract.IsChecked = $false; $chkThExtract.Foreground = Get-WpfColor "#CCCCCC"; $chkThExtract.VerticalAlignment = "Center"; $chkThExtract.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThImage   = New-Object System.Windows.Controls.CheckBox; $chkThImage.Content   = "Image";   $chkThImage.IsChecked   = $false; $chkThImage.Foreground   = Get-WpfColor "#CCCCCC"; $chkThImage.VerticalAlignment   = "Center"; $chkThImage.Margin   = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThColors  = New-Object System.Windows.Controls.CheckBox; $chkThColors.Content  = "Save Colors"; $chkThColors.IsChecked = $false; $chkThColors.Foreground = Get-WpfColor "#CCCCCC"; $chkThColors.VerticalAlignment = "Center"; $chkThColors.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThColors.ToolTip = "Write current color selections to the .3mf files for all cards in this theme"
    $chkThBOD     = New-Object System.Windows.Controls.CheckBox; $chkThBOD.Content     = "Create BOD"; $chkThBOD.IsChecked = $false; $chkThBOD.Foreground = Get-WpfColor "#CCCCCC"; $chkThBOD.VerticalAlignment = "Center"; $chkThBOD.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $chkThBOD.ToolTip = "Reduces the merged Full.3mf to the 5 pairs closest to centre and exports a BOD.gcode.3mf to the Printing Queue for all cards in this theme"
    $chkThPrintQ  = New-Object System.Windows.Controls.CheckBox; $chkThPrintQ.Content  = "Printing Queue"; $chkThPrintQ.IsChecked = $false; $chkThPrintQ.Foreground = Get-WpfColor "#CCCCCC"; $chkThPrintQ.VerticalAlignment = "Center"; $chkThPrintQ.Margin = New-Object System.Windows.Thickness(0,0,20,0)
    $chkThPrintQ.ToolTip = "Copies the existing Full.gcode.3mf to the Printing Queue folder with today's date for all cards in this theme"

    $btnThSelAll   = New-Object System.Windows.Controls.Button; $btnThSelAll.Content   = "Select All";   $btnThSelAll.Background   = Get-WpfColor "#2A2C35"; $btnThSelAll.Foreground   = Get-WpfColor "#FFFFFF"; $btnThSelAll.Width   = 85;  $btnThSelAll.Height   = 25; $btnThSelAll.BorderThickness   = 0; $btnThSelAll.Cursor   = [System.Windows.Input.Cursors]::Hand; $btnThSelAll.Margin   = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThDeselAll = New-Object System.Windows.Controls.Button; $btnThDeselAll.Content = "Deselect All"; $btnThDeselAll.Background = Get-WpfColor "#2A2C35"; $btnThDeselAll.Foreground = Get-WpfColor "#FFFFFF"; $btnThDeselAll.Width = 85;  $btnThDeselAll.Height = 25; $btnThDeselAll.BorderThickness = 0; $btnThDeselAll.Cursor = [System.Windows.Input.Cursors]::Hand; $btnThDeselAll.Margin = New-Object System.Windows.Thickness(0,0,20,0)
    $btnThRevert   = New-Object System.Windows.Controls.Button; $btnThRevert.Content   = "Revert Merge"; $btnThRevert.Background   = Get-WpfColor "#D95F5F"; $btnThRevert.Foreground   = Get-WpfColor "#FFFFFF"; $btnThRevert.Width   = 110; $btnThRevert.Height   = 25; $btnThRevert.BorderThickness   = 0; $btnThRevert.Cursor   = [System.Windows.Input.Cursors]::Hand; $btnThRevert.Margin   = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThProcess    = New-Object System.Windows.Controls.Button; $btnThProcess.Content    = "Process Theme";  $btnThProcess.Background    = Get-WpfColor "#4CAF72"; $btnThProcess.Foreground    = Get-WpfColor "#FFFFFF"; $btnThProcess.Width    = 115; $btnThProcess.Height = 25; $btnThProcess.BorderThickness    = 0; $btnThProcess.Cursor    = [System.Windows.Input.Cursors]::Hand; $btnThProcess.Margin    = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThRenameOnly = New-Object System.Windows.Controls.Button; $btnThRenameOnly.Content = "Rename Theme";   $btnThRenameOnly.Background = Get-WpfColor "#5A6A8A"; $btnThRenameOnly.Foreground = Get-WpfColor "#FFFFFF"; $btnThRenameOnly.Width = 110; $btnThRenameOnly.Height = 25; $btnThRenameOnly.BorderThickness = 0; $btnThRenameOnly.Cursor = [System.Windows.Input.Cursors]::Hand; $btnThRenameOnly.Margin = New-Object System.Windows.Thickness(0,0,8,0)
    $btnThRenameOnly.ToolTip = "Rename all cards in this theme, skipping color validation and heavy tasks"
    $btnThRefresh    = New-Object System.Windows.Controls.Button; $btnThRefresh.Content    = "Refresh Theme";  $btnThRefresh.Background    = Get-WpfColor "#5A78C4"; $btnThRefresh.Foreground    = Get-WpfColor "#FFFFFF"; $btnThRefresh.Width    = 115; $btnThRefresh.Height = 25; $btnThRefresh.BorderThickness    = 0; $btnThRefresh.Cursor    = [System.Windows.Input.Cursors]::Hand; $btnThRefresh.Margin = New-Object System.Windows.Thickness(0,0,8,0)

    $btnThSyncSkus = New-Object System.Windows.Controls.Button
    $btnThSyncSkus.Content = "Sync SKUs"; $btnThSyncSkus.Height = 25; $btnThSyncSkus.Padding = New-Object System.Windows.Thickness(10,0,10,0)
    $btnThSyncSkus.Background = Get-WpfColor "#2A5A7A"; $btnThSyncSkus.Foreground = Get-WpfColor "#FFFFFF"
    $btnThSyncSkus.BorderThickness = 0; $btnThSyncSkus.Cursor = [System.Windows.Input.Cursors]::Hand

    $btnThSyncSkus.Tag = $gpJob
    $btnThSyncSkus.Add_Click({
        $gp = $this.Tag

        # Scan all cards in the theme for missing SKUs that have a sibling match
        $pending = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($pj in $gp.Parents) {
            if ($pj.TxtSKU.Text -ne '') { continue }
            if ([string]::IsNullOrWhiteSpace($pj.DesignName)) { continue }
            $sib = Find-SiblingSkU $pj.FolderPath $pj.DesignName
            if ($null -eq $sib) { continue }
            $pending.Add(@{ PJob = $pj; SiblingInfo = $sib })
        }

        if ($pending.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "All designs in this theme already have SKUs assigned, or no sibling printer variants with SKUs were found.",
                "Sync SKUs", "OK", "Information") | Out-Null
            return
        }

        # Build confirmation preview
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Sync $($pending.Count) SKU(s) from sibling printer variants?")
        [void]$sb.AppendLine("")
        foreach ($item in $pending) {
            [void]$sb.AppendLine("  $($item.SiblingInfo.SKU)   <--   $($item.PJob.DesignName)   (from $($item.SiblingInfo.Printer))")
        }

        $result = [System.Windows.MessageBox]::Show(
            $sb.ToString(), "Sync SKUs - Confirm",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

        # Apply all syncs and update per-card UI
        $errors = 0
        foreach ($item in $pending) {
            $pj  = $item.PJob
            $sib = $item.SiblingInfo
            try {
                Save-SkuToTsv $sib.SKU $pj.FolderPath $pj.AnchorFile
                $pj.TxtSKU.Text = $sib.SKU
                if ($null -ne $pj.BtnSyncSku) { $pj.BtnSyncSku.Visibility = "Collapsed" }
                if ($null -ne $pj.SkuTag)     { $pj.SkuTag.BtnEdit.Visibility = "Visible" }
            } catch {
                $errors++
                Write-Log "Theme Sync SKU error ($($pj.FolderPath)): $($_.Exception.Message)" "ERROR"
            }
        }

        $msg = if ($errors -eq 0) { "Successfully synced $($pending.Count) SKU(s)." } else { "Synced $($pending.Count - $errors) SKU(s).  $errors error(s) - see log." }
        [System.Windows.MessageBox]::Show($msg, "Sync SKUs", "OK", "Information") | Out-Null
    }.GetNewClosure())

    $themeBarStack.Children.Add($chkThRename)    | Out-Null
    $themeBarStack.Children.Add($chkThMerge)     | Out-Null
    $themeBarStack.Children.Add($chkThSlice)     | Out-Null
    $themeBarStack.Children.Add($chkThExtract)   | Out-Null
    $themeBarStack.Children.Add($chkThImage)     | Out-Null
    $themeBarStack.Children.Add($chkThColors)    | Out-Null
    $themeBarStack.Children.Add($chkThBOD)       | Out-Null
    $themeBarStack.Children.Add($chkThPrintQ)    | Out-Null
    $themeBarStack.Children.Add($btnThSelAll)    | Out-Null
    $themeBarStack.Children.Add($btnThDeselAll)  | Out-Null
    $themeBarStack.Children.Add($btnThRevert)    | Out-Null
    $themeBarStack.Children.Add($btnThProcess)   | Out-Null
    $themeBarStack.Children.Add($btnThRenameOnly)| Out-Null
    $themeBarStack.Children.Add($btnThRefresh)   | Out-Null
    $themeBarStack.Children.Add($btnThSyncSkus)  | Out-Null
    $themeBar.Child = $themeBarStack
    $gpStack.Children.Add($themeBar) | Out-Null
    $gpJob.ThemeBar = $themeBar
    $gpJob.BtnThProcess = $btnThProcess
    $gpJob.BtnThRevert = $btnThRevert
    $themeBar.Visibility = if ($script:GlobalMode -eq "Editing") { "Collapsed" } else { "Visible" }

    # --- EDITING TASK BAR (shown only in Editing mode) ---
    $editingThemeBar = New-Object System.Windows.Controls.Border
    $editingThemeBar.Background = Get-WpfColor "#1C1E28"
    $editingThemeBar.BorderBrush = Get-WpfColor "#2A2C38"
    $editingThemeBar.BorderThickness = New-Object System.Windows.Thickness(0,0,0,1)
    $editingThemeBar.Padding = New-Object System.Windows.Thickness(15,10,15,10)
    $editingThemeBar.Visibility = if ($script:GlobalMode -eq "Editing") { "Visible" } else { "Collapsed" }

    $etbStack = New-Object System.Windows.Controls.StackPanel; $etbStack.Orientation = "Horizontal"

    $lblEtbHeader = New-Object System.Windows.Controls.TextBlock
    $lblEtbHeader.Text = "Editing Queue:"; $lblEtbHeader.Foreground = Get-WpfColor "#7AAABB"
    $lblEtbHeader.FontWeight = [System.Windows.FontWeights]::Bold; $lblEtbHeader.FontSize = 12
    $lblEtbHeader.VerticalAlignment = "Center"; $lblEtbHeader.Margin = New-Object System.Windows.Thickness(0,0,18,0)
    $etbStack.Children.Add($lblEtbHeader) | Out-Null

    $chkThEdReNest = New-Object System.Windows.Controls.CheckBox; $chkThEdReNest.Content = "Re-Nest"
    $chkThEdReNest.IsChecked = $true; $chkThEdReNest.Foreground = Get-WpfColor "#CCCCCC"
    $chkThEdReNest.VerticalAlignment = "Center"; $chkThEdReNest.Margin = New-Object System.Windows.Thickness(0,0,20,0)

    $btnThEdSelAll   = New-Object System.Windows.Controls.Button; $btnThEdSelAll.Content   = "Select All"
    $btnThEdSelAll.Background   = Get-WpfColor "#2A2C35"; $btnThEdSelAll.Foreground   = Get-WpfColor "#FFFFFF"
    $btnThEdSelAll.Width = 85; $btnThEdSelAll.Height = 25; $btnThEdSelAll.BorderThickness = 0
    $btnThEdSelAll.Cursor = [System.Windows.Input.Cursors]::Hand; $btnThEdSelAll.Margin = New-Object System.Windows.Thickness(0,0,8,0)

    $btnThEdDeselAll = New-Object System.Windows.Controls.Button; $btnThEdDeselAll.Content = "Deselect All"
    $btnThEdDeselAll.Background = Get-WpfColor "#2A2C35"; $btnThEdDeselAll.Foreground = Get-WpfColor "#FFFFFF"
    $btnThEdDeselAll.Width = 90; $btnThEdDeselAll.Height = 25; $btnThEdDeselAll.BorderThickness = 0
    $btnThEdDeselAll.Cursor = [System.Windows.Input.Cursors]::Hand; $btnThEdDeselAll.Margin = New-Object System.Windows.Thickness(0,0,20,0)

    $btnThEdQueueAll = New-Object System.Windows.Controls.Button; $btnThEdQueueAll.Content = "Queue Group"
    $btnThEdQueueAll.Background = Get-WpfColor "#3A5080"; $btnThEdQueueAll.Foreground = Get-WpfColor "#FFFFFF"
    $btnThEdQueueAll.Width = 105; $btnThEdQueueAll.Height = 25; $btnThEdQueueAll.BorderThickness = 0
    $btnThEdQueueAll.FontWeight = [System.Windows.FontWeights]::Bold
    $btnThEdQueueAll.Cursor = [System.Windows.Input.Cursors]::Hand

    $etbStack.Children.Add($chkThEdReNest)  | Out-Null
    $etbStack.Children.Add($btnThEdSelAll)  | Out-Null
    $etbStack.Children.Add($btnThEdDeselAll)    | Out-Null
    $etbStack.Children.Add($btnThEdQueueAll)    | Out-Null
    $editingThemeBar.Child = $etbStack
    $gpStack.Children.Add($editingThemeBar) | Out-Null
    $gpJob.EditingThemeBar = $editingThemeBar

    $parentListStack = New-Object System.Windows.Controls.StackPanel
    $parentListStack.Margin = New-Object System.Windows.Thickness(15)
    $gpStack.Children.Add($parentListStack) | Out-Null
    $gpJob.ParentListStack = $parentListStack

    foreach ($pKey in $parentDict.Keys) {
        $pJob = Build-PJob $pKey $parentDict[$pKey] $gpJob
        $gpJob.Parents.Add($pJob) | Out-Null
    }
    Update-GpFileCount $gpJob
    Update-ThRevertButtonState $gpJob
    Update-ThProcessButtonState $gpJob
    Update-GlobalProcessAllStatus

    # --- THEME TASK BAR HANDLERS (wired after Parents are populated) ---
    $chkThRename.Tag = $gpJob
    $chkThRename.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkRename.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThMerge.Tag = $gpJob
    $chkThMerge.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { if ($p.ChkMerge.IsEnabled) { $p.ChkMerge.IsChecked = $s } }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThSlice.Tag = $gpJob
    $chkThSlice.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkSlice.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThExtract.Tag = $gpJob
    $chkThExtract.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkExtract.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThImage.Tag = $gpJob
    $chkThImage.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkImage.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThColors.Tag = $gpJob
    $chkThColors.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkColors.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThBOD.Tag = $gpJob
    $chkThBOD.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkBOD.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $chkThPrintQ.Tag = $gpJob
    $chkThPrintQ.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { $p.ChkPrintQ.IsChecked = $s }; Update-ThProcessButtonState $this.Tag; Update-GlobalProcessAllStatus })

    $btnThSelAll.Tag = @{ GpJob = $gpJob; Chks = @{ Rename = $chkThRename; Merge = $chkThMerge; Slice = $chkThSlice; Extract = $chkThExtract; Image = $chkThImage; Colors = $chkThColors } }
    $btnThSelAll.Add_Click({
        $t = $this.Tag
        foreach ($p in $t.GpJob.Parents) {
            $p.ChkRename.IsChecked = $true; $p.ChkSlice.IsChecked = $true
            $p.ChkExtract.IsChecked = $true; $p.ChkImage.IsChecked = $true; $p.ChkColors.IsChecked = $true
            if ($p.ChkMerge.IsEnabled) { $p.ChkMerge.IsChecked = $true }
        }
        $t.Chks.Rename.IsChecked = $true; $t.Chks.Merge.IsChecked = $true
        $t.Chks.Slice.IsChecked = $true; $t.Chks.Extract.IsChecked = $true; $t.Chks.Image.IsChecked = $true; $t.Chks.Colors.IsChecked = $true
        Update-ThProcessButtonState $t.GpJob; Update-GlobalProcessAllStatus
    })

    $btnThDeselAll.Tag = @{ GpJob = $gpJob; Chks = @{ Rename = $chkThRename; Merge = $chkThMerge; Slice = $chkThSlice; Extract = $chkThExtract; Image = $chkThImage; Colors = $chkThColors; BOD = $chkThBOD; PrintQ = $chkThPrintQ } }
    $btnThDeselAll.Add_Click({
        $t = $this.Tag
        foreach ($p in $t.GpJob.Parents) {
            $p.ChkRename.IsChecked = $false; $p.ChkMerge.IsChecked = $false; $p.ChkSlice.IsChecked = $false
            $p.ChkExtract.IsChecked = $false; $p.ChkImage.IsChecked = $false; $p.ChkColors.IsChecked = $false; $p.ChkLogs.IsChecked = $false
            $p.ChkBOD.IsChecked = $false; $p.ChkPrintQ.IsChecked = $false
        }
        $t.Chks.Rename.IsChecked = $false; $t.Chks.Merge.IsChecked = $false
        $t.Chks.Slice.IsChecked = $false; $t.Chks.Extract.IsChecked = $false; $t.Chks.Image.IsChecked = $false; $t.Chks.Colors.IsChecked = $false
        $t.Chks.BOD.IsChecked = $false; $t.Chks.PrintQ.IsChecked = $false
        Update-ThProcessButtonState $t.GpJob; Update-GlobalProcessAllStatus
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
        Update-ThRevertButtonState $gp
    })

    $btnThProcess.Tag = $gpJob
    $btnThProcess.Add_Click({
        $gp = $this.Tag
        foreach ($p in $gp.Parents) { Enqueue-PJob $p $gp }
        if ($script:activeProcess -eq $null -and $script:processQueue.Count -gt 0) { Start-NextProcess }
    })

    $btnThRenameOnly.Tag = $gpJob
    $btnThRenameOnly.Add_Click({
        $gp = $this.Tag
        # Count cards that would be renamed (not yet queued/done, no collision)
        $eligible = @($gp.Parents | Where-Object { -not $_.IsQueued -and -not $_.IsDone -and -not $_.HasCollision })
        if ($eligible.Count -eq 0) { return }
        $result = [System.Windows.MessageBox]::Show(
            "Colors may be unmatched. This will rename files only for all $($eligible.Count) card(s) in this theme`n(merge, slice, extract, and image will be skipped).`n`nContinue?",
            "Rename Theme Only",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
        foreach ($p in $eligible) {
            $p.RenameOnlyBypass = $true
            $p.ChkRename.IsChecked = $true
            Enqueue-PJob $p $gp
        }
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
    $cbGpTag.Tag = $gpJob
    $cbGpTag.Add_SelectionChanged({ foreach ($p in $this.Tag.Parents) { Update-ParentPreview $p $this.Tag } })

    # --- EDITING THEME BAR HANDLERS (wired after Parents are populated) ---
    $chkThEdReNest.Tag = $gpJob
    $chkThEdReNest.Add_Click({ $s = [bool]$this.IsChecked; foreach ($p in $this.Tag.Parents) { if ($null -ne $p.ChkEdReNest) { $p.ChkEdReNest.IsChecked = $s } } })

    $btnThEdSelAll.Tag = $gpJob
    $btnThEdSelAll.Add_Click({
        foreach ($p in $this.Tag.Parents) {
            if ($null -ne $p.ChkEdReNest) { $p.ChkEdReNest.IsChecked = $true }
        }
        $chkThEdReNest.IsChecked = $true
    }.GetNewClosure())

    $btnThEdDeselAll.Tag = $gpJob
    $btnThEdDeselAll.Add_Click({
        foreach ($p in $this.Tag.Parents) {
            if ($null -ne $p.ChkEdReNest) { $p.ChkEdReNest.IsChecked = $false }
        }
        $chkThEdReNest.IsChecked = $false
    }.GetNewClosure())

    $btnThEdQueueAll.Tag = $gpJob
    $btnThEdQueueAll.Add_Click({
        $gp = $this.Tag
        foreach ($p in $gp.Parents) { Enqueue-EditJob $p $gp }
        if ($null -eq $script:editActiveJob) { Start-NextEditJob }
    })

    $mainStack.Children.Add($container) | Out-Null
}

# --- Edit queue functions ---

function Enqueue-EditJob($pJob, $gpJob) {
    if ($pJob.EdIsQueued) { return }
    $doReNest = ($null -ne $pJob.ChkEdReNest -and [bool]$pJob.ChkEdReNest.IsChecked)
    if (-not $doReNest) { return }   # nothing to queue right now
    $pJob.EdIsQueued = $true
    if ($null -ne $pJob.LblEdQueueStatus) {
        $pJob.LblEdQueueStatus.Text = "Queued"; $pJob.LblEdQueueStatus.Foreground = Get-WpfColor "#E8A135"
    }
    if ($null -ne $pJob.BtnEdQueue) { $pJob.BtnEdQueue.IsEnabled = $false }
    $script:editQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob })
}

function Start-NextEditJob {
    if ($null -ne $script:editActiveJob -or $script:editQueue.Count -eq 0) { return }
    $jobWrapper = $script:editQueue.Dequeue()
    $pJob  = $jobWrapper.PJob
    $gpJob = $jobWrapper.GpJob
    $script:editActiveJob = $jobWrapper

    if ($null -ne $pJob.LblEdQueueStatus) {
        $pJob.LblEdQueueStatus.Text = "Running..."; $pJob.LblEdQueueStatus.Foreground = Get-WpfColor "#F0A030"
    }

    # Trigger Re-Nest via the card's existing renestTag
    $rt = if ($null -ne $pJob.BtnRunRenest) { $pJob.BtnRunRenest.Tag } else { $null }
    if ($null -eq $rt -or [string]::IsNullOrEmpty($rt.FinalPath) -or -not (Test-Path -LiteralPath $rt.FinalPath)) {
        if ($null -ne $pJob.LblEdQueueStatus) {
            $pJob.LblEdQueueStatus.Text = "No Final.3mf"; $pJob.LblEdQueueStatus.Foreground = Get-WpfColor "#D95F5F"
        }
        $pJob.EdIsQueued = $false
        if ($null -ne $pJob.BtnEdQueue) { $pJob.BtnEdQueue.IsEnabled = $true }
        $script:editActiveJob = $null
        if ($script:editQueue.Count -gt 0) { Start-NextEditJob }
        return
    }

    # Reset Re-Nest review state and fire
    $rt.BorderReview.Visibility = "Collapsed"
    $rt.BtnRun.IsEnabled = $false; $rt.BtnRun.Content = "Running..."
    $rt.BtnRun.Background = Get-WpfColor "#333333"; $rt.BtnRun.Foreground = Get-WpfColor "#888888"
    $rt.LblStatus.Text = "Re-nesting..."; $rt.LblStatus.Foreground = Get-WpfColor "#F0A030"

    $tmpBase = Join-Path $env:TEMP ("renest_" + [System.Guid]::NewGuid().ToString("N"))
    $rt.LogOut = $tmpBase + "_out.txt"; $rt.LogErr = $tmpBase + "_err.txt"
    [System.IO.File]::WriteAllText($rt.LogOut, ""); [System.IO.File]::WriteAllText($rt.LogErr, "")
    $rt.RenestTempOut = Join-Path $env:TEMP ("renest_" + [System.Guid]::NewGuid().ToString("N") + ".3mf")
    $psArgs = @("-NoProfile","-ExecutionPolicy","Bypass","-File","`"$($rt.WorkerPath)`"","-FinalPath","`"$($rt.FinalPath)`"","-OutputPath","`"$($rt.RenestTempOut)`"","-NoConfirm")
    $rt.Proc = Start-Process -FilePath "powershell.exe" -ArgumentList $psArgs `
        -NoNewWindow -RedirectStandardOutput $rt.LogOut -RedirectStandardError $rt.LogErr -PassThru
    $script:_RenestActive = $rt

    $renestTimer = New-Object System.Windows.Threading.DispatcherTimer
    $renestTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $rt.Timer = $renestTimer
    $renestTimer.Add_Tick($script:_RenestTickSB)
    $renestTimer.Start()
}

# --- Renest tick handler (polls hidden worker process every 250 ms) ---
$script:_RenestActive = $null
$script:_RenestTickSB = {
    $t2 = $script:_RenestActive
    if ($null -eq $t2 -or $null -eq $t2.Proc) { return }
    if (-not $t2.Proc.HasExited) { return }

    # Process has finished â€” stop polling
    $t2.Timer.Stop()
    $script:_RenestActive = $null

    # Restore button immediately so UI is never stuck
    $t2.BtnRun.Content = "Run Re-Nest"; $t2.BtnRun.IsEnabled = $true
    $t2.BtnRun.Background = Get-WpfColor "#2E5A42"; $t2.BtnRun.Foreground = Get-WpfColor "#FFFFFF"

    # Read exit code safely â€” Process.ExitCode can throw on some versions
    $exitCode = -1
    try { $exitCode = [int]$t2.Proc.ExitCode } catch {}

    # Output path was pre-computed as a temp file before the worker launched
    $stemR  = [System.IO.Path]::GetFileNameWithoutExtension($t2.FinalPath) -replace '(?i)_Final$', ''
    $outDir = [System.IO.Path]::GetDirectoryName($t2.FinalPath)
    $t2.RenestPath = $t2.RenestTempOut

    # Move worker debug file to TEMP (keep permanent copy only if user requests it)
    $t2.DebugTempPath = ""
    try {
        $permDebugPath = Join-Path $outDir ($stemR + "_Renest_debug.txt")
        $errContent = if ($t2.LogErr -and (Test-Path -LiteralPath $t2.LogErr)) { [System.IO.File]::ReadAllText($t2.LogErr).Trim() } else { "" }
        if (Test-Path -LiteralPath $permDebugPath) {
            # Append stderr (if any) then move to temp
            if ($errContent) {
                [System.IO.File]::AppendAllText($permDebugPath, "`r`n--- stderr ---`r`n" + $errContent, [System.Text.Encoding]::UTF8)
            }
            $debugTmp = Join-Path $env:TEMP ("renest_debug_" + [System.Guid]::NewGuid().ToString("N") + ".txt")
            Move-Item -LiteralPath $permDebugPath -Destination $debugTmp -Force
            $t2.DebugTempPath = $debugTmp
        } elseif ($errContent) {
            # No debug file from worker but we have stderr - save that to temp
            $debugTmp = Join-Path $env:TEMP ("renest_debug_" + [System.Guid]::NewGuid().ToString("N") + ".txt")
            [System.IO.File]::WriteAllText($debugTmp, "--- stderr ---`r`n" + $errContent, [System.Text.Encoding]::UTF8)
            $t2.DebugTempPath = $debugTmp
        }
    } catch {}
    try { if ($t2.LogOut -and (Test-Path -LiteralPath $t2.LogOut)) { Remove-Item -LiteralPath $t2.LogOut -Force -ErrorAction SilentlyContinue } } catch {}
    try { if ($t2.LogErr -and (Test-Path -LiteralPath $t2.LogErr)) { Remove-Item -LiteralPath $t2.LogErr -Force -ErrorAction SilentlyContinue } } catch {}
    try { if ($t2.RotCorrTempPath -and (Test-Path -LiteralPath $t2.RotCorrTempPath)) { Remove-Item -LiteralPath $t2.RotCorrTempPath -Force -ErrorAction SilentlyContinue } } catch {}

    # Success = exit 0 OR the renest file was actually created (handles edge-case ExitCode issues)
    $renestExists = Test-Path -LiteralPath $t2.RenestPath
    if ($exitCode -eq 0 -or $renestExists) {
        $t2.LblStatus.Text = "Done - loading preview..."; $t2.LblStatus.Foreground = Get-WpfColor "#4CAF72"
        try {
            $tmpImgDir = Join-Path $env:TEMP ("renest_imgs_" + [System.Guid]::NewGuid().ToString("N"))
            New-Item -ItemType Directory -Path $tmpImgDir -Force | Out-Null
            # Only update the right (col 1) overlay with the new Renest top_1.png
            # Left (col 0) keeps its original Nest Source image loaded at build time
            $newPickPath = if ($renestExists) { Extract-3mfPickImage $t2.RenestPath $tmpImgDir "new" } else { $null }
            $t2.P.RnImgRight.Source = if ($newPickPath -and (Test-Path $newPickPath)) { Load-WpfImage $newPickPath } else { $null }
            if ($null -ne $t2.P.RnLblRight) { $t2.P.RnLblRight.Text = "Re-Nest" }
            # Ensure overlays are visible (they may have been hidden if in non-Editing mode)
            $t2.P.RnOvLeft.Visibility  = "Visible"
            $t2.P.RnOvRight.Visibility = "Visible"
            $t2.LblStatus.Text = "Done - review and confirm below."; $t2.LblStatus.Foreground = Get-WpfColor "#4CAF72"
        } catch {
            $t2.LblStatus.Text = "Done (preview error: $($_.Exception.Message))"; $t2.LblStatus.Foreground = Get-WpfColor "#F0A030"
        }
        $t2.BorderReview.Visibility = "Visible"
    } else {
        $t2.LblStatus.Text = "Failed (exit $exitCode)"; $t2.LblStatus.Foreground = Get-WpfColor "#D95F5F"
    }

    # If this run was triggered by the edit queue, update queue state and advance
    if ($null -ne $script:editActiveJob -and [object]::ReferenceEquals($script:editActiveJob.PJob, $t2.P)) {
        $qPJob = $script:editActiveJob.PJob
        $qPJob.EdIsQueued = $false
        $statusOk = ($exitCode -eq 0 -or (Test-Path -LiteralPath $t2.RenestPath))
        if ($null -ne $qPJob.LblEdQueueStatus) {
            $qPJob.LblEdQueueStatus.Text = if ($statusOk) { "Done" } else { "Failed" }
            $qPJob.LblEdQueueStatus.Foreground = Get-WpfColor $(if ($statusOk) { "#4CAF72" } else { "#D95F5F" })
        }
        if ($null -ne $qPJob.BtnEdQueue) { $qPJob.BtnEdQueue.IsEnabled = $true }
        $script:editActiveJob = $null
        if ($script:editQueue.Count -gt 0) { Start-NextEditJob }
    }
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
                        $raw = $statusText.Trim()
                        # Parse "SLICING... XX%" to drive progress bar; show clean label without the number
                        $slicePct = $null
                        if ($raw -match '^SLICING\.\.\.\s*(\d+)%') {
                            $slicePct = [int]$matches[1]
                            $raw = "SLICING..."
                        }
                        $txt = "[ $raw ]"
                        $pJob.CardStatusLabel.Text = $txt
                        $pJob.PickStatusLabel.Text = $txt
                        $pJob.BtnApply.Content = $raw
                        # Mirror to editing-panel status labels so progress is visible in Editing mode
                        $editFg = Get-WpfColor $(if ($raw -match '(?i)error') { "#D95F5F" } else { "#F0A030" })
                        $editTxt = if ($slicePct -ne $null) { "Exporting GCode... $slicePct%" } else { $raw }
                        if ($null -ne $pJob.RenestStatusLbl) {
                            $pJob.RenestStatusLbl.Text = $editTxt; $pJob.RenestStatusLbl.Foreground = $editFg
                        }
                        if ($null -ne $pJob.LblEdQueueStatus) {
                            $pJob.LblEdQueueStatus.Text = $editTxt; $pJob.LblEdQueueStatus.Foreground = $editFg
                        }
                        if ($slicePct -ne $null) {
                            $pJob.CardProgressBar.Value = $slicePct
                            $pJob.CardProgressBar.Visibility = "Visible"
                            $pJob.PickProgressBar.Value = $slicePct
                            $pJob.PickProgressBar.Visibility = "Visible"
                        } else {
                            $pJob.CardProgressBar.Visibility = "Collapsed"
                            $pJob.PickProgressBar.Visibility = "Collapsed"
                        }
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
                            if ($script:LibraryColors.Contains($selectedName)) {
                                $verifiedHex = $script:LibraryColors[$selectedName]
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

            # Clear progress bars and overlays (common to all job types)
            $pJob.CardProgressBar.Visibility = "Collapsed"; $pJob.CardProgressBar.Value = 0
            $pJob.PickProgressBar.Visibility = "Collapsed"; $pJob.PickProgressBar.Value = 0
            $pJob.ProcessingOverlay.Visibility = "Collapsed"
            $pJob.PickProcessingOverlay.Visibility = "Collapsed"
            $pJob.RowPanel.IsEnabled = $true

            $isSliceOnly = ($script:activeProcessJob.SliceOnly -eq $true)

            if ($isSliceOnly) {
                # Editing-mode slice export finished â€” leave card reusable, don't touch FilePrepPanel state
                $pJob.IsQueued = $false; $pJob.IsDone = $false
                if ($null -ne $pJob.RenestStatusLbl) {
                    $pJob.RenestStatusLbl.Text = "Export GCode complete!"; $pJob.RenestStatusLbl.Foreground = Get-WpfColor "#4CAF72"
                }
                if ($null -ne $pJob.LblEdQueueStatus) {
                    $pJob.LblEdQueueStatus.Text = "Done"; $pJob.LblEdQueueStatus.Foreground = Get-WpfColor "#4CAF72"
                }
                if ($null -ne $pJob.BtnEdQueue) { $pJob.BtnEdQueue.IsEnabled = $true }
                if ($null -ne $pJob.BtnRunRenest) { $pJob.BtnRunRenest.IsEnabled = $true }
            } else {
                # Normal merge/process job finished â€” KEEP/REVERT state
                $pJob.IsDone = $true
                $pJob.ChkRename.IsEnabled = $true; $pJob.ChkMerge.IsEnabled = $true; $pJob.ChkSlice.IsEnabled = $true
                $pJob.ChkExtract.IsEnabled = $true; $pJob.ChkImage.IsEnabled = $true

                $pJob.BtnApply.Content = "KEEP"; $pJob.BtnApply.Background = Get-WpfColor "#4CAF72"
                $pJob.BtnApply.IsEnabled = $true; $pJob.BtnApply.Width = 70
                # Mirror completion to editing panel in case mode was switched
                if ($null -ne $pJob.RenestStatusLbl) {
                    $pJob.RenestStatusLbl.Text = "Export GCode complete!"; $pJob.RenestStatusLbl.Foreground = Get-WpfColor "#4CAF72"
                }
                if ($null -ne $pJob.LblEdQueueStatus) {
                    $pJob.LblEdQueueStatus.Text = "Done"; $pJob.LblEdQueueStatus.Foreground = Get-WpfColor "#4CAF72"
                    if ($null -ne $pJob.BtnEdQueue) { $pJob.BtnEdQueue.IsEnabled = $true }
                }

                # Only enable revert controls if a Nest.3mf actually exists (failed merges won't have one)
                $nestNow = Get-ChildItem -Path $pJob.FolderPath -Filter "*Nest.3mf" -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '(?i)[._ ]Nest\.3mf$' } | Select-Object -First 1
                if ($nestNow) {
                    $pJob.BtnRevertDone.Visibility = "Visible"
                    if ($pJob.BtnRevertMerge) { $pJob.BtnRevertMerge.IsEnabled = $true; $pJob.BtnRevertMerge.Background = Get-WpfColor "#D95F5F"; $pJob.BtnRevertMerge.Foreground = Get-WpfColor "#FFFFFF"; $pJob.BtnRevertMerge.ToolTip = $null }
                } else {
                    $pJob.BtnRevertDone.Visibility = "Collapsed"
                    if ($pJob.BtnRevertMerge) { $pJob.BtnRevertMerge.IsEnabled = $false; $pJob.BtnRevertMerge.Background = Get-WpfColor "#3A3A3A"; $pJob.BtnRevertMerge.Foreground = Get-WpfColor "#666666"; $pJob.BtnRevertMerge.ToolTip = "No Nest.3mf found - merge may have failed" }
                }
            }

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
    $selectedPaths = [NativeFolderBrowser]::ShowDialog($hwnd, "Select folders containing 3MF source files")

    # 3. Stop if the user closed the window or clicked Cancel
    if ($null -eq $selectedPaths -or $selectedPaths.Count -eq 0) { return }

    $lblGlobalTitle.Text = "Scanning selected folders..."
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    # 4. Build queue using the same anchor-discovery logic as drag-and-drop
    # (finds Full > Nest > Final > any .3mf > .stl > .png per folder)
    $newGpQueue = Get-AnchorQueue $selectedPaths

    # Remove folders already loaded in the UI
    foreach ($gpPath in @($newGpQueue.Keys)) {
        foreach ($parentPath in @($newGpQueue[$gpPath].Keys)) {
            $exists = $false
            foreach ($j in $script:jobs) {
                foreach ($parentJob in $j.Parents) { if ($parentJob.FolderPath -eq $parentPath) { $exists = $true; break } }
                if ($exists) { break }
            }
            if ($exists) { $newGpQueue[$gpPath].Remove($parentPath) }
        }
        if ($newGpQueue[$gpPath].Count -eq 0) { $newGpQueue.Remove($gpPath) }
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

    $dbgLog = Join-Path $PSScriptRoot "..\launchers\CardQueueEditor_debug.txt"
    [System.IO.File]::WriteAllText($dbgLog, "[DROP] Started`r`nDropped: $($dropped -join ', ')`r`n")

    $lblGlobalTitle.Text = "Scanning dropped folders..."
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    try {
        [System.IO.File]::AppendAllText($dbgLog, "[STEP] Calling Get-AnchorQueue`r`n")
        $newGpQueue = Get-AnchorQueue $dropped
        [System.IO.File]::AppendAllText($dbgLog, "[STEP] Get-AnchorQueue returned $($newGpQueue.Count) group(s)`r`n")

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
                        [System.IO.File]::AppendAllText($dbgLog, "[STEP] Build-PJob (existing GP): $pKey`r`n")
                        $pJob = Build-PJob $pKey $newGpQueue[$gpPath][$pKey] $existingGp
                        $existingGp.Parents.Add($pJob) | Out-Null
                        [System.IO.File]::AppendAllText($dbgLog, "[STEP] Build-PJob done: $pKey`r`n")
                    }
                } else {
                    [System.IO.File]::AppendAllText($dbgLog, "[STEP] Build-GpJob: $gpPath`r`n")
                    Build-GpJob $gpPath $newGpQueue[$gpPath]
                    [System.IO.File]::AppendAllText($dbgLog, "[STEP] Build-GpJob done: $gpPath`r`n")
                }
            }
        }

        $lblGlobalTitle.Text = "Queue Dashboard ($($script:jobs.Count) Theme(s) found)"
        if ($script:jobs.Count -gt 0) { Update-GlobalProcessAllStatus }
        [System.IO.File]::AppendAllText($dbgLog, "[DROP] Complete. Jobs: $($script:jobs.Count)`r`n")
    } catch {
        $errMsg = $_.Exception.Message
        $errPos = $_.InvocationInfo.PositionMessage
        [System.IO.File]::AppendAllText($dbgLog, "[ERROR] $errMsg`r`nAt: $errPos`r`n")
        $lblGlobalTitle.Text = "Drop error - log saved next to dropped files"
        $dlgResult = [System.Windows.MessageBox]::Show(
            "An error occurred loading files.`n`nLog: $dbgLog`n`nOpen the log now?",
            "CardQueueEditor Error",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Error
        )
        if ($dlgResult -eq [System.Windows.MessageBoxResult]::Yes) {
            Start-Process notepad.exe $dbgLog
        }
    }
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

$btnStopQueue.Add_Click({
    $hasActive  = ($null -ne $script:activeProcess -and -not $script:activeProcess.HasExited)
    $hasQueued  = $script:processQueue.Count -gt 0

    if (-not $hasActive -and -not $hasQueued) {
        [System.Windows.MessageBox]::Show("There are no active or queued tasks to stop.", "Stop Queue", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    # Build a confirmation message showing what will be stopped
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Stop the current task and clear all remaining queued tasks?")
    if ($hasActive -and $null -ne $script:activeProcessJob) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Currently running:")
        [void]$sb.AppendLine("  $($script:activeProcessJob.PJob.AnchorFile.BaseName)")
    }
    if ($hasQueued) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Queued ($($script:processQueue.Count) task(s)):")
        foreach ($jw in $script:processQueue) {
            [void]$sb.AppendLine("  $($jw.PJob.AnchorFile.BaseName)")
        }
    }

    $result = [System.Windows.MessageBox]::Show(
        $sb.ToString().TrimEnd(),
        "Stop Queue",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

    # Helper: reset one pJob's UI back to ready state
    $resetPJob = {
        param($pj)
        $pj.IsQueued = $false
        $pj.BtnApply.Content    = "Add to Queue"
        $pj.BtnApply.Background = Get-WpfColor "#4CAF72"
        $pj.BtnApply.IsEnabled  = $true
        $pj.BtnApply.Width      = 100
        $pj.RowPanel.IsEnabled  = $true
        $pj.ProcessingOverlay.Visibility     = "Collapsed"
        $pj.PickProcessingOverlay.Visibility = "Collapsed"
        if ($null -ne $pj.RenestStatusLbl)  { $pj.RenestStatusLbl.Text  = "Stopped"; $pj.RenestStatusLbl.Foreground  = Get-WpfColor "#E8A135" }
        if ($null -ne $pj.LblEdQueueStatus) { $pj.LblEdQueueStatus.Text = ""       ; $pj.LblEdQueueStatus.Foreground = Get-WpfColor "#888A9A" }
    }

    # Kill the active worker process
    if ($hasActive) {
        try { $script:activeProcess.Kill() } catch {}
    }

    # Reset the currently running job card
    if ($null -ne $script:activeProcessJob) {
        & $resetPJob $script:activeProcessJob.PJob
    }

    # Drain the queue, resetting each card
    $affectedGpJobs = [System.Collections.Generic.HashSet[object]]::new()
    if ($null -ne $script:activeProcessJob) { [void]$affectedGpJobs.Add($script:activeProcessJob.GpJob) }
    while ($script:processQueue.Count -gt 0) {
        $jw = $script:processQueue.Dequeue()
        & $resetPJob $jw.PJob
        [void]$affectedGpJobs.Add($jw.GpJob)
    }

    # Re-enable theme/prefix/tag controls on affected grandparent groups
    # so the user can adjust settings and re-queue if needed
    foreach ($gp in $affectedGpJobs) {
        $gp.GpRenameConfirmed = $false
        if ($null -ne $gp.TBTheme)  { $gp.TBTheme.IsEnabled  = $true }
        if ($null -ne $gp.CbPrefix) { $gp.CbPrefix.IsEnabled = $true }
        if ($null -ne $gp.ChkSkip)  { $gp.ChkSkip.IsEnabled  = $true }
    }

    $script:activeProcess    = $null
    $script:activeProcessJob = $null

    [System.Windows.MessageBox]::Show(
        "Queue stopped.  All active and queued tasks have been cleared.",
        "Stop Queue",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
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

# â”€â”€ Import SKUs from CSV â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
$btnImportSkus.Add_Click({
    # Open file picker
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Title  = "Select SKU CSV File"
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.InitialDirectory = [System.Environment]::GetFolderPath('UserProfile')
    if ($ofd.ShowDialog() -ne $true) { return }

    try {
        # Read CSV â€” optional header row "SKU", then one SKU value per line.
        # Accepts any format matching the shared SKU pattern (3+ non-whitespace chars).
        $csvLines    = Get-Content $ofd.FileName -ErrorAction Stop
        $skusFromCsv = [System.Collections.Generic.List[string]]::new()
        $firstLine   = $true
        foreach ($line in $csvLines) {
            $val = $line.Trim().Trim('"')
            if ($firstLine) { $firstLine = $false; if ($val -ieq 'SKU') { continue } }
            if ($val -match $script:SkuPattern) { $skusFromCsv.Add($val) }
        }

        if ($skusFromCsv.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No valid SKUs found in:`n$($ofd.FileName)",
                "Import SKUs", "OK", "Warning") | Out-Null
            return
        }

        # Collect all pJobs that currently have no SKU, in queue order
        $targetJobs = [System.Collections.Generic.List[object]]::new()
        foreach ($gp in $script:jobs) {
            foreach ($p in $gp.Parents) {
                if ($null -ne $p.TxtSKU -and [string]::IsNullOrWhiteSpace($p.TxtSKU.Text)) {
                    $targetJobs.Add($p)
                }
            }
        }

        if ($targetJobs.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "All cards in the queue already have SKUs assigned.",
                "Import SKUs", "OK", "Information") | Out-Null
            return
        }

        $assignCount = [Math]::Min($skusFromCsv.Count, $targetJobs.Count)

        # Build confirmation list
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Assign $assignCount SKU(s) to queue cards?")
        [void]$sb.AppendLine("")
        for ($i = 0; $i -lt $assignCount; $i++) {
            $folderName = Split-Path $targetJobs[$i].FolderPath -Leaf
            [void]$sb.AppendLine("  $($skusFromCsv[$i])  ->  $folderName")
        }
        if ($targetJobs.Count -gt $skusFromCsv.Count) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("  ($($targetJobs.Count - $skusFromCsv.Count) card(s) will remain unassigned - CSV has fewer SKUs)")
        }
        if ($skusFromCsv.Count -gt $targetJobs.Count) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("  ($($skusFromCsv.Count - $targetJobs.Count) SKU(s) in CSV will not be used)")
        }

        $result = [System.Windows.MessageBox]::Show(
            $sb.ToString(), "Import SKUs - Confirm",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

        # Apply â€” bypass the lock, write to TSV, update the display box
        $errors = 0
        for ($i = 0; $i -lt $assignCount; $i++) {
            $p   = $targetJobs[$i]
            $sku = $skusFromCsv[$i]
            try {
                Save-SkuToTsv $sku $p.FolderPath $p.AnchorFile
                $p.TxtSKU.Text = $sku
            } catch {
                $errors++
                Write-Log "Import SKU error ($($p.FolderPath)): $($_.Exception.Message)" "ERROR"
            }
        }

        $msg = if ($errors -eq 0) {
            "Successfully assigned $assignCount SKU(s)."
        } else {
            "Assigned $($assignCount - $errors) SKU(s).  $errors error(s) - see debug log."
        }
        [System.Windows.MessageBox]::Show($msg, "Import SKUs", "OK", "Information") | Out-Null

    } catch {
        [System.Windows.MessageBox]::Show(
            "Error reading CSV:`n$($_.Exception.Message)",
            "Import SKUs", "OK", "Error") | Out-Null
    }
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

# â”€â”€ Global workspace mode buttons (File Prep | Editing | Review) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  LIBRARIES PANEL
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function Save-FilamentLibrary {
    Write-Log "Save-FilamentLibrary: start ($($script:LibraryColors.Count) entries)"
    try {
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("N/A,,,,,,,")
        foreach ($kv in $script:LibraryColors.GetEnumerator()) {
            $hex = $kv.Value  # #RRGGBBFF or #RRGGBB
            $r = [Convert]::ToInt32($hex.Substring(1,2),16)
            $g = [Convert]::ToInt32($hex.Substring(3,2),16)
            $b = [Convert]::ToInt32($hex.Substring(5,2),16)
            $gradCols = @("","","","")
            if ($script:LibraryGradients.Contains($kv.Key)) {
                $gradVals = $script:LibraryGradients[$kv.Key]
                for ($gi = 0; $gi -lt 4 -and $gi -lt $gradVals.Count; $gi++) {
                    $gradCols[$gi] = $gradVals[$gi]
                }
            }
            $lines.Add("$($kv.Key),$r,$g,$b,$($gradCols -join ',')")
        }
        [System.IO.File]::WriteAllLines($colorCsvPath, $lines, (New-Object System.Text.UTF8Encoding($false)))
        # Refresh $script:HexToName
        $script:HexToNameLib = @{}
        foreach ($kv in $script:LibraryColors.GetEnumerator()) {
            $hex9 = $kv.Value
            $hex7 = $hex9.Substring(0,7)
            $script:HexToName[$hex9] = $kv.Key; $script:HexToName[$hex7] = $kv.Key
        }
        Write-Log "Save-FilamentLibrary: success"
        return $true
    } catch {
        Write-Log "Save-FilamentLibrary: FAILED - $($_.Exception.Message)" "ERROR"
        Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"
        return $false
    }
}

function Save-PurgeDictionary {
    Write-Log "Save-PurgeDictionary: start ($($script:PurgeDict.Count) entries)"
    try {
        # Tab-delimited with a UTF-8 BOM â€” this is the format Excel actually writes back
        # to this .csv on save on this machine. Writing tab here too keeps the on-disk
        # format consistent regardless of whether the last save came from Excel or here.
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("Source_Filament`tTarget_Filament`tTuned`tTuned_Volume`tBase Volume")
        foreach ($row in $script:PurgeDict) {
            $tunedStr = if ($row.Tuned) { "TRUE" } else { "" }
            $lines.Add("$($row.Source_Filament)`t$($row.Target_Filament)`t$tunedStr`t$($row.Tuned_Volume)`t$($row.Base_Volume)")
        }
        [System.IO.File]::WriteAllLines($purgeDictPath, $lines, (New-Object System.Text.UTF8Encoding($true)))
        Write-Log "Save-PurgeDictionary: success"
        return $true
    } catch {
        Write-Log "Save-PurgeDictionary: FAILED - $($_.Exception.Message)" "ERROR"
        Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"
        return $false
    }
}

function Load-PurgeDictionary {
    $script:PurgeDict = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
    if (-not (Test-Path $purgeDictPath)) { Write-Log "Load-PurgeDictionary: file not found at $purgeDictPath" "WARN"; return }
    $lines = Get-Content -Path $purgeDictPath
    if ($lines.Count -eq 0) { return }

    # Accept either delimiter on read (older comma-saved copies may still be around) but
    # always rewrite as tab-delimited â€” that's the format Excel actually writes back to
    # this .csv on save on this machine, so the on-disk format stays consistent.
    $delim = if ($lines[0] -match "`t") { "`t" } else { "," }
    $header = $lines[0] -split $delim
    $colIndex = @{}
    for ($i = 0; $i -lt $header.Count; $i++) { $colIndex[$header[$i].Trim()] = $i }
    # Old format carried Test_Volume/Iterations columns we no longer need; migrate them away on load.
    # Also migrate any leftover comma-delimited copy to tab-delimited.
    $needsMigration = ($delim -eq ",") -or (($colIndex.ContainsKey("Iterations") -or $colIndex.ContainsKey("Test_Volume")) -and (-not $colIndex.ContainsKey("Tuned")))

    foreach ($line in ($lines | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split $delim
        $get = {
            param($col)
            if ($colIndex.ContainsKey($col) -and $colIndex[$col] -lt $parts.Count) { $parts[$colIndex[$col]].Trim() } else { "" }
        }
        $src      = & $get "Source_Filament"
        $tgt      = & $get "Target_Filament"
        $tunedVol = & $get "Tuned_Volume"
        $baseVol  = & $get "Base Volume"

        if ($needsMigration) {
            $tuned = -not [string]::IsNullOrWhiteSpace($tunedVol)
        } else {
            $tunedRaw = & $get "Tuned"
            $tuned = $tunedRaw -match '(?i)^(true|yes|1|x)$'
        }

        $srcHex = if ($script:LibraryColors.Contains($src)) { $script:LibraryColors[$src].Substring(0,7) } else { "#3A3C4A" }
        $tgtHex = if ($script:LibraryColors.Contains($tgt)) { $script:LibraryColors[$tgt].Substring(0,7) } else { "#3A3C4A" }

        $row = New-Object PurgeDictRow
        $row.Load($src, $tgt, $tuned, $tunedVol, $baseVol, $srcHex, $tgtHex)
        $script:PurgeDict.Add($row) | Out-Null
    }

    if ($needsMigration) {
        Write-Log "Load-PurgeDictionary: migrating CSV to drop Test_Volume/Iterations columns"
        Save-PurgeDictionary | Out-Null
    }
    Write-Log "Load-PurgeDictionary: loaded $($script:PurgeDict.Count) entries"
}

Load-PurgeDictionary
function Save-NamesLibrary {
    try {
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.AppendLine("# --- Shared configuration for BambuScripts -----------------------------------")
        [void]$sb.AppendLine("# Single source of truth for grandparent theme names and printer prefixes.")
        [void]$sb.AppendLine("# Dot-source this file from any worker script that needs these values:")
        [void]$sb.AppendLine("#   . (Join-Path `$PSScriptRoot `"..\libraries\NamesLibrary.ps1`")")
        [void]$sb.AppendLine("")
        # Themes â€” wrap every 6 to match original style
        [void]$sb.AppendLine("`$script:GpThemes = @(")
        $chunk = [System.Collections.Generic.List[string]]::new()
        foreach ($t in $script:GpThemes) {
            $chunk.Add("'$t'")
            if ($chunk.Count -eq 6) {
                [void]$sb.AppendLine("    " + ($chunk -join ', ') + ",")
                $chunk.Clear()
            }
        }
        if ($chunk.Count -gt 0) { [void]$sb.AppendLine("    " + ($chunk -join ', ')) }
        [void]$sb.AppendLine(")")
        [void]$sb.AppendLine("")
        # Printer prefixes
        [void]$sb.AppendLine("`$script:PrinterPrefixes = @(" + (($script:PrinterPrefixes | ForEach-Object{"'$_'"}) -join ', ') + ")")
        [void]$sb.AppendLine("")
        # Tags
        [void]$sb.AppendLine("`$script:Tags = @(" + (($script:Tags | ForEach-Object{"'$_'"}) -join ', ') + ")")
        [void]$sb.AppendLine("")
        # TagLabels
        [void]$sb.AppendLine("`$script:TagLabels = @{")
        foreach ($kv in $script:TagLabels.GetEnumerator()) {
            [void]$sb.AppendLine("    '$($kv.Key)'   = '$($kv.Value)'")
        }
        [void]$sb.AppendLine("}")
        [System.IO.File]::WriteAllText($namesLibPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
        return $true
    } catch { return $false }
}

# â”€â”€ Script-scope nav state (used by Set-NavSection after Build-LibrariesPanel returns) â”€â”€
$script:LibNavState  = @{ NavSecs = $null; NavBtns = $null }
$script:TagEditState = @{ TagBox = $null; LblBox = $null; SelectedTag = ""; AddBtn = $null; DirtyTags = @{} }

function Set-NavSection([string]$name) {
    $navSecs = $script:LibNavState.NavSecs; $navBtns = $script:LibNavState.NavBtns
    if ($null -eq $navSecs -or $null -eq $navBtns) { return }
    foreach ($kv in $navSecs.GetEnumerator()) { $kv.Value.Visibility = if ($kv.Key -eq $name) { "Visible" } else { "Collapsed" } }
    foreach ($kv in $navBtns.GetEnumerator()) {
        $kv.Value.Background = if ($kv.Key -eq $name) { Get-WpfColor "#252630" } else { Get-WpfColor "#1C1D23" }
        $kv.Value.Foreground = if ($kv.Key -eq $name) { Get-WpfColor "#FFFFFF" } else { Get-WpfColor "#C0C4D0" }
    }
}

function Update-PickerFromHsv($st) {
    if ($st.Updating) { return }; $st.Updating = $true
    try {
        $rgb = Hsv-To-Rgb $st.H $st.S $st.V
        $r = $rgb[0]; $g = $rgb[1]; $b = $rgb[2]
        $hex6 = "#{0:X2}{1:X2}{2:X2}" -f $r,$g,$b
        $st.Swatch.Background = Get-WpfColor $hex6
        $st.RBox.Text = $r; $st.GBox.Text = $g; $st.BBox.Text = $b
        $st.HexBox.Text = $hex6
        # Hue bar indicator
        $hw = $st.HueCanvas.ActualWidth
        if ($hw -gt 0) {
            $hx = $st.H / 360.0 * $hw - 2
            [System.Windows.Controls.Canvas]::SetLeft($st.HueIndicator, [Math]::Max(0,$hx))
            $st.HueIndicator.Height = $st.HueCanvas.ActualHeight
        }
        # SV square hue stop
        $pRgb = Hsv-To-Rgb $st.H 1.0 1.0
        $st.SvHueStop.Color = [System.Windows.Media.Color]::FromRgb($pRgb[0],$pRgb[1],$pRgb[2])
        # SV indicator + ensure rects are sized (SizeChanged may not have fired yet)
        $sw = $st.SvCanvas.ActualWidth; $sh = $st.SvCanvas.ActualHeight
        if ($sw -gt 0 -and $sh -gt 0) {
            if ($null -ne $st.SvColorRect -and ($st.SvColorRect.Width -ne $sw -or $st.SvColorRect.Height -ne $sh)) {
                $st.SvColorRect.Width = $sw; $st.SvColorRect.Height = $sh
            }
            if ($null -ne $st.SvDarkRect -and ($st.SvDarkRect.Width -ne $sw -or $st.SvDarkRect.Height -ne $sh)) {
                $st.SvDarkRect.Width = $sw; $st.SvDarkRect.Height = $sh
            }
            [System.Windows.Controls.Canvas]::SetLeft($st.SvDot, $st.S * $sw - 6)
            [System.Windows.Controls.Canvas]::SetTop($st.SvDot,  (1-$st.V) * $sh - 6)
        }
    } finally { $st.Updating = $false }
}

function Update-PickerFromRgb($st) {
    $rV=0;$gV=0;$bV=0
    [int]::TryParse($st.RBox.Text,[ref]$rV)|Out-Null
    [int]::TryParse($st.GBox.Text,[ref]$gV)|Out-Null
    [int]::TryParse($st.BBox.Text,[ref]$bV)|Out-Null
    $rV=[Math]::Max(0,[Math]::Min(255,$rV)); $gV=[Math]::Max(0,[Math]::Min(255,$gV)); $bV=[Math]::Max(0,[Math]::Min(255,$bV))
    $hsv = Rgb-To-Hsv $rV $gV $bV
    $st.H=$hsv[0]; $st.S=$hsv[1]; $st.V=$hsv[2]
    Update-PickerFromHsv $st
}

function Load-FilamentEntry([string]$entryName, $st) {
    Write-Log "Load-FilamentEntry: '$entryName'"
    $hex9 = $script:LibraryColors[$entryName]
    if (-not $hex9) { Write-Log "Load-FilamentEntry: not found in library" "WARN"; return }
    $r=[Convert]::ToInt32($hex9.Substring(1,2),16)
    $g=[Convert]::ToInt32($hex9.Substring(3,2),16)
    $b=[Convert]::ToInt32($hex9.Substring(5,2),16)
    $st.NameBox.Text = $entryName; $st.EditingName = $entryName
    # Suppress TextChanged cascade while batch-setting all three boxes
    $st.Updating = $true
    $st.RBox.Text = "$r"; $st.GBox.Text = "$g"; $st.BBox.Text = "$b"
    $st.Updating = $false
    Update-PickerFromRgb $st   # positions the HSV picker visuals
    # Restore exact library values - the HSV round-trip in Update-PickerFromRgb can
    # introduce +/-1 rounding drift on any channel, so we pin the boxes back to the
    # canonical values.  Use Updating=true to suppress a second TextChanged cascade.
    $st.Updating = $true
    $st.RBox.Text = "$r"; $st.GBox.Text = "$g"; $st.BBox.Text = "$b"
    $st.HexBox.Text = $hex9.Substring(0, 7)
    $st.Updating = $false
    $st.StatusLbl.Text = "Loaded: $entryName"; $st.StatusLbl.Foreground = Get-WpfColor "#555868"
}

function Rebuild-FilamentList($st) {
    Write-Log "Rebuild-FilamentList: start"
    $st.FilamentStack.Children.Clear()
    foreach ($kv in $script:LibraryColors.GetEnumerator()) {
        $eName = $kv.Key; $eHex = $kv.Value
        $row = New-Object System.Windows.Controls.Border
        $row.Padding = New-Object System.Windows.Thickness(8,5,8,5)
        $row.BorderBrush = Get-WpfColor "#23242E"; $row.BorderThickness = New-Object System.Windows.Thickness(0,0,0,1)
        $row.Cursor = [System.Windows.Input.Cursors]::Hand
        $row.Background = Get-WpfColor "#0D0E10"

        $rowStack = New-Object System.Windows.Controls.StackPanel; $rowStack.Orientation = "Horizontal"
        $swatch = New-Object System.Windows.Controls.Border
        $swatch.Width = 18; $swatch.Height = 18; $swatch.CornerRadius = New-Object System.Windows.CornerRadius(3)
        $swatch.Background = Get-WpfColor ($eHex.Substring(0,7))
        $swatch.Margin = New-Object System.Windows.Thickness(0,0,8,0); $swatch.VerticalAlignment = "Center"
        $rowStack.Children.Add($swatch) | Out-Null
        $nameTb = New-Object System.Windows.Controls.TextBlock; $nameTb.Text = $eName
        $nameTb.Foreground = Get-WpfColor "#C0C4D0"; $nameTb.FontSize = 11; $nameTb.VerticalAlignment = "Center"
        $rowStack.Children.Add($nameTb) | Out-Null
        $row.Child = $rowStack

        $capturedName = $eName; $capturedState = $st
        $row.Add_MouseEnter({ $this.Background = Get-WpfColor "#1A1C24" })
        $row.Add_MouseLeave({ $this.Background = Get-WpfColor "#0D0E10" })
        $row.Add_MouseLeftButtonDown({
            try {
                Write-Log "FilamentRow.MouseDown: '$capturedName'"
                Load-FilamentEntry $capturedName $capturedState
            } catch { Write-Log "FilamentRow.MouseDown EXCEPTION: $($_.Exception.Message)" "ERROR" }
        }.GetNewClosure())
        $st.FilamentStack.Children.Add($row) | Out-Null
    }
    Write-Log "Rebuild-FilamentList: done ($($script:LibraryColors.Count) rows)"
}

function Rebuild-TagsList($stack) {
    # Remove all rows except the header (index 0)
    while ($stack.Children.Count -gt 1) { $stack.Children.RemoveAt(1) }
    # Capture $script: variables as locals â€” $script: scope is NOT reliably accessible
    # inside .GetNewClosure() closures running on the WPF dispatcher thread.
    $tes       = $script:TagEditState
    $tagLabels = $script:TagLabels
    foreach ($tag in $script:Tags) {
        $lbl = if ($tagLabels.ContainsKey($tag)) { $tagLabels[$tag] } else { "" }
        $row = New-Object System.Windows.Controls.Grid; $row.Margin = New-Object System.Windows.Thickness(0,1,0,1)
        $row.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(80) })) | Out-Null
        $row.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
        $isDirty = $tes.DirtyTags.ContainsKey($tag)
        $fgTag = if ($isDirty) { Get-WpfColor "#E8903A" } else { Get-WpfColor "#E0E3EC" }
        $fgLbl = if ($isDirty) { Get-WpfColor "#C07030" } else { Get-WpfColor "#C0C4D0" }
        $tbTag = New-Object System.Windows.Controls.TextBlock; $tbTag.Text=$tag; $tbTag.Foreground=$fgTag; $tbTag.FontSize=12; $tbTag.VerticalAlignment="Center"; [System.Windows.Controls.Grid]::SetColumn($tbTag,0)
        $tbLbl = New-Object System.Windows.Controls.TextBlock; $tbLbl.Text=$lbl; $tbLbl.Foreground=$fgLbl; $tbLbl.FontSize=12; $tbLbl.VerticalAlignment="Center"; [System.Windows.Controls.Grid]::SetColumn($tbLbl,1)
        $row.Children.Add($tbTag)|Out-Null; $row.Children.Add($tbLbl)|Out-Null
        $row.Cursor = [System.Windows.Input.Cursors]::Hand
        # Capture loop variables and stack explicitly â€” never rely on $this.Tag inside closures
        $capturedTag   = $tag
        $capturedStack = $stack
        $capturedRow   = $row
        $row.Add_MouseEnter({
            if ($capturedTag -ne $tes.SelectedTag) { $capturedRow.Background = Get-WpfColor "#252630" }
        }.GetNewClosure())
        $row.Add_MouseLeave({
            if ($capturedTag -ne $tes.SelectedTag) { $capturedRow.Background = $null }
        }.GetNewClosure())
        $row.Add_MouseLeftButtonDown({
            if ($tes.SelectedTag -eq $capturedTag) {
                # Toggle off â€” clicking selected row again deselects it
                $capturedRow.Background = $null
                $tes.SelectedTag = ""
                if ($null -ne $tes.TagBox) { $tes.TagBox.Text = "" }
                if ($null -ne $tes.LblBox) { $tes.LblBox.Text = "" }
                if ($null -ne $tes.AddBtn) { $tes.AddBtn.Content = "Add" }
            } else {
                # Select this row, clear all others
                foreach ($sib in $capturedStack.Children) {
                    if ($sib -is [System.Windows.FrameworkElement] -and
                        $null -ne $sib.Tag -and $sib.Tag -is [string]) {
                        $sib.Background = $null
                    }
                }
                $capturedRow.Background = Get-WpfColor "#2A3A5A"
                $tes.SelectedTag = $capturedTag
                if ($null -ne $tes.TagBox) {
                    $tes.TagBox.Text = $capturedTag
                    $tes.LblBox.Text = if ($tagLabels.ContainsKey($capturedTag)) { $tagLabels[$capturedTag] } else { "" }
                }
                if ($null -ne $tes.AddBtn) { $tes.AddBtn.Content = "Edit" }
            }
        }.GetNewClosure())
        # Use a simple string tag so sibling-clear loop works without hashtable method calls
        $row.Tag = $capturedTag
        $stack.Children.Add($row) | Out-Null
    }
}

# â”€â”€ Script-scope helpers for Naming Conventions closures â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# These functions are called BY NAME from inside .GetNewClosure() closures, so
# they execute in the main script's scope and can safely read/write $script: vars.
function Set-GpThemes([string[]]$items)        { $script:GpThemes        = $items }
function Set-PrinterPrefixes([string[]]$items) { $script:PrinterPrefixes = $items }

function Invoke-AddTagEntry([string]$tn, [string]$tl, $stack) {
    Write-Log "Invoke-AddTagEntry: tn='$tn' tl='$tl'"
    $tes = $script:TagEditState
    if ($null -eq $tes) { Write-Log "Invoke-AddTagEntry: TagEditState is null" "ERROR"; return }
    $oldSel = $tes.SelectedTag
    if (-not [string]::IsNullOrWhiteSpace($oldSel) -and $oldSel -ne $tn -and ($script:Tags -contains $oldSel)) {
        $script:Tags = @($script:Tags | ForEach-Object { if ($_ -eq $oldSel) { $tn } else { $_ } })
        $script:TagLabels.Remove($oldSel) | Out-Null
        $tes.DirtyTags.Remove($oldSel) | Out-Null
    } elseif (-not ($script:Tags -contains $tn)) {
        $script:Tags += $tn
    }
    $script:TagLabels[$tn] = $tl
    $tes.DirtyTags[$tn] = $true
    $tes.SelectedTag = ""
    if ($null -ne $tes.TagBox) { $tes.TagBox.Text = "" }
    if ($null -ne $tes.LblBox) { $tes.LblBox.Text = "" }
    if ($null -ne $tes.AddBtn) { $tes.AddBtn.Content = "Add" }
    Rebuild-TagsList $stack
    Write-Log "Invoke-AddTagEntry: done (Tags=$($script:Tags.Count))"
}

function Invoke-RemoveTagEntry([string]$tn, $stack) {
    Write-Log "Invoke-RemoveTagEntry: tn='$tn'"
    $tes = $script:TagEditState
    if ($null -eq $tes) { Write-Log "Invoke-RemoveTagEntry: TagEditState is null" "ERROR"; return }
    $script:Tags = @($script:Tags | Where-Object { $_ -ne $tn })
    $script:TagLabels.Remove($tn) | Out-Null
    $tes.DirtyTags.Remove($tn) | Out-Null
    $tes.SelectedTag = ""
    if ($null -ne $tes.TagBox) { $tes.TagBox.Text = "" }
    if ($null -ne $tes.LblBox) { $tes.LblBox.Text = "" }
    if ($null -ne $tes.AddBtn) { $tes.AddBtn.Content = "Add" }
    Rebuild-TagsList $stack
    Write-Log "Invoke-RemoveTagEntry: done (Tags=$($script:Tags.Count))"
}

function Invoke-SaveTagsSection($stack) {
    Write-Log "Invoke-SaveTagsSection: start"
    if (Save-NamesLibrary) {
        $script:TagEditState.DirtyTags = @{}
        Rebuild-TagsList $stack
        Write-Log "Invoke-SaveTagsSection: success"
        return $true
    }
    Write-Log "Invoke-SaveTagsSection: Save-NamesLibrary returned false" "ERROR"
    return $false
}

function Build-LibrariesPanel {
    # â”€â”€ Root panel: overlaps the ScrollViewer in Grid Row 1 â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $outerGrid = $window.Content
    $root = New-Object System.Windows.Controls.Grid
    $root.Background = Get-WpfColor "#0D0E10"
    $root.Visibility = "Collapsed"
    [System.Windows.Controls.Grid]::SetRow($root, 1)
    $outerGrid.Children.Add($root) | Out-Null

    $root.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(170) })) | Out-Null
    $root.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null

    # â”€â”€ Left nav sidebar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $sidebar = New-Object System.Windows.Controls.Border
    $sidebar.Background = Get-WpfColor "#1C1D23"
    $sidebar.BorderBrush = Get-WpfColor "#2A2C38"; $sidebar.BorderThickness = New-Object System.Windows.Thickness(0,0,1,0)
    [System.Windows.Controls.Grid]::SetColumn($sidebar, 0)
    $root.Children.Add($sidebar) | Out-Null

    $sideStack = New-Object System.Windows.Controls.StackPanel
    $sideStack.Margin = New-Object System.Windows.Thickness(0,12,0,0)
    $sidebar.Child = $sideStack

    $libHdr = New-Object System.Windows.Controls.TextBlock
    $libHdr.Text = "LIBRARIES"; $libHdr.FontSize = 10; $libHdr.FontWeight = [System.Windows.FontWeights]::Bold
    $libHdr.Foreground = Get-WpfColor "#555868"; $libHdr.Margin = New-Object System.Windows.Thickness(14,0,0,10)
    $sideStack.Children.Add($libHdr) | Out-Null

    function New-NavBtn([string]$label) {
        $b = New-Object System.Windows.Controls.Button
        $b.Content = $label; $b.Height = 34; $b.FontSize = 12
        $b.Background = Get-WpfColor "#1C1D23"; $b.Foreground = Get-WpfColor "#C0C4D0"
        $b.BorderThickness = 0; $b.Cursor = [System.Windows.Input.Cursors]::Hand
        $b.HorizontalContentAlignment = "Left"
        $b.Padding = New-Object System.Windows.Thickness(14,0,0,0)
        return $b
    }
    $btnNavFilaments = New-NavBtn "Filaments"
    $btnNavNaming    = New-NavBtn "Naming Conventions"
    $btnNavPurge     = New-NavBtn "Purge Dictionary"
    $sideStack.Children.Add($btnNavFilaments) | Out-Null
    $sideStack.Children.Add($btnNavNaming)    | Out-Null
    $sideStack.Children.Add($btnNavPurge)     | Out-Null

    # â”€â”€ Content area â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $contentGrid = New-Object System.Windows.Controls.Grid
    [System.Windows.Controls.Grid]::SetColumn($contentGrid, 1)
    $root.Children.Add($contentGrid) | Out-Null

    # Two overlapping section panels â€” show one at a time
    $secFilaments = New-Object System.Windows.Controls.Grid; $secFilaments.Background = Get-WpfColor "#0D0E10"; $secFilaments.Visibility = "Visible"
    $secNaming    = New-Object System.Windows.Controls.Grid; $secNaming.Background    = Get-WpfColor "#0D0E10"; $secNaming.Visibility    = "Collapsed"
    $secPurge     = New-Object System.Windows.Controls.Grid; $secPurge.Background     = Get-WpfColor "#0D0E10"; $secPurge.Visibility     = "Collapsed"
    $contentGrid.Children.Add($secFilaments) | Out-Null
    $contentGrid.Children.Add($secNaming)    | Out-Null
    $contentGrid.Children.Add($secPurge)     | Out-Null

    # Nav button click logic â€” populate script-scope LibNavState so Set-NavSection works after this function returns
    $navBtns = @{ Filaments=$btnNavFilaments; "Naming Conventions"=$btnNavNaming; "Purge Dictionary"=$btnNavPurge }
    $navSecs = @{ Filaments=$secFilaments;    "Naming Conventions"=$secNaming;    "Purge Dictionary"=$secPurge }
    $script:LibNavState.NavBtns = $navBtns
    $script:LibNavState.NavSecs = $navSecs
    Set-NavSection "Filaments"
    $btnNavFilaments.Add_Click({ Write-Log "btnNavFilaments: clicked"; Set-NavSection "Filaments"          })
    $btnNavNaming.Add_Click({    Write-Log "btnNavNaming: clicked";    Set-NavSection "Naming Conventions" })
    $btnNavPurge.Add_Click({     Write-Log "btnNavPurge: clicked";     Set-NavSection "Purge Dictionary"   })

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  FILAMENTS SECTION
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    $secFilaments.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(260) })) | Out-Null
    $secFilaments.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null

    # â”€â”€ Filament list (left) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $listBorder = New-Object System.Windows.Controls.Border
    $listBorder.BorderBrush = Get-WpfColor "#2A2C38"; $listBorder.BorderThickness = New-Object System.Windows.Thickness(0,0,1,0)
    [System.Windows.Controls.Grid]::SetColumn($listBorder, 0)
    $secFilaments.Children.Add($listBorder) | Out-Null

    $listDock = New-Object System.Windows.Controls.DockPanel
    $listBorder.Child = $listDock

    # Bottom: Add New button
    $listFooter = New-Object System.Windows.Controls.Border
    $listFooter.Background = Get-WpfColor "#1C1D23"; $listFooter.BorderBrush = Get-WpfColor "#2A2C38"
    $listFooter.BorderThickness = New-Object System.Windows.Thickness(0,1,0,0); $listFooter.Padding = New-Object System.Windows.Thickness(8,6,8,6)
    [System.Windows.Controls.DockPanel]::SetDock($listFooter, "Bottom")
    $listDock.Children.Add($listFooter) | Out-Null

    $btnAddNew = New-Object System.Windows.Controls.Button
    $btnAddNew.Content = "+ Add New Filament"; $btnAddNew.Height = 28; $btnAddNew.FontSize = 11
    $btnAddNew.Background = Get-WpfColor "#2E5A42"; $btnAddNew.Foreground = Get-WpfColor "#FFFFFF"
    $btnAddNew.BorderThickness = 0; $btnAddNew.Cursor = [System.Windows.Input.Cursors]::Hand
    $listFooter.Child = $btnAddNew

    $listScroll = New-Object System.Windows.Controls.ScrollViewer
    $listScroll.VerticalScrollBarVisibility = "Auto"
    $listScroll.HorizontalScrollBarVisibility = "Disabled"
    $listDock.Children.Add($listScroll) | Out-Null

    $filamentStack = New-Object System.Windows.Controls.StackPanel
    $listScroll.Content = $filamentStack

    # â”€â”€ Edit panel (right) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $editScroll = New-Object System.Windows.Controls.ScrollViewer
    $editScroll.VerticalScrollBarVisibility = "Auto"; $editScroll.HorizontalScrollBarVisibility = "Disabled"
    $editScroll.Padding = New-Object System.Windows.Thickness(20,16,20,20)
    [System.Windows.Controls.Grid]::SetColumn($editScroll, 1)
    $secFilaments.Children.Add($editScroll) | Out-Null

    $editStack = New-Object System.Windows.Controls.StackPanel
    $editScroll.Content = $editStack

    $editHdr = New-Object System.Windows.Controls.TextBlock
    $editHdr.Text = "Edit Filament"; $editHdr.FontSize = 14; $editHdr.FontWeight = [System.Windows.FontWeights]::Bold
    $editHdr.Foreground = Get-WpfColor "#C8CFDD"; $editHdr.Margin = New-Object System.Windows.Thickness(0,0,0,14)
    $editStack.Children.Add($editHdr) | Out-Null

    # Name field
    $nameLbl = New-Object System.Windows.Controls.TextBlock; $nameLbl.Text = "Name"
    $nameLbl.Foreground = Get-WpfColor "#888A9A"; $nameLbl.FontSize = 11; $nameLbl.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $editStack.Children.Add($nameLbl) | Out-Null
    $tbName = New-Object System.Windows.Controls.TextBox; $tbName.Height = 28; $tbName.FontSize = 12
    $tbName.Background = Get-WpfColor "#1C1D23"; $tbName.Foreground = Get-WpfColor "#E0E3EC"
    $tbName.BorderBrush = Get-WpfColor "#3A3C4A"; $tbName.BorderThickness = New-Object System.Windows.Thickness(1)
    $tbName.Padding = New-Object System.Windows.Thickness(6,0,6,0); $tbName.Margin = New-Object System.Windows.Thickness(0,0,0,14)
    $editStack.Children.Add($tbName) | Out-Null

    # Color swatch
    $swatchBorder = New-Object System.Windows.Controls.Border
    $swatchBorder.Height = 48; $swatchBorder.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $swatchBorder.Background = Get-WpfColor "#FF0000"; $swatchBorder.Margin = New-Object System.Windows.Thickness(0,0,0,14)
    $editStack.Children.Add($swatchBorder) | Out-Null

    # Hue bar  â”€â”€ rainbow gradient Canvas + indicator
    $hueLbl = New-Object System.Windows.Controls.TextBlock; $hueLbl.Text = "Hue"
    $hueLbl.Foreground = Get-WpfColor "#888A9A"; $hueLbl.FontSize = 11; $hueLbl.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $editStack.Children.Add($hueLbl) | Out-Null

    $hueOuter = New-Object System.Windows.Controls.Border
    $hueOuter.Height = 24; $hueOuter.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $hueOuter.Margin = New-Object System.Windows.Thickness(0,0,0,10); $hueOuter.ClipToBounds = $true
    $editStack.Children.Add($hueOuter) | Out-Null

    $hueCanvas = New-Object System.Windows.Controls.Canvas
    $hueOuter.Child = $hueCanvas

    $hueBrush = New-Object System.Windows.Media.LinearGradientBrush
    $hueBrush.StartPoint = [System.Windows.Point]::new(0,0.5); $hueBrush.EndPoint = [System.Windows.Point]::new(1,0.5)
    foreach ($pair in @(@(0,"#FF0000"),@(0.167,"#FFFF00"),@(0.333,"#00FF00"),@(0.5,"#00FFFF"),@(0.667,"#0000FF"),@(0.833,"#FF00FF"),@(1,"#FF0000"))) {
        $gs = New-Object System.Windows.Media.GradientStop
        $gs.Color  = [System.Windows.Media.ColorConverter]::ConvertFromString($pair[1])
        $gs.Offset = $pair[0]
        $hueBrush.GradientStops.Add($gs) | Out-Null
    }
    $hueRect = New-Object System.Windows.Shapes.Rectangle
    $hueRect.Fill = $hueBrush; $hueRect.Height = 24
    [System.Windows.Controls.Canvas]::SetLeft($hueRect, 0); [System.Windows.Controls.Canvas]::SetTop($hueRect, 0)
    $hueCanvas.Children.Add($hueRect) | Out-Null

    # Hue position indicator (thin white rect with dark border)
    $hueInd = New-Object System.Windows.Controls.Border
    $hueInd.Width = 4; $hueInd.Background = Get-WpfColor "#FFFFFF"
    $hueInd.BorderBrush = Get-WpfColor "#000000"; $hueInd.BorderThickness = New-Object System.Windows.Thickness(1)
    [System.Windows.Controls.Canvas]::SetTop($hueInd, 0)
    $hueCanvas.Children.Add($hueInd) | Out-Null

    # Hue canvas needs to stretch â€” bind Width via SizeChanged and Loaded
    $sizeHueRect = { if ($hueCanvas.ActualWidth -gt 0) { $hueRect.Width = $hueCanvas.ActualWidth } }.GetNewClosure()
    $hueCanvas.Add_SizeChanged($sizeHueRect)
    $hueCanvas.Add_Loaded($sizeHueRect)

    # SV (saturation/value) square
    $svLbl = New-Object System.Windows.Controls.TextBlock; $svLbl.Text = "Saturation / Brightness"
    $svLbl.Foreground = Get-WpfColor "#888A9A"; $svLbl.FontSize = 11; $svLbl.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $editStack.Children.Add($svLbl) | Out-Null

    $svOuter = New-Object System.Windows.Controls.Border
    $svOuter.Height = 200; $svOuter.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $svOuter.Margin = New-Object System.Windows.Thickness(0,0,0,14); $svOuter.ClipToBounds = $true
    $editStack.Children.Add($svOuter) | Out-Null

    $svCanvas = New-Object System.Windows.Controls.Canvas
    $svCanvas.Background = [System.Windows.Media.Brushes]::Transparent  # needed for hit-testing
    $svOuter.Child = $svCanvas

    # Layer 1: horizontal whiteâ†’hue gradient
    $svHueBrush = New-Object System.Windows.Media.LinearGradientBrush
    $svHueBrush.StartPoint = [System.Windows.Point]::new(0,0.5); $svHueBrush.EndPoint = [System.Windows.Point]::new(1,0.5)
    $svStopWhite = New-Object System.Windows.Media.GradientStop; $svStopWhite.Color = [System.Windows.Media.Colors]::White; $svStopWhite.Offset = 0
    $svStopHue   = New-Object System.Windows.Media.GradientStop; $svStopHue.Color   = [System.Windows.Media.Colors]::Red;   $svStopHue.Offset   = 1
    $svHueBrush.GradientStops.Add($svStopWhite) | Out-Null; $svHueBrush.GradientStops.Add($svStopHue) | Out-Null
    $svColorRect = New-Object System.Windows.Shapes.Rectangle; $svColorRect.Fill = $svHueBrush
    [System.Windows.Controls.Canvas]::SetLeft($svColorRect,0); [System.Windows.Controls.Canvas]::SetTop($svColorRect,0)
    $svCanvas.Children.Add($svColorRect) | Out-Null

    # Layer 2: vertical transparentâ†’black gradient
    $svDarkBrush = New-Object System.Windows.Media.LinearGradientBrush
    $svDarkBrush.StartPoint = [System.Windows.Point]::new(0.5,0); $svDarkBrush.EndPoint = [System.Windows.Point]::new(0.5,1)
    $svStopTrans = New-Object System.Windows.Media.GradientStop; $svStopTrans.Color = [System.Windows.Media.Color]::FromArgb(0,0,0,0); $svStopTrans.Offset = 0
    $svStopBlack = New-Object System.Windows.Media.GradientStop; $svStopBlack.Color = [System.Windows.Media.Colors]::Black; $svStopBlack.Offset = 1
    $svDarkBrush.GradientStops.Add($svStopTrans) | Out-Null; $svDarkBrush.GradientStops.Add($svStopBlack) | Out-Null
    $svDarkRect = New-Object System.Windows.Shapes.Rectangle; $svDarkRect.Fill = $svDarkBrush
    [System.Windows.Controls.Canvas]::SetLeft($svDarkRect,0); [System.Windows.Controls.Canvas]::SetTop($svDarkRect,0)
    $svCanvas.Children.Add($svDarkRect) | Out-Null

    # Crosshair indicator
    $svDot = New-Object System.Windows.Shapes.Ellipse
    $svDot.Width = 12; $svDot.Height = 12
    $svDot.Stroke = Get-WpfColor "#FFFFFF"; $svDot.StrokeThickness = 2
    $svDot.Fill = [System.Windows.Media.Brushes]::Transparent
    $svDot.IsHitTestVisible = $false
    $svCanvas.Children.Add($svDot) | Out-Null

    # Size the gradient rects whenever the canvas renders or resizes
    $sizeSvRects = {
        $w = $svCanvas.ActualWidth; $h = $svCanvas.ActualHeight
        if ($w -gt 0 -and $h -gt 0) {
            $svColorRect.Width = $w; $svColorRect.Height = $h
            $svDarkRect.Width  = $w; $svDarkRect.Height  = $h
        }
    }.GetNewClosure()
    $svCanvas.Add_SizeChanged($sizeSvRects)
    $svCanvas.Add_Loaded($sizeSvRects)

    # RGB boxes
    $rgbLbl = New-Object System.Windows.Controls.TextBlock; $rgbLbl.Text = "R / G / B  (0-255)"
    $rgbLbl.Foreground = Get-WpfColor "#888A9A"; $rgbLbl.FontSize = 11; $rgbLbl.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $editStack.Children.Add($rgbLbl) | Out-Null

    $rgbGrid = New-Object System.Windows.Controls.Grid; $rgbGrid.Margin = New-Object System.Windows.Thickness(0,0,0,10)
    foreach ($w in @(4,1,4,1,4)) { $rgbGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = if($w-eq 1){[System.Windows.GridLength]::new(8)}else{[System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)} })) | Out-Null }
    $editStack.Children.Add($rgbGrid) | Out-Null

    function New-RgbBox {
        $tb = New-Object System.Windows.Controls.TextBox; $tb.Height = 30; $tb.FontSize = 13; $tb.TextAlignment = "Center"
        $tb.Background = Get-WpfColor "#1C1D23"; $tb.Foreground = Get-WpfColor "#E0E3EC"
        $tb.BorderBrush = Get-WpfColor "#3A3C4A"; $tb.BorderThickness = New-Object System.Windows.Thickness(1)
        return $tb
    }
    $tbR = New-RgbBox; $tbG = New-RgbBox; $tbB = New-RgbBox
    [System.Windows.Controls.Grid]::SetColumn($tbR,0); [System.Windows.Controls.Grid]::SetColumn($tbG,2); [System.Windows.Controls.Grid]::SetColumn($tbB,4)
    $rgbGrid.Children.Add($tbR) | Out-Null; $rgbGrid.Children.Add($tbG) | Out-Null; $rgbGrid.Children.Add($tbB) | Out-Null

    $hexLbl = New-Object System.Windows.Controls.TextBlock; $hexLbl.Text = "Hex"
    $hexLbl.Foreground = Get-WpfColor "#888A9A"; $hexLbl.FontSize = 11; $hexLbl.Margin = New-Object System.Windows.Thickness(0,0,0,4)
    $editStack.Children.Add($hexLbl) | Out-Null
    $tbHex = New-Object System.Windows.Controls.TextBox; $tbHex.Height = 28; $tbHex.FontSize = 12
    $tbHex.Background = Get-WpfColor "#1C1D23"; $tbHex.Foreground = Get-WpfColor "#888A9A"
    $tbHex.BorderBrush = Get-WpfColor "#2A2C38"; $tbHex.BorderThickness = New-Object System.Windows.Thickness(1)
    $tbHex.Padding = New-Object System.Windows.Thickness(6,0,6,0); $tbHex.IsReadOnly = $true
    $tbHex.Margin = New-Object System.Windows.Thickness(0,0,0,18)
    $editStack.Children.Add($tbHex) | Out-Null

    # Save / Delete buttons
    $saveBtnGrid = New-Object System.Windows.Controls.Grid; $saveBtnGrid.Margin = New-Object System.Windows.Thickness(0,0,0,8)
    $saveBtnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $saveBtnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(10) })) | Out-Null
    $saveBtnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $editStack.Children.Add($saveBtnGrid) | Out-Null

    $btnSaveFilament = New-Object System.Windows.Controls.Button; $btnSaveFilament.Content = "Save / Update"
    $btnSaveFilament.Height = 32; $btnSaveFilament.FontSize = 12; $btnSaveFilament.FontWeight = [System.Windows.FontWeights]::Bold
    $btnSaveFilament.Background = Get-WpfColor "#3A5080"; $btnSaveFilament.Foreground = Get-WpfColor "#FFFFFF"
    $btnSaveFilament.BorderThickness = 0; $btnSaveFilament.Cursor = [System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnSaveFilament, 0); $saveBtnGrid.Children.Add($btnSaveFilament) | Out-Null

    $btnDelFilament = New-Object System.Windows.Controls.Button; $btnDelFilament.Content = "Delete"
    $btnDelFilament.Height = 32; $btnDelFilament.FontSize = 12
    $btnDelFilament.Background = Get-WpfColor "#6B2828"; $btnDelFilament.Foreground = Get-WpfColor "#FFFFFF"
    $btnDelFilament.BorderThickness = 0; $btnDelFilament.Cursor = [System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnDelFilament, 2); $saveBtnGrid.Children.Add($btnDelFilament) | Out-Null

    $libStatusLbl = New-Object System.Windows.Controls.TextBlock
    $libStatusLbl.FontSize = 11; $libStatusLbl.Foreground = Get-WpfColor "#555868"
    $libStatusLbl.Margin = New-Object System.Windows.Thickness(0,4,0,0)
    $editStack.Children.Add($libStatusLbl) | Out-Null

    # Capture script-scope collections as locals â€” $script: scope is NOT reliably
    # accessible inside .GetNewClosure() closures running on the WPF dispatcher thread.
    $capturedLibColors = $script:LibraryColors
    $capturedHexToName = $script:HexToName
    $capturedLibGradients = $script:LibraryGradients

    # â”€â”€ Shared picker state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $pickerState = @{
        H = 0.0; S = 1.0; V = 1.0
        Updating = $false
        EditingName = ""     # name of the entry currently loaded ("" = new)
        HueCanvas = $hueCanvas; HueIndicator = $hueInd
        SvCanvas  = $svCanvas;  SvDot = $svDot; SvHueStop = $svStopHue
        SvColorRect = $svColorRect; SvDarkRect = $svDarkRect
        Swatch    = $swatchBorder
        RBox = $tbR; GBox = $tbG; BBox = $tbB; HexBox = $tbHex
        NameBox = $tbName
        FilamentStack = $filamentStack
        StatusLbl = $libStatusLbl
        SaveBtn = $btnSaveFilament; DelBtn = $btnDelFilament
    }

    # â”€â”€ Hue canvas mouse events â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Use captured canvas/pickerState variables â€” avoids $this/$_ capture ambiguity
    # from .GetNewClosure(). WPF passes (sender, eventArgs) as $args[0]/$args[1];
    # those are never captured and are safe to use for position queries.
    $hueCanvas.Add_MouseLeftButtonDown({
        try {
            $hueCanvas.CaptureMouse()
            $w = $hueCanvas.ActualWidth; if ($w -le 0) { return }
            $pickerState.H = [Math]::Max(0.0,[Math]::Min(359.9, $args[1].GetPosition($hueCanvas).X / $w * 360.0))
            Update-PickerFromHsv $pickerState
        } catch { Write-Log "hueCanvas.MouseDown EXCEPTION: $($_.Exception.Message)" "ERROR" }
    }.GetNewClosure())
    $hueCanvas.Add_MouseMove({
        try {
            if (-not $hueCanvas.IsMouseCaptured) { return }
            $w = $hueCanvas.ActualWidth; if ($w -le 0) { return }
            $pickerState.H = [Math]::Max(0.0,[Math]::Min(359.9, $args[1].GetPosition($hueCanvas).X / $w * 360.0))
            Update-PickerFromHsv $pickerState
        } catch { Write-Log "hueCanvas.MouseMove EXCEPTION: $($_.Exception.Message)" "ERROR" }
    }.GetNewClosure())
    $hueCanvas.Add_MouseLeftButtonUp({ $hueCanvas.ReleaseMouseCapture() }.GetNewClosure())

    # â”€â”€ SV canvas mouse events â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Uses $svOuter (Border) for hit-testing and size â€” more reliable than the
    # Canvas whose Background was null (non-hit-testable) before children sized.
    # Position via Mouse.GetPosition to avoid any $args capture ambiguity.
    # IMPORTANT: use 0.0/1.0 double literals in Min/Max â€” PowerShell selects the
    # Int32 overload when the first arg is an integer literal, which rounds the
    # double result to 0 or 1 and causes the "locked to quadrants" behaviour.
    $svOuter.Add_MouseLeftButtonDown({
        try {
            $svOuter.CaptureMouse()
            $w = $svOuter.ActualWidth; $h = $svOuter.ActualHeight
            if ($w -le 0 -or $h -le 0) { return }
            $p = [System.Windows.Input.Mouse]::GetPosition($svOuter)
            $pickerState.S = [Math]::Max(0.0,[Math]::Min(1.0, $p.X / $w))
            $pickerState.V = [Math]::Max(0.0,[Math]::Min(1.0, 1.0 - $p.Y / $h))
            Update-PickerFromHsv $pickerState
        } catch { Write-Log "svOuter.MouseDown EXCEPTION: $($_.Exception.Message)" "ERROR" }
    }.GetNewClosure())
    $svOuter.Add_MouseMove({
        try {
            if (-not $svOuter.IsMouseCaptured) { return }
            $w = $svOuter.ActualWidth; $h = $svOuter.ActualHeight
            if ($w -le 0 -or $h -le 0) { return }
            $p = [System.Windows.Input.Mouse]::GetPosition($svOuter)
            $pickerState.S = [Math]::Max(0.0,[Math]::Min(1.0, $p.X / $w))
            $pickerState.V = [Math]::Max(0.0,[Math]::Min(1.0, 1.0 - $p.Y / $h))
            Update-PickerFromHsv $pickerState
        } catch { Write-Log "svOuter.MouseMove EXCEPTION: $($_.Exception.Message)" "ERROR" }
    }.GetNewClosure())
    $svOuter.Add_MouseLeftButtonUp({ $svOuter.ReleaseMouseCapture() }.GetNewClosure())

    # â”€â”€ RGB box TextChanged â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    foreach ($box in @($tbR,$tbG,$tbB)) {
        $box.Add_TextChanged({
            if ($pickerState.Updating) { return }
            Update-PickerFromRgb $pickerState
        }.GetNewClosure())
    }

    # â”€â”€ Save / Update â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $btnSaveFilament.Add_Click({
        try {
            Write-Log "btnSaveFilament: clicked"
            $nm = $pickerState.NameBox.Text.Trim()
            Write-Log "btnSaveFilament: name='$nm'"
            if ([string]::IsNullOrWhiteSpace($nm)) { $pickerState.StatusLbl.Text = "Name required."; return }
            $rV=0;$gV=0;$bV=0
            [int]::TryParse($pickerState.RBox.Text,[ref]$rV)|Out-Null
            [int]::TryParse($pickerState.GBox.Text,[ref]$gV)|Out-Null
            [int]::TryParse($pickerState.BBox.Text,[ref]$bV)|Out-Null
            $rV=[Math]::Max(0,[Math]::Min(255,$rV));$gV=[Math]::Max(0,[Math]::Min(255,$gV));$bV=[Math]::Max(0,[Math]::Min(255,$bV))
            $hex9 = "#{0:X2}{1:X2}{2:X2}FF" -f $rV,$gV,$bV
            Write-Log "btnSaveFilament: hex=$hex9, editingName='$($pickerState.EditingName)'"
            if ($pickerState.EditingName -and $pickerState.EditingName -ne $nm -and $capturedLibColors.Contains($pickerState.EditingName)) {
                # Capture old hex BEFORE removing so we can clean up HexToName correctly
                $oldHex = $capturedLibColors[$pickerState.EditingName]
                $capturedLibColors.Remove($pickerState.EditingName) | Out-Null
                if ($oldHex) { $capturedHexToName.Remove($oldHex) | Out-Null; $capturedHexToName.Remove($oldHex.Substring(0,7)) | Out-Null }
                if ($capturedLibGradients.Contains($pickerState.EditingName)) {
                    $capturedLibGradients[$nm] = $capturedLibGradients[$pickerState.EditingName]
                    $capturedLibGradients.Remove($pickerState.EditingName) | Out-Null
                }
            }
            $capturedLibColors[$nm] = $hex9; $capturedHexToName[$hex9] = $nm; $capturedHexToName[$hex9.Substring(0,7)] = $nm
            $pickerState.EditingName = $nm
            Write-Log "btnSaveFilament: calling Save-FilamentLibrary"
            if (Save-FilamentLibrary) {
                $pickerState.StatusLbl.Text = "Saved."; $pickerState.StatusLbl.Foreground = Get-WpfColor "#4CAF72"
            } else {
                $pickerState.StatusLbl.Text = "Save failed!"; $pickerState.StatusLbl.Foreground = Get-WpfColor "#D95F5F"
            }
            Write-Log "btnSaveFilament: calling Rebuild-FilamentList"
            Rebuild-FilamentList $pickerState
            Write-Log "btnSaveFilament: done"
        } catch {
            Write-Log "btnSaveFilament: EXCEPTION $($_.Exception.GetType().FullName): $($_.Exception.Message)" "ERROR"
            Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"
            throw
        }
    }.GetNewClosure())

    # â”€â”€ Delete â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $btnDelFilament.Add_Click({
        $nm = $pickerState.EditingName
        if ([string]::IsNullOrWhiteSpace($nm) -or -not $capturedLibColors.Contains($nm)) {
            $pickerState.StatusLbl.Text = "Nothing selected to delete."; return
        }
        $capturedLibColors.Remove($nm) | Out-Null
        $capturedLibGradients.Remove($nm) | Out-Null
        if (Save-FilamentLibrary) {
            $pickerState.EditingName = ""; $pickerState.NameBox.Text = ""
            $pickerState.RBox.Text="0"; $pickerState.GBox.Text="0"; $pickerState.BBox.Text="0"
            $pickerState.StatusLbl.Text = "Deleted: $nm"; $pickerState.StatusLbl.Foreground = Get-WpfColor "#E8A135"
        } else {
            $pickerState.StatusLbl.Text = "Delete (save) failed!"; $pickerState.StatusLbl.Foreground = Get-WpfColor "#D95F5F"
        }
        Rebuild-FilamentList $pickerState
    }.GetNewClosure())

    # â”€â”€ Add New button â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $btnAddNew.Add_Click({
        $pickerState.EditingName = ""
        $pickerState.NameBox.Text = "New Filament"
        $pickerState.RBox.Text="128"; $pickerState.GBox.Text="128"; $pickerState.BBox.Text="128"
        Update-PickerFromRgb $pickerState
        $pickerState.StatusLbl.Text = "New entry - fill in name and color then Save."
        $pickerState.StatusLbl.Foreground = Get-WpfColor "#888A9A"
    }.GetNewClosure())

    # Initial list build
    Rebuild-FilamentList $pickerState
    if ($script:LibraryColors.Count -gt 0) {
        $firstName = @($script:LibraryColors.Keys)[0]
        Load-FilamentEntry $firstName $pickerState
    }

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  NAMING CONVENTIONS SECTION â€” Themes | Printer Prefixes | Tags & Labels
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    foreach ($w in @(240, 16, 240, 16, 1)) {
        $cd = New-Object System.Windows.Controls.ColumnDefinition
        $cd.Width = if ($w -eq 1) { [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star) } `
                    else          { [System.Windows.GridLength]::new($w) }
        $secNaming.ColumnDefinitions.Add($cd) | Out-Null
    }

    function Build-NameListSection($parentGrid, [string]$title, [scriptblock]$getItems, [scriptblock]$setItems, [scriptblock]$onSave, [int]$col = 0) {
        $g = New-Object System.Windows.Controls.Grid; $g.Margin = New-Object System.Windows.Thickness(20,16,20,16)
        $g.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
        $g.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) | Out-Null
        $g.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
        $g.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
        [System.Windows.Controls.Grid]::SetColumn($g, $col); $parentGrid.Children.Add($g) | Out-Null

        $hdr = New-Object System.Windows.Controls.TextBlock; $hdr.Text = $title
        $hdr.FontSize = 14; $hdr.FontWeight = [System.Windows.FontWeights]::Bold
        $hdr.Foreground = Get-WpfColor "#C8CFDD"; $hdr.Margin = New-Object System.Windows.Thickness(0,0,0,10)
        [System.Windows.Controls.Grid]::SetRow($hdr, 0); $g.Children.Add($hdr) | Out-Null

        $sv = New-Object System.Windows.Controls.ScrollViewer; $sv.VerticalScrollBarVisibility = "Auto"
        $sv.HorizontalScrollBarVisibility = "Disabled"
        $sv.Background = Get-WpfColor "#1C1D23"
        $sv.BorderBrush = Get-WpfColor "#2A2C38"; $sv.BorderThickness = New-Object System.Windows.Thickness(1)
        [System.Windows.Controls.Grid]::SetRow($sv, 1); $g.Children.Add($sv) | Out-Null

        $listBox = New-Object System.Windows.Controls.ListBox
        $listBox.Background = Get-WpfColor "#1C1D23"; $listBox.Foreground = Get-WpfColor "#C0C4D0"
        $listBox.BorderThickness = 0; $listBox.FontSize = 12; $listBox.Padding = New-Object System.Windows.Thickness(4)
        foreach ($item in (& $getItems)) {
            $lbi = New-Object System.Windows.Controls.ListBoxItem
            $lbi.Content = $item; $lbi.Foreground = Get-WpfColor "#C0C4D0"; $lbi.Tag = "clean"
            $listBox.Items.Add($lbi) | Out-Null
        }
        $sv.Content = $listBox

        # Add row
        $addGrid = New-Object System.Windows.Controls.Grid; $addGrid.Margin = New-Object System.Windows.Thickness(0,8,0,6)
        $addGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
        $addGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(8) })) | Out-Null
        $addGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(70) })) | Out-Null
        [System.Windows.Controls.Grid]::SetRow($addGrid, 2); $g.Children.Add($addGrid) | Out-Null

        $tbNew = New-Object System.Windows.Controls.TextBox; $tbNew.Height = 28; $tbNew.FontSize = 12; $tbNew.Padding = New-Object System.Windows.Thickness(6,0,6,0)
        $tbNew.Background = Get-WpfColor "#1C1D23"; $tbNew.Foreground = Get-WpfColor "#E0E3EC"
        $tbNew.BorderBrush = Get-WpfColor "#3A3C4A"; $tbNew.BorderThickness = New-Object System.Windows.Thickness(1)
        [System.Windows.Controls.Grid]::SetColumn($tbNew, 0); $addGrid.Children.Add($tbNew) | Out-Null

        $btnAdd = New-Object System.Windows.Controls.Button; $btnAdd.Content = "Add"; $btnAdd.Height = 28; $btnAdd.FontSize = 11
        $btnAdd.Background = Get-WpfColor "#2E5A42"; $btnAdd.Foreground = Get-WpfColor "#FFFFFF"
        $btnAdd.BorderThickness = 0; $btnAdd.Cursor = [System.Windows.Input.Cursors]::Hand
        [System.Windows.Controls.Grid]::SetColumn($btnAdd, 2); $addGrid.Children.Add($btnAdd) | Out-Null

        # Delete + Save row
        $actionGrid = New-Object System.Windows.Controls.Grid
        $actionGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
        $actionGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(8) })) | Out-Null
        $actionGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
        [System.Windows.Controls.Grid]::SetRow($actionGrid, 3); $g.Children.Add($actionGrid) | Out-Null

        $btnDel = New-Object System.Windows.Controls.Button; $btnDel.Content = "Delete Selected"; $btnDel.Height = 28; $btnDel.FontSize = 11
        $btnDel.Background = Get-WpfColor "#6B2828"; $btnDel.Foreground = Get-WpfColor "#FFFFFF"
        $btnDel.BorderThickness = 0; $btnDel.Cursor = [System.Windows.Input.Cursors]::Hand
        [System.Windows.Controls.Grid]::SetColumn($btnDel, 0); $actionGrid.Children.Add($btnDel) | Out-Null

        $btnSave = New-Object System.Windows.Controls.Button; $btnSave.Content = "Save Changes"; $btnSave.Height = 28; $btnSave.FontSize = 11; $btnSave.FontWeight = [System.Windows.FontWeights]::Bold
        $btnSave.Background = Get-WpfColor "#3A5080"; $btnSave.Foreground = Get-WpfColor "#FFFFFF"
        $btnSave.BorderThickness = 0; $btnSave.Cursor = [System.Windows.Input.Cursors]::Hand
        [System.Windows.Controls.Grid]::SetColumn($btnSave, 2); $actionGrid.Children.Add($btnSave) | Out-Null

        # Wire up events
        # Click list item â†’ populate text box + switch button to "Edit"
        $listBox.Tag = @{ TextBox=$tbNew; AddBtn=$btnAdd }
        $listBox.Add_SelectionChanged({
            $t = $this.Tag
            if ($null -ne $this.SelectedItem) {
                $t.TextBox.Text = "$($this.SelectedItem.Content)"
                $t.AddBtn.Content = "Edit"
            } else {
                $t.TextBox.Text = ""
                $t.AddBtn.Content = "Add"
            }
        }.GetNewClosure())

        # Toggle-deselect: clicking an already-selected row deselects it
        $capturedListBox = $listBox
        $capturedTbNew   = $tbNew
        $capturedBtnAdd  = $btnAdd
        $capturedListBox.Add_PreviewMouseLeftButtonDown({
            $hit = $args[1].OriginalSource
            # Walk up the visual tree to find the ListBoxItem that was clicked
            $lbi = $hit
            while ($null -ne $lbi -and -not ($lbi -is [System.Windows.Controls.ListBoxItem])) {
                $lbi = [System.Windows.Media.VisualTreeHelper]::GetParent($lbi)
            }
            if ($null -ne $lbi -and ($lbi -is [System.Windows.Controls.ListBoxItem]) -and $lbi.IsSelected) {
                # Deselect and clear the text box; mark event handled so WPF
                # doesn't immediately re-select the item on mouse-up
                $capturedListBox.SelectedItem = $null
                $capturedTbNew.Text = ""
                $capturedBtnAdd.Content = "Add"
                $args[1].Handled = $true
            }
        }.GetNewClosure())

        $btnAdd.Tag = @{ Box=$tbNew; List=$listBox; Btn=$btnAdd }
        $btnAdd.Add_Click({
            $t = $this.Tag; $v = $t.Box.Text.Trim(); if ([string]::IsNullOrWhiteSpace($v)) { return }
            $sel = $t.List.SelectedItem   # a ListBoxItem or $null
            if ($null -ne $sel -and "$($sel.Content)" -ne $v) {
                # Rename in place â€” mark dirty
                $sel.Content = $v
                $sel.Foreground = Get-WpfColor "#E8903A"
                $sel.Tag = "dirty"
            } elseif ($null -eq $sel -or "$($sel.Content)" -ne $v) {
                # Check for duplicate content
                $dup = $false
                foreach ($lbi in $t.List.Items) { if ("$($lbi.Content)" -eq $v) { $dup = $true; break } }
                if (-not $dup) {
                    $lbi = New-Object System.Windows.Controls.ListBoxItem
                    $lbi.Content = $v; $lbi.Foreground = Get-WpfColor "#E8903A"; $lbi.Tag = "dirty"
                    $t.List.Items.Add($lbi) | Out-Null
                    $t.List.SelectedItem = $lbi
                }
            }
            $t.Box.Text = ""
            $t.Btn.Content = "Add"
            $t.List.SelectedItem = $null
        }.GetNewClosure())

        $btnDel.Tag = $listBox
        $btnDel.Add_Click({
            $lb = $this.Tag
            if ($null -ne $lb.SelectedItem) { $lb.Items.Remove($lb.SelectedItem) }
        }.GetNewClosure())

        $btnSave.Tag = @{ List=$listBox; SetItems=$setItems; OnSave=$onSave; Sv=$sv }
        $btnSave.Add_Click({
            $t = $this.Tag
            $items = @($t.List.Items | ForEach-Object { "$($_.Content)" })
            & $t.SetItems $items
            $ok = & $t.OnSave
            if ($ok) {
                # Clear dirty highlights
                foreach ($lbi in $t.List.Items) { $lbi.Foreground = Get-WpfColor "#C0C4D0"; $lbi.Tag = "clean" }
                $t.Sv.BorderBrush = Get-WpfColor "#4CAF72"
            } else {
                $t.Sv.BorderBrush = Get-WpfColor "#D95F5F"
            }
        }.GetNewClosure())

        return $listBox
    }

    # Themes â€” column 0
    Build-NameListSection $secNaming "Themes" `
        { $script:GpThemes } `
        { param($v) Set-GpThemes $v } `
        { Write-Log "Save-NamesLibrary (Themes): called"; $ok = Save-NamesLibrary; Write-Log "Save-NamesLibrary (Themes): $(if($ok){'success'}else{'FAILED'})"; $ok } 0 | Out-Null

    # Printer Prefixes â€” column 2
    Build-NameListSection $secNaming "Printer Prefixes" `
        { $script:PrinterPrefixes } `
        { param($v) Set-PrinterPrefixes $v } `
        { Write-Log "Save-NamesLibrary (Prefixes): called"; $ok = Save-NamesLibrary; Write-Log "Save-NamesLibrary (Prefixes): $(if($ok){'success'}else{'FAILED'})"; $ok } 2 | Out-Null

    # Tags & Labels â€” column 4
    $tagsContainer = New-Object System.Windows.Controls.Grid; $tagsContainer.Margin = New-Object System.Windows.Thickness(0,16,20,16)
    $tagsContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
    $tagsContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) | Out-Null
    $tagsContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
    $tagsContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($tagsContainer, 4); $secNaming.Children.Add($tagsContainer) | Out-Null

    $tagsHdr = New-Object System.Windows.Controls.TextBlock; $tagsHdr.Text = "Tags & Labels"
    $tagsHdr.FontSize = 14; $tagsHdr.FontWeight = [System.Windows.FontWeights]::Bold
    $tagsHdr.Foreground = Get-WpfColor "#C8CFDD"; $tagsHdr.Margin = New-Object System.Windows.Thickness(0,0,0,10)
    [System.Windows.Controls.Grid]::SetRow($tagsHdr, 0); $tagsContainer.Children.Add($tagsHdr) | Out-Null

    $tagsSv = New-Object System.Windows.Controls.ScrollViewer; $tagsSv.VerticalScrollBarVisibility = "Auto"
    $tagsSv.Background = Get-WpfColor "#1C1D23"; $tagsSv.BorderBrush = Get-WpfColor "#2A2C38"; $tagsSv.BorderThickness = New-Object System.Windows.Thickness(1)
    [System.Windows.Controls.Grid]::SetRow($tagsSv, 1); $tagsContainer.Children.Add($tagsSv) | Out-Null

    $tagsListStack = New-Object System.Windows.Controls.StackPanel; $tagsListStack.Margin = New-Object System.Windows.Thickness(4)
    $tagsSv.Content = $tagsListStack

    # Header row
    $tagsColHdr = New-Object System.Windows.Controls.Grid; $tagsColHdr.Margin = New-Object System.Windows.Thickness(0,0,0,2)
    $tagsColHdr.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(80) })) | Out-Null
    $tagsColHdr.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $th1 = New-Object System.Windows.Controls.TextBlock; $th1.Text="Tag";   $th1.Foreground=Get-WpfColor "#555868"; $th1.FontSize=11; [System.Windows.Controls.Grid]::SetColumn($th1,0)
    $th2 = New-Object System.Windows.Controls.TextBlock; $th2.Text="Label"; $th2.Foreground=Get-WpfColor "#555868"; $th2.FontSize=11; [System.Windows.Controls.Grid]::SetColumn($th2,1)
    $tagsColHdr.Children.Add($th1)|Out-Null; $tagsColHdr.Children.Add($th2)|Out-Null
    $tagsListStack.Children.Add($tagsColHdr) | Out-Null

    Rebuild-TagsList $tagsListStack


    # Add row for tags
    $tagsAddGrid = New-Object System.Windows.Controls.Grid; $tagsAddGrid.Margin = New-Object System.Windows.Thickness(0,8,0,6)
    $tagsAddGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(80) })) | Out-Null
    $tagsAddGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(8) })) | Out-Null
    $tagsAddGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $tagsAddGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(8) })) | Out-Null
    $tagsAddGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(60) })) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($tagsAddGrid,2); $tagsContainer.Children.Add($tagsAddGrid) | Out-Null

    function New-TagInput {
        $tb = New-Object System.Windows.Controls.TextBox; $tb.Height=28; $tb.FontSize=12; $tb.Padding=New-Object System.Windows.Thickness(6,0,6,0)
        $tb.Background=Get-WpfColor "#1C1D23"; $tb.Foreground=Get-WpfColor "#E0E3EC"
        $tb.BorderBrush=Get-WpfColor "#3A3C4A"; $tb.BorderThickness=New-Object System.Windows.Thickness(1)
        return $tb
    }
    $tbTagName = New-TagInput; [System.Windows.Controls.Grid]::SetColumn($tbTagName,0); $tagsAddGrid.Children.Add($tbTagName)|Out-Null
    $tbTagLbl  = New-TagInput; [System.Windows.Controls.Grid]::SetColumn($tbTagLbl, 2); $tagsAddGrid.Children.Add($tbTagLbl)|Out-Null
    $btnAddTag = New-Object System.Windows.Controls.Button; $btnAddTag.Content="Add"; $btnAddTag.Height=28; $btnAddTag.FontSize=11
    $btnAddTag.Background=Get-WpfColor "#2E5A42"; $btnAddTag.Foreground=Get-WpfColor "#FFFFFF"; $btnAddTag.BorderThickness=0; $btnAddTag.Cursor=[System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnAddTag,4); $tagsAddGrid.Children.Add($btnAddTag)|Out-Null
    # Register inputs so Rebuild-TagsList click handlers can populate them
    $script:TagEditState.TagBox = $tbTagName
    $script:TagEditState.LblBox = $tbTagLbl
    $script:TagEditState.AddBtn = $btnAddTag

    $tagsActGrid = New-Object System.Windows.Controls.Grid
    $tagsActGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $tagsActGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(8) })) | Out-Null
    $tagsActGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($tagsActGrid,3); $tagsContainer.Children.Add($tagsActGrid) | Out-Null

    $btnDelTag  = New-Object System.Windows.Controls.Button; $btnDelTag.Content="Delete Selected"; $btnDelTag.Height=28; $btnDelTag.FontSize=11
    $btnDelTag.Background=Get-WpfColor "#6B2828"; $btnDelTag.Foreground=Get-WpfColor "#FFFFFF"; $btnDelTag.BorderThickness=0; $btnDelTag.Cursor=[System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnDelTag,0); $tagsActGrid.Children.Add($btnDelTag)|Out-Null

    $btnSaveTags = New-Object System.Windows.Controls.Button; $btnSaveTags.Content="Save Changes"; $btnSaveTags.Height=28; $btnSaveTags.FontSize=11; $btnSaveTags.FontWeight=[System.Windows.FontWeights]::Bold
    $btnSaveTags.Background=Get-WpfColor "#3A5080"; $btnSaveTags.Foreground=Get-WpfColor "#FFFFFF"; $btnSaveTags.BorderThickness=0; $btnSaveTags.Cursor=[System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Grid]::SetColumn($btnSaveTags,2); $tagsActGrid.Children.Add($btnSaveTags)|Out-Null

    # Capture $script: variables as locals â€” $script: scope is NOT reliably accessible
    # inside .GetNewClosure() closures running on the WPF dispatcher thread.
    $capturedTagEditState = $script:TagEditState
    $capturedTagsStack    = $tagsListStack

    $btnAddTag.Add_Click({
        Write-Log "btnAddTag: clicked"
        try {
            $tn = $capturedTagEditState.TagBox.Text.Trim()
            $tl = $capturedTagEditState.LblBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($tn)) { Write-Log "btnAddTag: tag name empty, skipping"; return }
            Invoke-AddTagEntry $tn $tl $capturedTagsStack
        } catch { Write-Log "btnAddTag EXCEPTION: $($_.Exception.Message)" "ERROR"; Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"; throw }
    }.GetNewClosure())

    $btnDelTag.Add_Click({
        Write-Log "btnDelTag: clicked"
        try {
            $tn = $capturedTagEditState.SelectedTag
            if ([string]::IsNullOrWhiteSpace($tn)) { Write-Log "btnDelTag: nothing selected, skipping"; return }
            Invoke-RemoveTagEntry $tn $capturedTagsStack
        } catch { Write-Log "btnDelTag EXCEPTION: $($_.Exception.Message)" "ERROR"; Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"; throw }
    }.GetNewClosure())

    $btnSaveTags.Add_Click({
        Write-Log "btnSaveTags: clicked"
        try {
            if (Invoke-SaveTagsSection $capturedTagsStack) {
                $this.Background = Get-WpfColor "#4CAF72"
            } else {
                $this.Background = Get-WpfColor "#D95F5F"
            }
        } catch { Write-Log "btnSaveTags EXCEPTION: $($_.Exception.Message)" "ERROR"; Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"; throw }
    }.GetNewClosure())

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  PURGE DICTIONARY SECTION
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    $purgeContainer = New-Object System.Windows.Controls.Grid
    $purgeContainer.Margin = New-Object System.Windows.Thickness(20,16,20,20)
    $purgeContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
    $purgeContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
    $purgeContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) | Out-Null
    $purgeContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) | Out-Null
    $secPurge.Children.Add($purgeContainer) | Out-Null

    $purgeHdrBar = New-Object System.Windows.Controls.DockPanel
    $purgeHdrBar.LastChildFill = $false; $purgeHdrBar.Margin = New-Object System.Windows.Thickness(0,0,0,10)
    [System.Windows.Controls.Grid]::SetRow($purgeHdrBar, 0); $purgeContainer.Children.Add($purgeHdrBar) | Out-Null

    $purgeHdr = New-Object System.Windows.Controls.TextBlock; $purgeHdr.Text = "Purge Dictionary"
    $purgeHdr.FontSize = 14; $purgeHdr.FontWeight = [System.Windows.FontWeights]::Bold
    $purgeHdr.Foreground = Get-WpfColor "#C8CFDD"
    [System.Windows.Controls.DockPanel]::SetDock($purgeHdr, [System.Windows.Controls.Dock]::Left)
    $purgeHdrBar.Children.Add($purgeHdr) | Out-Null

    $purgeAvgSavings = New-Object System.Windows.Controls.TextBlock; $purgeAvgSavings.Text = "Avg Savings (tuned): -"
    $purgeAvgSavings.FontSize = 13; $purgeAvgSavings.FontWeight = [System.Windows.FontWeights]::SemiBold
    $purgeAvgSavings.Foreground = Get-WpfColor "#7FD8A0"; $purgeAvgSavings.VerticalAlignment = "Center"
    [System.Windows.Controls.DockPanel]::SetDock($purgeAvgSavings, [System.Windows.Controls.Dock]::Right)
    $purgeHdrBar.Children.Add($purgeAvgSavings) | Out-Null

    # â”€â”€ Filter bar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $purgeFilterPanel = New-Object System.Windows.Controls.StackPanel
    $purgeFilterPanel.Orientation = "Horizontal"; $purgeFilterPanel.Margin = New-Object System.Windows.Thickness(0,0,0,12)
    [System.Windows.Controls.Grid]::SetRow($purgeFilterPanel, 1); $purgeContainer.Children.Add($purgeFilterPanel) | Out-Null

    function New-PurgeFilterCombo([string]$labelText, [double]$width) {
        $grp = New-Object System.Windows.Controls.StackPanel
        $grp.Orientation = "Vertical"; $grp.Margin = New-Object System.Windows.Thickness(0,0,16,0)
        $purgeFilterPanel.Children.Add($grp) | Out-Null

        $lbl = New-Object System.Windows.Controls.TextBlock; $lbl.Text = $labelText; $lbl.FontSize = 11
        $lbl.Foreground = Get-WpfColor "#888A9A"; $lbl.Margin = New-Object System.Windows.Thickness(0,0,0,4)
        $grp.Children.Add($lbl) | Out-Null

        $cmb = New-Object System.Windows.Controls.ComboBox
        $cmb.Width = $width; $cmb.Height = 28
        $cmb.Background = Get-WpfColor "#1C1D23"; $cmb.Foreground = Get-WpfColor "#E0E3EC"
        $cmb.BorderBrush = Get-WpfColor "#3A3C4A"; $cmb.BorderThickness = New-Object System.Windows.Thickness(1)
        $cmb.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, [System.Windows.Controls.ToolBar]::ComboBoxStyleKey)
        $cmb.Resources[[System.Windows.SystemColors]::WindowBrushKey]        = Get-WpfColor "#1C1D23"
        $cmb.Resources[[System.Windows.SystemColors]::WindowTextBrushKey]    = Get-WpfColor "#E0E3EC"
        $cmb.Resources[[System.Windows.SystemColors]::HighlightBrushKey]     = Get-WpfColor "#3A5080"
        $cmb.Resources[[System.Windows.SystemColors]::HighlightTextBrushKey] = Get-WpfColor "#FFFFFF"
        $cmbItemStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
        $cmbItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, (Get-WpfColor "#1C1D23"))))
        $cmbItemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, (Get-WpfColor "#E0E3EC"))))
        $cmb.ItemContainerStyle = $cmbItemStyle
        $grp.Children.Add($cmb) | Out-Null
        return $cmb
    }

    # Make a filter combo typable/searchable - mirrors the filament color-picker combo pattern:
    # IsEditable + manual CollectionView filtering as the user types, with Enter/Tab/Escape handling.
    function Enable-PurgeComboTypeAhead([System.Windows.Controls.ComboBox]$combo, [string[]]$options) {
        $items = New-Object System.Collections.ObjectModel.ObservableCollection[string]
        foreach ($opt in $options) { [void]$items.Add($opt) }
        $combo.ItemsSource = $items
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($items)

        $combo.IsEditable = $true
        $combo.IsTextSearchEnabled = $false
        $combo.SelectedIndex = 0

        $state = @{ Filtering = $false; NeedsFilter = $false; View = $view }
        $combo.Tag = $state

        $combo.Add_PreviewTextInput({
            param($s, $e)
            $s.Tag.NeedsFilter = $true
        }) | Out-Null

        $combo.Add_PreviewKeyDown({
            param($s, $e)
            $st = $s.Tag
            switch ($e.Key) {
                { $_ -eq [System.Windows.Input.Key]::Back -or $_ -eq [System.Windows.Input.Key]::Delete } { $st.NeedsFilter = $true }
                { $_ -eq [System.Windows.Input.Key]::Down -or $_ -eq [System.Windows.Input.Key]::Up } {
                    if (-not $s.IsDropDownOpen -and $s.Items.Count -gt 0) { $s.IsDropDownOpen = $true; $e.Handled = $true }
                }
                { $_ -eq [System.Windows.Input.Key]::Enter -or $_ -eq [System.Windows.Input.Key]::Tab } {
                    if ($s.IsDropDownOpen -and $s.Items.Count -gt 0) {
                        $accepted = if ($s.SelectedItem) { $s.SelectedItem.ToString() } else { $s.Items[0].ToString() }
                        $st.Filtering = $true
                        $s.IsDropDownOpen = $false
                        $st.View.Filter = $null
                        $st.Filtering = $false
                        $s.SelectedItem = $accepted
                        $tb = $s.Template.FindName("PART_EditableTextBox", $s)
                        if ($tb) { $tb.SelectAll() }
                        $e.Handled = $true
                    }
                }
                { $_ -eq [System.Windows.Input.Key]::Escape } {
                    if ($s.IsDropDownOpen) {
                        $st.Filtering = $true
                        $s.IsDropDownOpen = $false
                        $st.View.Filter = $null
                        $st.Filtering = $false
                        if ($s.SelectedItem) { $s.Text = $s.SelectedItem.ToString() }
                        $e.Handled = $true
                    }
                }
            }
        }) | Out-Null

        $textChangedHandler = [System.Windows.Controls.TextChangedEventHandler]{
            param($s, $e)
            $st = $s.Tag
            if ($st.Filtering -or -not $st.NeedsFilter) { return }
            $st.NeedsFilter = $false
            $st.Filtering = $true
            $typed = $s.Text
            if ([string]::IsNullOrWhiteSpace($typed)) {
                $st.View.Filter = $null
                $s.IsDropDownOpen = $false
            } else {
                $tl = $typed.ToLower()
                $st.View.Filter = [Predicate[object]]{ param($item) $item.ToString().ToLower().Contains($tl) }.GetNewClosure()
                $s.IsDropDownOpen = (-not $st.View.IsEmpty)
            }
            $st.Filtering = $false
        }
        $combo.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, $textChangedHandler)

        $combo.Add_DropDownClosed({
            param($s, $e)
            $st = $s.Tag
            if ($st.Filtering) { return }
            $st.Filtering = $true
            $st.View.Filter = $null
            $st.Filtering = $false
        }) | Out-Null
    }

    $cmbPurgeFrom = New-PurgeFilterCombo "From" 220
    $purgeFromOptions = @("(All)") + @($script:PurgeDict | Select-Object -ExpandProperty Source_Filament -Unique | Sort-Object)
    Enable-PurgeComboTypeAhead $cmbPurgeFrom $purgeFromOptions

    $cmbPurgeTuned = New-PurgeFilterCombo "Tuned" 130
    Enable-PurgeComboTypeAhead $cmbPurgeTuned @("(All)","Tuned only","Not tuned")

    $cmbPurgeVol = New-PurgeFilterCombo "Tuned Volume" 150
    Enable-PurgeComboTypeAhead $cmbPurgeVol @("(All)","Has value","Empty")

    $purgeGrid = New-Object System.Windows.Controls.DataGrid
    $purgeGrid.AutoGenerateColumns = $false
    $purgeGrid.CanUserAddRows = $false
    $purgeGrid.CanUserSortColumns = $true
    $purgeGrid.HeadersVisibility = "Column"
    $purgeGrid.GridLinesVisibility = "Horizontal"
    $purgeGrid.Background = Get-WpfColor "#1C1D23"
    $purgeGrid.Foreground = Get-WpfColor "#E0E3EC"
    $purgeGrid.RowBackground = Get-WpfColor "#1C1D23"
    $purgeGrid.AlternatingRowBackground = Get-WpfColor "#21222A"
    $purgeGrid.BorderBrush = Get-WpfColor "#2A2C38"; $purgeGrid.BorderThickness = New-Object System.Windows.Thickness(1)
    $purgeGrid.RowHeaderWidth = 0
    [System.Windows.Controls.Grid]::SetRow($purgeGrid, 2); $purgeContainer.Children.Add($purgeGrid) | Out-Null

    # Column headers default to a washed-out gray on this dark theme â€” force a legible style
    $purgeHdrStyle = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
    $purgeHdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BackgroundProperty, (Get-WpfColor "#252630"))))
    $purgeHdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::ForegroundProperty, (Get-WpfColor "#E8EBF2"))))
    $purgeHdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::FontWeightProperty, [System.Windows.FontWeights]::SemiBold)))
    $purgeHdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::PaddingProperty, (New-Object System.Windows.Thickness(8,6,8,6)))))
    $purgeHdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BorderBrushProperty, (Get-WpfColor "#2A2C38"))))
    $purgeHdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BorderThicknessProperty, (New-Object System.Windows.Thickness(0,0,1,1)))))
    $purgeGrid.ColumnHeaderStyle = $purgeHdrStyle

    # Cell highlight: green while saved-but-unedited, orange while dirty (dirty takes precedence).
    # Built from a XAML string so WPF's own TypeConverter handles the bool->Value comparison --
    # code-behind DataTrigger Value=$true can silently fail in PowerShell's WPF interop.
    function New-PurgeCellStyle([string]$dirtyProp, [string]$savedProp) {
        $xaml = @”
<Style TargetType='DataGridCell'
       xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
       xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
  <Style.Triggers>
    <DataTrigger Binding='{Binding $savedProp}' Value='True'>
      <Setter Property='Background' Value='#1E472E'/>
      <Setter Property='Foreground' Value='#88DDAA'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding $dirtyProp}' Value='True'>
      <Setter Property='Background' Value='#B9711F'/>
      <Setter Property='Foreground' Value='#1A1A1A'/>
    </DataTrigger>
  </Style.Triggers>
</Style>
“@
        $xr = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xaml)))
        return [System.Windows.Markup.XamlReader]::Load($xr)
    }

    # Swatch + name template for the From/To columns
    function New-PurgeSwatchTemplate([string]$hexProp, [string]$textProp) {
        $xaml = @"
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
              xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <StackPanel Orientation="Horizontal" Margin="6,2,6,2">
    <Border Width="14" Height="14" CornerRadius="2" Margin="0,0,8,0" VerticalAlignment="Center"
            BorderBrush="#3A3C4A" BorderThickness="1">
      <Border.Background>
        <SolidColorBrush Color="{Binding $hexProp}"/>
      </Border.Background>
    </Border>
    <TextBlock Text="{Binding $textProp}" Foreground="#E0E3EC" VerticalAlignment="Center"/>
  </StackPanel>
</DataTemplate>
"@
        $xr = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xaml)))
        return [System.Windows.Markup.XamlReader]::Load($xr)
    }

    # Centered text for numeric-style columns
    $purgeCenterStyle = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    $purgeCenterStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::TextAlignmentProperty, [System.Windows.TextAlignment]::Center))) | Out-Null
    $purgeCenterEditStyle = New-Object System.Windows.Style([System.Windows.Controls.TextBox])
    $purgeCenterEditStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBox]::TextAlignmentProperty, [System.Windows.TextAlignment]::Center))) | Out-Null

    $colFrom = New-Object System.Windows.Controls.DataGridTemplateColumn
    $colFrom.Header = "From"; $colFrom.CellTemplate = New-PurgeSwatchTemplate "Source_Hex" "Source_Filament"
    $colFrom.SortMemberPath = "Source_Filament"
    $colFrom.Width = New-Object System.Windows.Controls.DataGridLength(220)
    $colFrom.CellStyle = New-PurgeCellStyle "FromDirty" "FromSaved"
    $purgeGrid.Columns.Add($colFrom) | Out-Null

    $colTo = New-Object System.Windows.Controls.DataGridTemplateColumn
    $colTo.Header = "To"; $colTo.CellTemplate = New-PurgeSwatchTemplate "Target_Hex" "Target_Filament"
    $colTo.SortMemberPath = "Target_Filament"
    $colTo.Width = New-Object System.Windows.Controls.DataGridLength(220)
    $colTo.CellStyle = New-PurgeCellStyle "ToDirty" "ToSaved"
    $purgeGrid.Columns.Add($colTo) | Out-Null

    $colTuned = New-Object System.Windows.Controls.DataGridCheckBoxColumn
    $colTuned.Header = "Tuned"; $colTuned.Binding = New-Object System.Windows.Data.Binding("Tuned")
    $colTuned.Width = 60
    $colTuned.CellStyle = New-PurgeCellStyle "TunedDirty" "TunedSaved"
    $purgeGrid.Columns.Add($colTuned) | Out-Null

    $colTunedVol = New-Object System.Windows.Controls.DataGridTextColumn
    $colTunedVol.Header = "Tuned Volume"; $colTunedVol.Binding = New-Object System.Windows.Data.Binding("Tuned_Volume")
    $colTunedVol.Width = 110
    $colTunedVol.ElementStyle = $purgeCenterStyle; $colTunedVol.EditingElementStyle = $purgeCenterEditStyle
    $colTunedVol.CellStyle = New-PurgeCellStyle "TunedVolumeDirty" "TunedVolumeSaved"
    $purgeGrid.Columns.Add($colTunedVol) | Out-Null

    $colBaseVol = New-Object System.Windows.Controls.DataGridTextColumn
    $colBaseVol.Header = "Base Volume"; $colBaseVol.Binding = New-Object System.Windows.Data.Binding("Base_Volume")
    $colBaseVol.Width = 110
    $colBaseVol.ElementStyle = $purgeCenterStyle; $colBaseVol.EditingElementStyle = $purgeCenterEditStyle
    $colBaseVol.CellStyle = New-PurgeCellStyle "BaseVolumeDirty" "BaseVolumeSaved"
    $purgeGrid.Columns.Add($colBaseVol) | Out-Null

    $colSavings = New-Object System.Windows.Controls.DataGridTextColumn
    $colSavings.Header = "% Savings"; $colSavings.Binding = New-Object System.Windows.Data.Binding("Savings_Pct")
    $colSavings.Width = 100; $colSavings.IsReadOnly = $true
    $colSavings.ElementStyle = $purgeCenterStyle
    $purgeGrid.Columns.Add($colSavings) | Out-Null

    $purgeGrid.ItemsSource = $script:PurgeDict
    $purgeView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:PurgeDict)
    $purgeView.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("Source_Filament", [System.ComponentModel.ListSortDirection]::Ascending))) | Out-Null

    # Rolling average % savings across tuned combos (untuned/unparseable rows are excluded)
    $capturedPurgeAvgSavings = $purgeAvgSavings
    $capturedPurgeDictForAvg = $script:PurgeDict
    $updatePurgeAvgSavings = {
        $vals = @()
        foreach ($r in $capturedPurgeDictForAvg) {
            if ($r.SavingsValue -ne $null) { $vals += [double]$r.SavingsValue }
        }
        if ($vals.Count -gt 0) {
            $avg = ($vals | Measure-Object -Average).Average
            $capturedPurgeAvgSavings.Text = "Avg Savings (tuned): {0:N1}% across {1} combo{2}" -f $avg, $vals.Count, $(if ($vals.Count -eq 1) { "" } else { "s" })
        } else {
            $capturedPurgeAvgSavings.Text = "Avg Savings (tuned): -"
        }
    }.GetNewClosure()
    & $updatePurgeAvgSavings

    # â”€â”€ Filter predicate â€” re-applied whenever a filter combo changes â”€â”€â”€â”€â”€â”€â”€â”€
    $capturedPurgeView      = $purgeView
    $capturedCmbPurgeFrom   = $cmbPurgeFrom
    $capturedCmbPurgeTuned  = $cmbPurgeTuned
    $capturedCmbPurgeVol    = $cmbPurgeVol

    $applyPurgeFilter = {
        $fromSel  = $capturedCmbPurgeFrom.SelectedItem
        $tunedSel = $capturedCmbPurgeTuned.SelectedItem
        $volSel   = $capturedCmbPurgeVol.SelectedItem
        $capturedPurgeView.Filter = [Predicate[Object]]{
            param($item)
            if ($fromSel -and $fromSel -ne "(All)" -and $item.Source_Filament -ne $fromSel) { return $false }
            if ($tunedSel -eq "Tuned only" -and -not $item.Tuned) { return $false }
            if ($tunedSel -eq "Not tuned" -and $item.Tuned) { return $false }
            $hasVol = -not [string]::IsNullOrWhiteSpace($item.Tuned_Volume)
            if ($volSel -eq "Has value" -and -not $hasVol) { return $false }
            if ($volSel -eq "Empty" -and $hasVol) { return $false }
            return $true
        }
        $capturedPurgeView.Refresh()
    }.GetNewClosure()

    $cmbPurgeFrom.Add_SelectionChanged({ Write-Log "cmbPurgeFrom: changed"; & $applyPurgeFilter }.GetNewClosure())
    $cmbPurgeTuned.Add_SelectionChanged({ Write-Log "cmbPurgeTuned: changed"; & $applyPurgeFilter }.GetNewClosure())
    $cmbPurgeVol.Add_SelectionChanged({ Write-Log "cmbPurgeVol: changed"; & $applyPurgeFilter }.GetNewClosure())

    $purgeBtnGrid = New-Object System.Windows.Controls.Grid; $purgeBtnGrid.Margin = New-Object System.Windows.Thickness(0,10,0,0)
    $purgeBtnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(160) })) | Out-Null
    $purgeBtnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(8) })) | Out-Null
    $purgeBtnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(140) })) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($purgeBtnGrid, 3); $purgeContainer.Children.Add($purgeBtnGrid) | Out-Null

    $btnSavePurge = New-Object System.Windows.Controls.Button; $btnSavePurge.Content = "Save Changes"
    $btnSavePurge.Height = 32; $btnSavePurge.FontSize = 12; $btnSavePurge.FontWeight = [System.Windows.FontWeights]::Bold
    $btnSavePurge.Background = Get-WpfColor "#3A5080"; $btnSavePurge.Foreground = Get-WpfColor "#FFFFFF"
    $btnSavePurge.BorderThickness = 0; $btnSavePurge.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnSavePurge.IsEnabled = $false
    [System.Windows.Controls.Grid]::SetColumn($btnSavePurge, 0); $purgeBtnGrid.Children.Add($btnSavePurge) | Out-Null

    $btnDiscardPurge = New-Object System.Windows.Controls.Button; $btnDiscardPurge.Content = "Discard Changes"
    $btnDiscardPurge.Height = 32; $btnDiscardPurge.FontSize = 12; $btnDiscardPurge.FontWeight = [System.Windows.FontWeights]::Bold
    $btnDiscardPurge.Background = Get-WpfColor "#3A3A3A"; $btnDiscardPurge.Foreground = Get-WpfColor "#666666"
    $btnDiscardPurge.BorderThickness = 0; $btnDiscardPurge.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnDiscardPurge.IsEnabled = $false
    [System.Windows.Controls.Grid]::SetColumn($btnDiscardPurge, 2); $purgeBtnGrid.Children.Add($btnDiscardPurge) | Out-Null

    $capturedPurgeGrid     = $purgeGrid
    $capturedBtnSavePurge  = $btnSavePurge

    # IsEnabled is driven entirely by PropertyChanged on each PurgeDictRow (below).
    # [action]{} blocks inside BeginInvoke lose the $script: scope when run from a
    # GetNewClosure() handler, so relying on $script:PurgeDict there was unreliable.
    # CellEditEnding is kept only to refresh the avg-savings label.
    $capturedUpdateAvgSavings2 = $updatePurgeAvgSavings
    $purgeGrid.Add_CellEditEnding({
        $upd = $capturedUpdateAvgSavings2
        $capturedPurgeGrid.Dispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::Background, [action]{ & $upd }) | Out-Null
    }.GetNewClosure())

    $capturedBtnSavePurge3    = $btnSavePurge
    $capturedBtnDiscardPurge3 = $btnDiscardPurge
    $capturedPurgeDict3       = $script:PurgeDict
    $capturedUpdateAvgSavings3 = $updatePurgeAvgSavings
    $purgeRowChangedHandler = [System.ComponentModel.PropertyChangedEventHandler]({
        param($sender, $e)
        if ($e.PropertyName -eq "IsDirty") {
            $anyDirty = $false
            foreach ($r in $capturedPurgeDict3) { if ($r.IsDirty) { $anyDirty = $true; break } }
            if ($anyDirty) {
                $capturedBtnSavePurge3.Background    = Get-WpfColor "#3A5080"
                $capturedBtnDiscardPurge3.Background = Get-WpfColor "#6B3030"
                $capturedBtnDiscardPurge3.Foreground = Get-WpfColor "#FFFFFF"
            } else {
                $capturedBtnDiscardPurge3.Background = Get-WpfColor "#3A3A3A"
                $capturedBtnDiscardPurge3.Foreground = Get-WpfColor "#666666"
            }
            $capturedBtnSavePurge3.IsEnabled    = $anyDirty
            $capturedBtnDiscardPurge3.IsEnabled = $anyDirty
            & $capturedUpdateAvgSavings3
        }
    }.GetNewClosure())
    foreach ($row in $script:PurgeDict) {
        $row.add_PropertyChanged($purgeRowChangedHandler)
    }

    # Captured for use inside save/discard clicks (handler must be detached before bulk ops)
    $capturedRowHandler       = $purgeRowChangedHandler
    $capturedBtnDiscardPurge4 = $btnDiscardPurge
    $capturedBtnSavePurge4    = $btnSavePurge
    $capturedPurgeDict4       = $script:PurgeDict
    $capturedPurgeGrid4       = $purgeGrid

    $btnDiscardPurge.Add_Click({
        Write-Log "btnDiscardPurge: clicked"
        $capturedPurgeGrid4.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null
        # Detach handler so per-row RevertToBaseline doesn't trigger N^2 dirty scans
        foreach ($r in $capturedPurgeDict4) { $r.remove_PropertyChanged($capturedRowHandler) }
        foreach ($r in $capturedPurgeDict4) { if ($r.IsDirty) { $r.RevertToBaseline() } }
        foreach ($r in $capturedPurgeDict4) { $r.add_PropertyChanged($capturedRowHandler) }
        $capturedBtnSavePurge4.IsEnabled    = $false
        $capturedBtnDiscardPurge4.IsEnabled = $false
        $capturedBtnDiscardPurge4.Background = Get-WpfColor "#3A3A3A"
        $capturedBtnDiscardPurge4.Foreground = Get-WpfColor "#666666"
    }.GetNewClosure())

    $capturedBtnDiscardForSave = $btnDiscardPurge
    $capturedPurgeDictForSave  = $script:PurgeDict
    $capturedRowHandlerForSave = $purgeRowChangedHandler
    $btnSavePurge.Add_Click({
        Write-Log "btnSavePurge: clicked"
        try {
            $capturedPurgeGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) | Out-Null
            if (Save-PurgeDictionary) {
                # Detach handler so per-row CommitBaseline doesn't trigger N^2 dirty scans
                foreach ($r in $capturedPurgeDictForSave) { $r.remove_PropertyChanged($capturedRowHandlerForSave) }
                foreach ($r in $capturedPurgeDictForSave) { $r.CommitBaseline() }
                foreach ($r in $capturedPurgeDictForSave) { $r.add_PropertyChanged($capturedRowHandlerForSave) }
                $this.Background = Get-WpfColor "#4CAF72"
                $this.IsEnabled  = $false
                $capturedBtnDiscardForSave.IsEnabled = $false
                $capturedBtnDiscardForSave.Background = Get-WpfColor "#3A3A3A"
                $capturedBtnDiscardForSave.Foreground = Get-WpfColor "#666666"
            } else {
                $this.Background = Get-WpfColor "#D95F5F"
            }
        } catch { Write-Log "btnSavePurge EXCEPTION: $($_.Exception.Message)" "ERROR"; Write-Log "  STACK: $($_.ScriptStackTrace)" "ERROR"; throw }
    }.GetNewClosure())

    $script:LibrariesPanel = $root
    return $root
}

function New-ModeButton([string]$label, [int]$width, [string]$bg, [string]$fg, [bool]$active) {
    $b = New-Object System.Windows.Controls.Button
    $b.Content = $label; $b.Width = $width; $b.Height = 34
    $b.FontWeight = [System.Windows.FontWeights]::Bold; $b.FontSize = 12
    $b.Background = Get-WpfColor $bg; $b.Foreground = Get-WpfColor $fg
    $b.BorderThickness = New-Object System.Windows.Thickness(1)
    $borderColor = if ($active) { "#5A78C4" } else { "#3A3D50" }
    $b.BorderBrush = Get-WpfColor $borderColor
    $b.Cursor = [System.Windows.Input.Cursors]::Hand
    $b.Margin = New-Object System.Windows.Thickness(4,0,4,0)
    $b.IsEnabled = -not $active
    return $b
}

$script:BtnModeFilePr   = New-ModeButton "File Prep" 85  "#3A5080" "#FFFFFF" $true
$script:BtnModeEditing  = New-ModeButton "Editing"   75  "#252630" "#7A7D90" $false
$script:BtnModeReview   = New-ModeButton "Review"    70  "#252630" "#7A7D90" $false
$script:BtnModeLibraries = New-ModeButton "Libraries" 80  "#252630" "#7A7D90" $false

$script:BtnModeFilePr.Add_Click({   Set-GlobalMode "FilePr"   })
$script:BtnModeEditing.Add_Click({  Set-GlobalMode "Editing"  })
$script:BtnModeReview.Add_Click({   Set-GlobalMode "Review"   })
$script:BtnModeLibraries.Add_Click({ Write-Log "BtnModeLibraries: clicked"; Set-GlobalMode "Libraries" })

$topModeBar.Children.Add($script:BtnModeFilePr)    | Out-Null
$topModeBar.Children.Add($script:BtnModeEditing)   | Out-Null
$topModeBar.Children.Add($script:BtnModeReview)    | Out-Null

# â”€â”€ Libraries button sits LEFT of the Browse button in the center column â”€â”€
# Navigate up from BtnBrowse: StackPanel â†’ header Grid (col 1)
$browseStack  = $btnBrowse.Parent            # vertical StackPanel (Browse + hint text)
$headerGrid   = $browseStack.Parent          # the Grid inside the header Border

# Pull the browse stack out, wrap it with the Libraries button in a horizontal SP
$headerGrid.Children.Remove($browseStack) | Out-Null

$centerWrap = New-Object System.Windows.Controls.StackPanel
$centerWrap.Orientation = "Horizontal"
$centerWrap.HorizontalAlignment = "Center"
$centerWrap.VerticalAlignment   = "Center"
[System.Windows.Controls.Grid]::SetColumn($centerWrap, 1)

$script:BtnModeLibraries.Margin = New-Object System.Windows.Thickness(0,0,12,0)
$script:BtnModeLibraries.VerticalAlignment = "Center"
$centerWrap.Children.Add($script:BtnModeLibraries) | Out-Null

$browseStack.VerticalAlignment = "Center"
$centerWrap.Children.Add($browseStack) | Out-Null

$headerGrid.Children.Add($centerWrap) | Out-Null

# Build the Libraries panel (attaches itself to the window's root Grid)
Build-LibrariesPanel | Out-Null

Write-Log "Entering window.ShowDialog()"
try {
    $window.ShowDialog() | Out-Null
    Write-Log "window.ShowDialog() returned normally"
} catch {
    Write-Log "window.ShowDialog() THREW: $($_.Exception.GetType().FullName): $($_.Exception.Message)" "FATAL"
    Write-Log "  STACK: $($_.ScriptStackTrace)" "FATAL"
}
Write-Log "Script exiting"
try { Stop-Transcript } catch {}
