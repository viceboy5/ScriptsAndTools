#Requires -Version 3.0
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Hide console window -----------------------------------------------------
try {
    Add-Type -Name ConsoleHider -Namespace Win32 -MemberDefinition @'
        [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]   public static extern bool   ShowWindow(IntPtr hWnd, int nCmdShow);
'@
    [Win32.ConsoleHider]::ShowWindow([Win32.ConsoleHider]::GetConsoleWindow(), 0) | Out-Null
} catch { }

# --- IFileOpenDialog COM -----------------------------------------------------
Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class ModernFolderPicker {

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("D57C7288-D4AD-4768-BE02-9D969532D960"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
        [PreserveSig] int Show(IntPtr hwndOwner);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IntPtr psi);
        void SetFolder(IntPtr psi);
        void GetFolder(out IntPtr ppsi);
        void GetCurrentSelection(out IntPtr ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IntPtr psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IntPtr ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    static readonly Guid CLSID_FileOpenDialog = new Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7");
    const uint FOS_PICKFOLDERS     = 0x00000020;
    const uint FOS_FORCEFILESYSTEM = 0x00000040;
    const uint SIGDN_FILESYSPATH   = 0x80058000;

    public static string Pick(IntPtr owner, string title) {
        try {
            Type t        = Type.GetTypeFromCLSID(CLSID_FileOpenDialog);
            object inst   = Activator.CreateInstance(t);
            var    dialog = (IFileOpenDialog)inst;
            try {
                uint opts;
                dialog.GetOptions(out opts);
                dialog.SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                if (!string.IsNullOrEmpty(title)) dialog.SetTitle(title);
                int hr = dialog.Show(owner);
                if (hr != 0) return null;
                IShellItem item;
                dialog.GetResult(out item);
                string path;
                item.GetDisplayName(SIGDN_FILESYSPATH, out path);
                Marshal.ReleaseComObject(item);
                return path;
            } finally {
                Marshal.ReleaseComObject(dialog);
            }
        } catch {
            return null;
        }
    }
}
'@

# --- Colors ------------------------------------------------------------------
function clr($hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
$cBG      = clr "#111214"
$cBG2     = clr "#16171B"
$cBG3     = clr "#1C1D23"
$cBG4     = clr "#0D0E10"
$cBorder  = clr "#2A2C35"
$cText    = clr "#DDE0E8"
$cMuted   = clr "#5A5D6B"
$cAmber   = clr "#E8A135"
$cBlue    = clr "#6DAEE0"
$cRed     = clr "#D95F5F"
$cGreen   = clr "#4CAF72"
$cDropHL  = clr "#1E2010"
$cInput   = clr "#1E2028"
$cLabel   = clr "#44475A"
$cThemeBG = clr "#181920"
$cThemeFG = clr "#4A4D5C"
$cOldTag  = clr "#3A3D4A"
$cOldText = clr "#6B6E7A"
$cNewTag  = clr "#1E3328"
$cNewText = clr "#4CAF72"
$cBtnOff  = clr "#1E2028"
$cBtnOffT = clr "#3A3D4A"
$cBtnOn   = clr "#2D5A3D"
$cBtnOnT  = clr "#6FD98F"
$cBtnOnBd = clr "#4CAF72"
$cFileSep = clr "#1A1C22"

# --- Data stores -------------------------------------------------------------
# db[gpKey]       = { Name; Path; Parents: OrderedDict }
# db[gpKey].Parents[ppKey] = { Name; Path; Files: List<{Path;Name;Suffix;Extension;Stem;Parts}> }
# gpRows[gpKey]   = { NewBox; ThemeLabels: List; ParentKeys: List; OldName }
# parentRows[ppKey] = { Char; Adj; ThemeLabel; NewLabel; GpKey; OldName; FilePreviews: List }
$script:db         = [System.Collections.Specialized.OrderedDictionary]::new()
$script:gpRows     = @{}
$script:parentRows = @{}
$script:resizeList = [System.Collections.Generic.List[hashtable]]::new()

# --- ParseFile: detect suffix and extension ----------------------------------
# Compound extensions are checked first (longest match wins).
# Suffix = last token when stem is split on delimiter runs ( space . _ - ).
function script:ParseFile([string]$filename) {
    $compoundExts = @('.gcode.3mf', '.gcode.stl', '.gcode.step', '.f3d.3mf')
    $ext = $null
    foreach ($ce in $compoundExts) {
        if ($filename.ToLower().EndsWith($ce.ToLower())) {
            $ext = $filename.Substring($filename.Length - $ce.Length)
            break
        }
    }
    if (-not $ext) { $ext = [System.IO.Path]::GetExtension($filename) }
    if (-not $ext) { $ext = '' }

    $stem  = $filename.Substring(0, $filename.Length - $ext.Length)
    $parts = [string[]]( ($stem -split '[\s._-]+') | Where-Object { $_ -ne '' } )

    # .png files carry no suffix token — the whole stem is just the name
    if ($ext -ieq '.png') {
        return @{ Suffix = ''; Extension = $ext; Stem = $stem; Parts = $parts }
    }

    $suffix = if ($parts.Count -gt 0) { $parts[-1] } else { $stem }

    return @{ Suffix = $suffix; Extension = $ext; Stem = $stem; Parts = $parts }
}

# --- SmartFill: detect Character and Adj from the Full.3mf anchor in a parent
# After rename:  CuteDuck[_Rare]_Farm_Full.3mf  -> Char=CuteDuck, Adj=Rare
# Before rename: Cute Duck_Full.3mf             -> Char=CuteDuck, Adj=
function script:SmartFill([string]$ppKey, [string]$gpName) {
    $pd = $null
    foreach ($gk in $script:db.Keys) {
        if ($script:db[$gk].Parents.Contains($ppKey)) { $pd = $script:db[$gk].Parents[$ppKey]; break }
    }
    if (-not $pd) { return @{ Char = ''; Adj = '' } }

    # Find the Full.3mf anchor in this parent
    $anchor = $null
    foreach ($fi in $pd.Files) {
        if ($fi.Suffix -ieq 'Full' -and $fi.Extension -ieq '.3mf') { $anchor = $fi; break }
    }
    if (-not $anchor) { return @{ Char = ''; Adj = '' } }

    $stem    = $anchor.Stem
    $escaped = [regex]::Escape($gpName)

    # Already renamed: ends with _Theme_Full
    if ($gpName -ne '' -and $stem -imatch "^(.+)_${escaped}_Full$") {
        $prefix   = $Matches[1]
        $segments = $prefix -split '_'
        if ($segments.Count -ge 2) {
            # First segment = Character, rest = Adj (joined, stripping separators)
            $adj = ($segments[1..($segments.Count-1)]) -join ''
            return @{ Char = $segments[0]; Adj = $adj }
        }
        return @{ Char = $prefix; Adj = '' }
    }

    # Original file: all parts before the last one (Full), joined without separator
    $parts = $anchor.Parts
    if ($parts.Count -ge 2) {
        $charRaw = $parts[0..($parts.Count - 2)] -join ''
        return @{ Char = $charRaw; Adj = '' }
    }
    return @{ Char = ($parts -join ''); Adj = '' }
}

# --- Helpers -----------------------------------------------------------------
function script:NoSpaces([string]$s) { $s -replace '[^a-zA-Z0-9]', '' }

function script:InputLayout([int]$panelW) {
    $lm = 12; $rm = 12; $sepW = 16; $gap = 4
    $avail  = $panelW - $lm - $rm - ($sepW * 2) - ($gap * 4)
    $wChar  = [int]($avail * 0.36)
    $wAdj   = [int]($avail * 0.28)
    $wTheme = $avail - $wChar - $wAdj
    $xChar  = $lm
    $xSep1  = $xChar + $wChar + $gap
    $xAdj   = $xSep1 + $sepW + $gap
    $xSep2  = $xAdj + $wAdj + $gap
    $xTheme = $xSep2 + $sepW + $gap
    return @{ xChar=$xChar; wChar=$wChar; xSep1=$xSep1; xAdj=$xAdj; wAdj=$wAdj; xSep2=$xSep2; xTheme=$xTheme; wTheme=$wTheme }
}

# Factory functions -- each call creates a truly independent scope so the
# captured control references don't get stomped by subsequent loop iterations.
function script:ApplyResize {
    $w = [Math]::Max(620, $scroll.ClientSize.Width - 8)
    $lay = script:InputLayout $w
    $halfW = [int](($w - 8 - 60) / 2)

    foreach ($entry in $script:resizeList) {
        if ($entry.Type -eq 'gp') {
            $entry.Panel.Width  = $w
            $entry.NewBox.Width = $w - 290
        } elseif ($entry.Type -eq 'parent') {
            $entry.Panel.Width    = $w
            $entry.Divider.Width  = $w - 16
            $entry.Char.Location  = New-Object System.Drawing.Point($lay.xChar,  56)
            $entry.Char.Width     = $lay.wChar
            $entry.Sep1.Location  = New-Object System.Drawing.Point($lay.xSep1, 57)
            $entry.Adj.Location   = New-Object System.Drawing.Point($lay.xAdj,   56)
            $entry.Adj.Width      = $lay.wAdj
            $entry.Sep2.Location  = New-Object System.Drawing.Point($lay.xSep2, 57)
            $entry.Theme.Location = New-Object System.Drawing.Point($lay.xTheme, 56)
            $entry.Theme.Width    = $lay.wTheme
            $entry.LChar.Location = New-Object System.Drawing.Point($lay.xChar,  42)
            $entry.LAdj.Location  = New-Object System.Drawing.Point($lay.xAdj,   42)
            $entry.LTheme.Location= New-Object System.Drawing.Point($lay.xTheme, 42)
            foreach ($fi in $entry.Files) {
                $fi.RowPanel.Width    = $w - 8
                $fi.OldLabel.Width    = $halfW
                $fi.Arrow.Location    = New-Object System.Drawing.Point(($halfW + 62), $fi.Arrow.Location.Y)
                $fi.NewLabel.Width    = $halfW
                $fi.NewLabel.Location = New-Object System.Drawing.Point(($halfW + 82), 0)
            }
        }
    }
}

# Anchors are still Full.3mf -- the presence of one qualifies the parent folder
function script:IsMatch([string]$name) { $name -imatch 'full\.3mf' }

# --- RegisterParent: add a parent folder and ALL its files to db -------------
function script:RegisterParent([string]$parentPath, $gpDir) {
    $gpKey  = if ($gpDir) { $gpDir.FullName } else { '__ROOT__' }
    $gpName = if ($gpDir) { $gpDir.Name     } else { '(root)' }

    # Already registered
    if ($script:db.Contains($gpKey) -and $script:db[$gpKey].Parents.Contains($parentPath)) { return }

    if (-not $script:db.Contains($gpKey)) {
        $script:db[$gpKey] = [PSCustomObject]@{
            Name    = $gpName
            Path    = $gpKey
            Parents = [System.Collections.Specialized.OrderedDictionary]::new()
        }
    }
    $gd = $script:db[$gpKey]

    if (-not $gd.Parents.Contains($parentPath)) {
        $parentDir = [System.IO.DirectoryInfo]::new($parentPath)
        $gd.Parents[$parentPath] = [PSCustomObject]@{
            Name  = $parentDir.Name
            Path  = $parentPath
            Files = [System.Collections.Generic.List[object]]::new()
        }
        # Collect ALL files in this folder (non-recursive), sorted by name
        try {
            Get-ChildItem -LiteralPath $parentPath -File -ErrorAction SilentlyContinue |
                Sort-Object Name |
                ForEach-Object {
                    $parsed = script:ParseFile $_.Name
                    $gd.Parents[$parentPath].Files.Add([PSCustomObject]@{
                        Path      = $_.FullName
                        Name      = $_.Name
                        Suffix    = $parsed.Suffix
                        Extension = $parsed.Extension
                        Stem      = $parsed.Stem
                        Parts     = $parsed.Parts
                    })
                }
        } catch { }
    }
}

# --- ScanPath ----------------------------------------------------------------
function script:ScanPath([string]$rawPath) {
    $path = $rawPath.Trim('"').Trim()
    if ([System.IO.Directory]::Exists($path)) {
        try {
            Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { script:IsMatch $_.Name } |
                ForEach-Object { script:RegisterParent $_.Directory.FullName $_.Directory.Parent }
        } catch { }
    } elseif ([System.IO.File]::Exists($path)) {
        if (script:IsMatch ([System.IO.Path]::GetFileName($path))) {
            $fi = [System.IO.FileInfo]::new($path)
            script:RegisterParent $fi.Directory.FullName $fi.Directory.Parent
        }
    }
}

# --- Name computation --------------------------------------------------------
# File:   Character[_Adj]_Theme_Suffix + Extension
# Parent: Theme_Character[_Adj]

function script:ComputeFileNew([string]$ppKey, [string]$suffix, [string]$ext) {
    $pRow  = $script:parentRows[$ppKey]; if (-not $pRow) { return '' }
    $gpRow = $script:gpRows[$pRow.GpKey]
    $theme = if ($gpRow) { script:NoSpaces $gpRow.NewBox.Text.Trim() } else { '' }
    $char  = script:NoSpaces $pRow.Char.Text.Trim()
    $adj   = script:NoSpaces $pRow.Adj.Text.Trim()
    if ($char -eq '') { return '' }
    $parts = @($char)
    if ($adj    -ne '') { $parts += $adj    }
    if ($theme  -ne '') { $parts += $theme  }
    if ($suffix -ne '') { $parts += $suffix }
    return ($parts -join '_') + $ext
}

function script:ComputeParentNew([string]$ppKey) {
    $pRow  = $script:parentRows[$ppKey]; if (-not $pRow) { return '' }
    $gpRow = $script:gpRows[$pRow.GpKey]
    $theme = if ($gpRow) { script:NoSpaces $gpRow.NewBox.Text.Trim() } else { '' }
    $char  = script:NoSpaces $pRow.Char.Text.Trim()
    $adj   = script:NoSpaces $pRow.Adj.Text.Trim()
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($theme -ne '') { $parts.Add($theme) }
    if ($char  -ne '') { $parts.Add($char)  }
    if ($adj   -ne '') { $parts.Add($adj)   }
    if ($parts.Count -eq 0) { return '' }
    return $parts -join '_'
}

# --- Live update chain -------------------------------------------------------
function script:UpdateFromGp([string]$gpKey) {
    $gpRow = $script:gpRows[$gpKey]; if (-not $gpRow) { return }
    $newTheme = script:NoSpaces $gpRow.NewBox.Text.Trim()
    foreach ($lbl in $gpRow.ThemeLabels) { $lbl.Text = "  $newTheme" }
    foreach ($ppKey in $gpRow.ParentKeys) {
        $pRow = $script:parentRows[$ppKey]
        if ($pRow) {
            $v = script:ComputeParentNew $ppKey
            $pRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { 'NEW:  (enter fields to preview)' }
            foreach ($fp in $pRow.FilePreviews) {
                $fv = script:ComputeFileNew $ppKey $fp.Suffix $fp.Extension
                $fp.NewLabel.Text = if ($fv) { $fv } else { '(enter Character)' }
            }
        }
    }
    script:UpdateRenameButton
}

function script:UpdateFromParent([string]$ppKey) {
    $pRow = $script:parentRows[$ppKey]; if (-not $pRow) { return }
    $v = script:ComputeParentNew $ppKey
    $pRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { 'NEW:  (enter fields to preview)' }
    foreach ($fp in $pRow.FilePreviews) {
        $fv = script:ComputeFileNew $ppKey $fp.Suffix $fp.Extension
        $fp.NewLabel.Text = if ($fv) { $fv } else { '(enter Character)' }
    }
    script:UpdateRenameButton
}

# --- RENAME button state -----------------------------------------------------
function script:StyleRenameBtn([bool]$on) {
    if ($on) {
        $btnRename.BackColor                         = $cBtnOn
        $btnRename.ForeColor                         = $cBtnOnT
        $btnRename.FlatAppearance.BorderSize         = 1
        $btnRename.FlatAppearance.BorderColor        = $cBtnOnBd
        $btnRename.FlatAppearance.MouseOverBackColor = clr '#3A7050'
        $btnRename.Cursor                            = [System.Windows.Forms.Cursors]::Hand
    } else {
        $btnRename.BackColor                         = $cBtnOff
        $btnRename.ForeColor                         = $cBtnOffT
        $btnRename.FlatAppearance.BorderSize         = 1
        $btnRename.FlatAppearance.BorderColor        = $cBorder
        $btnRename.FlatAppearance.MouseOverBackColor = $cBtnOff
        $btnRename.Cursor                            = [System.Windows.Forms.Cursors]::Default
    }
}

function script:UpdateRenameButton {
    if ($script:db.Count -eq 0) { $btnRename.Enabled = $false; script:StyleRenameBtn $false; return }
    $ok = $true
    foreach ($gk in $script:gpRows.Keys) {
        if ((script:NoSpaces $script:gpRows[$gk].NewBox.Text.Trim()) -eq '') { $ok = $false; break }
    }
    if ($ok) {
        foreach ($ppKey in $script:parentRows.Keys) {
            if ((script:NoSpaces $script:parentRows[$ppKey].Char.Text.Trim()) -eq '') { $ok = $false; break }
        }
    }
    $btnRename.Enabled = $ok
    script:StyleRenameBtn $ok
}

# --- Conflict resolution dialog ----------------------------------------------
# Shows two conflicting files side by side with dates/sizes, returns path to delete or $null to abort.
function script:ShowConflictDialog([string]$newName, [string]$pathA, [string]$pathB) {
    $fiA = [System.IO.FileInfo]::new($pathA)
    $fiB = [System.IO.FileInfo]::new($pathB)

    function FmtSize($bytes) {
        if ($bytes -ge 1MB) { return "{0:N1} MB" -f ($bytes / 1MB) }
        return "{0:N0} KB" -f ($bytes / 1KB)
    }

    $dlg                  = New-Object System.Windows.Forms.Form
    $dlg.Text             = 'Name Conflict'
    $dlg.ClientSize       = New-Object System.Drawing.Size(680, 310)
    $dlg.FormBorderStyle  = 'FixedDialog'
    $dlg.StartPosition    = 'CenterParent'
    $dlg.BackColor        = clr '#111214'
    $dlg.ForeColor        = clr '#DDE0E8'
    $dlg.MaximizeBox      = $false
    $dlg.MinimizeBox      = $false

    $title                = New-Object System.Windows.Forms.Label
    $title.Text           = 'Rename Conflict'
    $title.Font           = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $title.ForeColor      = clr '#D95F5F'
    $title.AutoSize       = $true
    $title.Location       = New-Object System.Drawing.Point(20, 16)
    $dlg.Controls.Add($title)

    $sub                  = New-Object System.Windows.Forms.Label
    $sub.Text             = "Both files below would be renamed to:  $newName"
    $sub.Font             = New-Object System.Drawing.Font('Segoe UI', 9)
    $sub.ForeColor        = clr '#6B6E7A'
    $sub.AutoSize         = $true
    $sub.Location         = New-Object System.Drawing.Point(20, 46)
    $dlg.Controls.Add($sub)

    $newLbl               = New-Object System.Windows.Forms.Label
    $newLbl.Text          = $newName
    $newLbl.Font          = New-Object System.Drawing.Font('Consolas', 10, [System.Drawing.FontStyle]::Bold)
    $newLbl.ForeColor     = clr '#E8A135'
    $newLbl.AutoSize      = $true
    $newLbl.Location      = New-Object System.Drawing.Point(20, 64)
    $dlg.Controls.Add($newLbl)

    # Helper to build a file card
    function MakeCard($fi, $x, $label) {
        $card                  = New-Object System.Windows.Forms.Panel
        $card.Size             = New-Object System.Drawing.Size(300, 140)
        $card.Location         = New-Object System.Drawing.Point($x, 100)
        $card.BackColor        = clr '#16171B'

        $cardBorder = New-Object System.Windows.Forms.Panel
        $cardBorder.Size     = New-Object System.Drawing.Size(300, 140)
        $cardBorder.Location = New-Object System.Drawing.Point(0, 0)
        $cardBorder.BackColor = clr '#2A2C35'
        $cardBorder.Add_Paint({
            param($s,$e)
            $e.Graphics.Clear($s.BackColor)
        })

        $hdr                   = New-Object System.Windows.Forms.Label
        $hdr.Text              = $label
        $hdr.Font              = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $hdr.ForeColor         = clr '#44475A'
        $hdr.AutoSize          = $true
        $hdr.Location          = New-Object System.Drawing.Point(12, 10)
        $card.Controls.Add($hdr)

        $fname                 = New-Object System.Windows.Forms.Label
        $fname.Text            = $fi.Name
        $fname.Font            = New-Object System.Drawing.Font('Consolas', 8)
        $fname.ForeColor       = clr '#DDE0E8'
        $fname.AutoSize        = $false
        $fname.Size            = New-Object System.Drawing.Size(276, 32)
        $fname.Location        = New-Object System.Drawing.Point(12, 26)
        $fname.TextAlign       = 'MiddleLeft'
        $card.Controls.Add($fname)

        $mod                   = New-Object System.Windows.Forms.Label
        $mod.Text              = 'Modified:  ' + $fi.LastWriteTime.ToString('yyyy-MM-dd  h:mm tt')
        $mod.Font              = New-Object System.Drawing.Font('Segoe UI', 8)
        $mod.ForeColor         = clr '#6DAEE0'
        $mod.AutoSize          = $true
        $mod.Location          = New-Object System.Drawing.Point(12, 62)
        $card.Controls.Add($mod)

        $sz                    = New-Object System.Windows.Forms.Label
        $sz.Text               = 'Size:  ' + (FmtSize $fi.Length)
        $sz.Font               = New-Object System.Drawing.Font('Segoe UI', 8)
        $sz.ForeColor          = clr '#6B6E7A'
        $sz.AutoSize           = $true
        $sz.Location           = New-Object System.Drawing.Point(12, 80)
        $card.Controls.Add($sz)

        $created               = New-Object System.Windows.Forms.Label
        $created.Text          = 'Created:  ' + $fi.CreationTime.ToString('yyyy-MM-dd  h:mm tt')
        $created.Font          = New-Object System.Drawing.Font('Segoe UI', 8)
        $created.ForeColor     = clr '#6B6E7A'
        $created.AutoSize      = $true
        $created.Location      = New-Object System.Drawing.Point(12, 97)
        $card.Controls.Add($created)

        return $card
    }

    $cardA = MakeCard $fiA 20  'FILE A  (will be renamed ->'
    $cardB = MakeCard $fiB 350 'FILE B  (will be renamed ->'

    # Newer badge
    if ($fiA.LastWriteTime -gt $fiB.LastWriteTime) {
        $nb = New-Object System.Windows.Forms.Label; $nb.Text = 'NEWER'
        $nb.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $nb.ForeColor = clr '#4CAF72'; $nb.BackColor = clr '#1E3328'
        $nb.AutoSize = $true; $nb.Location = New-Object System.Drawing.Point(12, 118); $cardA.Controls.Add($nb)
    } elseif ($fiB.LastWriteTime -gt $fiA.LastWriteTime) {
        $nb = New-Object System.Windows.Forms.Label; $nb.Text = 'NEWER'
        $nb.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $nb.ForeColor = clr '#4CAF72'; $nb.BackColor = clr '#1E3328'
        $nb.AutoSize = $true; $nb.Location = New-Object System.Drawing.Point(12, 118); $cardB.Controls.Add($nb)
    }

    $dlg.Controls.Add($cardA)
    $dlg.Controls.Add($cardB)

    # Instruction
    $inst              = New-Object System.Windows.Forms.Label
    $inst.Text         = 'Choose which file to DELETE before renaming, or skip this conflict:'
    $inst.Font         = New-Object System.Drawing.Font('Segoe UI', 8)
    $inst.ForeColor    = clr '#5A5D6B'
    $inst.AutoSize     = $true
    $inst.Location     = New-Object System.Drawing.Point(20, 250)
    $dlg.Controls.Add($inst)

    $result = $null

    function MakeBtn($text, $bg, $fg, $x, $w) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text; $b.Size = New-Object System.Drawing.Size($w, 28)
        $b.Location = New-Object System.Drawing.Point($x, 272)
        $b.FlatStyle = 'Flat'; $b.BackColor = $bg; $b.ForeColor = $fg
        $b.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 8)
        $b.FlatAppearance.BorderSize = 0
        $b.Cursor = [System.Windows.Forms.Cursors]::Hand
        $dlg.Controls.Add($b)
        return $b
    }

    $btnDelA   = MakeBtn "Delete File A" (clr '#3A1A1A') (clr '#D95F5F') 20  130
    $btnDelB   = MakeBtn "Delete File B" (clr '#3A1A1A') (clr '#D95F5F') 162 130
    $btnSkip   = MakeBtn "Skip This Conflict" (clr '#1E2028') (clr '#5A5D6B') 304 160
    $btnAbort  = MakeBtn "Abort Rename" (clr '#1E2028') (clr '#5A5D6B') 474 160

    $btnDelA.Add_Click({  $script:conflictResult = 'A'; $dlg.Close() })
    $btnDelB.Add_Click({  $script:conflictResult = 'B'; $dlg.Close() })
    $btnSkip.Add_Click({  $script:conflictResult = 'Skip'; $dlg.Close() })
    $btnAbort.Add_Click({ $script:conflictResult = 'Abort'; $dlg.Close() })

    $script:conflictResult = 'Abort'
    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
    return $script:conflictResult
}

# --- Perform renames ---------------------------------------------------------
function script:DoRename {
    $errors  = [System.Collections.Generic.List[string]]::new()
    $ops     = [System.Collections.Generic.List[object]]::new()
    $renamed = 0

    # Build file ops
    foreach ($ppKey in $script:parentRows.Keys) {
        $pRow = $script:parentRows[$ppKey]
        foreach ($fp in $pRow.FilePreviews) {
            $newName = script:ComputeFileNew $ppKey $fp.Suffix $fp.Extension
            if ($newName -eq '' -or -not [System.IO.File]::Exists($fp.Path)) { continue }
            $dir     = [System.IO.Path]::GetDirectoryName($fp.Path)
            $newPath = [System.IO.Path]::Combine($dir, $newName)
            if ($fp.Path -ne $newPath) {
                $ops.Add([PSCustomObject]@{ Type = 'file'; OldPath = $fp.Path; NewPath = $newPath })
            }
        }
    }

    # --- Conflict detection: two source files -> same target path ------------
    # Group file ops by NewPath and find duplicates
    $byTarget = @{}
    foreach ($op in $ops) {
        if ($op.Type -ne 'file') { continue }
        $key = $op.NewPath.ToLower()
        if (-not $byTarget.ContainsKey($key)) { $byTarget[$key] = [System.Collections.Generic.List[object]]::new() }
        $byTarget[$key].Add($op)
    }

    $toRemoveFromOps = [System.Collections.Generic.List[object]]::new()

    foreach ($key in $byTarget.Keys) {
        $group = $byTarget[$key]
        if ($group.Count -lt 2) { continue }

        # Show conflict dialog for the first two in this group
        $opA = $group[0]; $opB = $group[1]
        $choice = script:ShowConflictDialog ([System.IO.Path]::GetFileName($opA.NewPath)) $opA.OldPath $opB.OldPath

        if ($choice -eq 'Abort') { return }

        if ($choice -eq 'A') {
            # Delete file A, remove its op so only B gets renamed
            try { [System.IO.File]::Delete($opA.OldPath) } catch { $errors.Add("Could not delete $([System.IO.Path]::GetFileName($opA.OldPath)): $($_.Exception.Message)") }
            $toRemoveFromOps.Add($opA)
        } elseif ($choice -eq 'B') {
            # Delete file B, remove its op
            try { [System.IO.File]::Delete($opB.OldPath) } catch { $errors.Add("Could not delete $([System.IO.Path]::GetFileName($opB.OldPath)): $($_.Exception.Message)") }
            $toRemoveFromOps.Add($opB)
        } else {
            # Skip -- remove both from ops so neither gets renamed
            $toRemoveFromOps.Add($opA)
            $toRemoveFromOps.Add($opB)
        }
    }

    # Remove resolved/skipped ops
    foreach ($dead in $toRemoveFromOps) { $ops.Remove($dead) | Out-Null }

    # --- Also detect collision with an existing file that isn't being renamed -
    $stillConflicts = [System.Collections.Generic.List[string]]::new()
    foreach ($op in $ops) {
        if ($op.Type -ne 'file') { continue }
        if ([System.IO.File]::Exists($op.NewPath) -and $op.NewPath -ne $op.OldPath) {
            $stillConflicts.Add("$([System.IO.Path]::GetFileName($op.OldPath))  ->  $([System.IO.Path]::GetFileName($op.NewPath))  (target already exists on disk)")
        }
    }
    if ($stillConflicts.Count -gt 0) {
        $msg = "The following targets already exist on disk and would be overwritten.`nResolve these manually before proceeding:`n`n" + ($stillConflicts -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, 'Cannot Proceed', 'OK', 'Warning') | Out-Null
        return
    }

    # Parent folders
    foreach ($ppKey in $script:parentRows.Keys) {
        $newName = script:ComputeParentNew $ppKey
        if ($newName -eq '' -or -not [System.IO.Directory]::Exists($ppKey)) { continue }
        $parent  = [System.IO.Path]::GetDirectoryName($ppKey)
        $newPath = [System.IO.Path]::Combine($parent, $newName)
        if ($ppKey -ne $newPath) {
            $ops.Add([PSCustomObject]@{ Type = 'parent'; OldPath = $ppKey; NewPath = $newPath })
        }
    }

    # Grandparent folders
    foreach ($gk in $script:gpRows.Keys) {
        $newName = script:NoSpaces $script:gpRows[$gk].NewBox.Text.Trim()
        if ($newName -eq '' -or $gk -eq '__ROOT__' -or -not [System.IO.Directory]::Exists($gk)) { continue }
        $parent  = [System.IO.Path]::GetDirectoryName($gk)
        $newPath = [System.IO.Path]::Combine($parent, $newName)
        if ($gk -ne $newPath) {
            $ops.Add([PSCustomObject]@{ Type = 'gp'; OldPath = $gk; NewPath = $newPath })
        }
    }

    if ($ops.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Nothing to rename - all names already match.', 'No Changes', 'OK', 'Information') | Out-Null
        return
    }

    $summary = "About to rename $($ops.Count) item(s).`n`n"
    $maxShow = [Math]::Min($ops.Count, 14)
    for ($i = 0; $i -lt $maxShow; $i++) {
        $o = $ops[$i]
        $summary += [System.IO.Path]::GetFileName($o.OldPath) + "`n  -> " + [System.IO.Path]::GetFileName($o.NewPath) + "`n"
    }
    if ($ops.Count -gt 14) { $summary += "... and $($ops.Count - 14) more`n" }
    $summary += "`nProceed?"

    if ([System.Windows.Forms.MessageBox]::Show($summary, 'Confirm Rename', 'YesNo', 'Question') -ne 'Yes') { return }

    foreach ($typeOrder in @('file', 'parent', 'gp')) {
        foreach ($op in $ops) {
            if ($op.Type -ne $typeOrder) { continue }
            try {
                if ($op.Type -eq 'file') {
                    if ([System.IO.File]::Exists($op.OldPath)) { [System.IO.File]::Move($op.OldPath, $op.NewPath); $renamed++ }
                } else {
                    if ([System.IO.Directory]::Exists($op.OldPath)) { [System.IO.Directory]::Move($op.OldPath, $op.NewPath); $renamed++ }
                }
            } catch {
                $errors.Add("$([System.IO.Path]::GetFileName($op.OldPath)): $($_.Exception.Message)")
            }
        }
    }

    if ($errors.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Successfully renamed $renamed item(s).`n`nUse Clear All and re-scan to reload.", 'Done', 'OK', 'Information') | Out-Null
    } else {
        $msg = "Renamed $renamed item(s) with $($errors.Count) error(s):`n`n" + ($errors -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, 'Partial Success', 'OK', 'Warning') | Out-Null
    }
}

# --- Build scrollable panel --------------------------------------------------
function script:RebuildPanel {
    # Save user-typed values (keyed per parent folder)
    $savedChar  = @{}; $savedAdj = @{}; $savedGpNew = @{}
    foreach ($ppKey in $script:parentRows.Keys) {
        $r = $script:parentRows[$ppKey]
        $savedChar[$ppKey] = $r.Char.Text
        $savedAdj[$ppKey]  = $r.Adj.Text
    }
    foreach ($gk in $script:gpRows.Keys) { $savedGpNew[$gk] = $script:gpRows[$gk].NewBox.Text }

    $script:gpRows     = @{}
    $script:parentRows = @{}
    $script:resizeList.Clear()

    $scroll.SuspendLayout()
    $scroll.Controls.Clear()

    if ($script:db.Count -eq 0) {
        $lbl           = New-Object System.Windows.Forms.Label
        $lbl.Text      = 'No results yet - browse or drop a folder above to begin.'
        $lbl.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
        $lbl.ForeColor = $cMuted; $lbl.AutoSize = $true
        $lbl.Location  = New-Object System.Drawing.Point(20, 20)
        $scroll.Controls.Add($lbl)
        $scroll.ResumeLayout()
        script:UpdateRenameButton
        return
    }

    $innerW = [Math]::Max(620, $scroll.ClientSize.Width - 8)
    $y = 6

    foreach ($gpKey in $script:db.Keys) {
        $gd = $script:db[$gpKey]
        $totalFiles = 0
        foreach ($pk in $gd.Parents.Keys) { $totalFiles += $gd.Parents[$pk].Files.Count }
        $badge = if ($totalFiles -eq 1) { '1 file' } else { "$totalFiles files" }

        $gpRow = @{
            NewBox      = $null
            ThemeLabels = [System.Collections.Generic.List[System.Windows.Forms.Label]]::new()
            ParentKeys  = [System.Collections.Generic.List[string]]::new()
            OldName     = $gd.Name
        }
        $script:gpRows[$gpKey] = $gpRow

        # -- Grandparent header (56px) ----------------------------------------
        $gpPanel           = New-Object System.Windows.Forms.Panel
        $gpPanel.BackColor = $cBG3
        $gpPanel.Size      = New-Object System.Drawing.Size($innerW, 56)
        $gpPanel.Location  = New-Object System.Drawing.Point(4, $y)

        $gpNameLbl           = New-Object System.Windows.Forms.Label
        $gpNameLbl.Text      = "  $($gd.Name)"
        $gpNameLbl.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
        $gpNameLbl.ForeColor = $cAmber; $gpNameLbl.AutoSize = $true
        $gpNameLbl.Location  = New-Object System.Drawing.Point(4, 6)
        $gpPanel.Controls.Add($gpNameLbl)

        $gpBadge           = New-Object System.Windows.Forms.Label
        $gpBadge.Text      = "[$badge]"
        $gpBadge.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpBadge.ForeColor = $cMuted; $gpBadge.AutoSize = $true
        $gpBadge.Location  = New-Object System.Drawing.Point(($gpNameLbl.PreferredWidth + 10), 8)
        $gpPanel.Controls.Add($gpBadge)

        $gpOldTag = New-Object System.Windows.Forms.Label
        $gpOldTag.Text = 'OLD:'; $gpOldTag.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $gpOldTag.ForeColor = $cOldTag; $gpOldTag.AutoSize = $true
        $gpOldTag.Location  = New-Object System.Drawing.Point(8, 36)
        $gpPanel.Controls.Add($gpOldTag)

        $gpOldVal = New-Object System.Windows.Forms.Label
        $gpOldVal.Text = $gd.Name; $gpOldVal.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpOldVal.ForeColor = $cOldText; $gpOldVal.AutoSize = $true
        $gpOldVal.Location  = New-Object System.Drawing.Point(38, 36)
        $gpPanel.Controls.Add($gpOldVal)

        $gpNewTag = New-Object System.Windows.Forms.Label
        $gpNewTag.Text = 'NEW:'; $gpNewTag.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $gpNewTag.ForeColor = $cNewTag; $gpNewTag.AutoSize = $true
        $gpNewTag.Location  = New-Object System.Drawing.Point(250, 36)
        $gpPanel.Controls.Add($gpNewTag)

        $gpNewBox                  = New-Object System.Windows.Forms.TextBox
        $gpNewBox.Text             = if ($savedGpNew.ContainsKey($gpKey)) { $savedGpNew[$gpKey] } else { $gd.Name }
        $gpNewBox.Font             = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpNewBox.BackColor        = $cInput; $gpNewBox.ForeColor = $cNewText
        $gpNewBox.BorderStyle      = 'FixedSingle'
        $gpNewBox.Size             = New-Object System.Drawing.Size(($innerW - 290), 20)
        $gpNewBox.Location         = New-Object System.Drawing.Point(282, 33)
        $gpNewBox.Tag              = $gpKey
        $gpPanel.Controls.Add($gpNewBox)
        $gpRow.NewBox = $gpNewBox

        $script:resizeList.Add(@{ Type="gp"; Panel=$gpPanel; NewBox=$gpNewBox })

        $gpNewBox.Add_TextChanged({
            param($s, $e)
            $cur = $s.SelectionStart; $old = $s.Text; $n = $old -replace '[^a-zA-Z0-9]', ''
            if ($old -ne $n) { $s.Text = $n; $s.SelectionStart = [Math]::Max(0, $cur - ($old.Length - $n.Length)); $s.SelectionLength = 0 }
            script:UpdateFromGp $s.Tag
        })

        $scroll.Controls.Add($gpPanel)
        $y += 60

        foreach ($ppKey in $gd.Parents.Keys) {
            $pd = $gd.Parents[$ppKey]
            $gpRow.ParentKeys.Add($ppKey)

            # SmartFill defaults
            $fill        = script:SmartFill $ppKey $gd.Name
            $defaultChar = if ($savedChar.ContainsKey($ppKey)) { $savedChar[$ppKey] } else { $fill.Char }
            $defaultAdj  = if ($savedAdj.ContainsKey($ppKey))  { $savedAdj[$ppKey]  } else { $fill.Adj  }

            $fileCount = $pd.Files.Count
            # Heights: name=22, old/new folder=20, col labels=16, inputs=28, divider=5, files=(38 each), padding=6
            $pBlockH = 97 + ($fileCount * 38)

            # -- Parent block -------------------------------------------------
            $pPanel           = New-Object System.Windows.Forms.Panel
            $pPanel.BackColor = $cBG2
            $pPanel.Size      = New-Object System.Drawing.Size($innerW, $pBlockH)
            $pPanel.Location  = New-Object System.Drawing.Point(4, $y)

            # Parent folder name (y=4)
            $pNameLbl = New-Object System.Windows.Forms.Label
            $pNameLbl.Text = "    $($pd.Name)"
            $pNameLbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $pNameLbl.ForeColor = $cBlue; $pNameLbl.AutoSize = $true
            $pNameLbl.Location  = New-Object System.Drawing.Point(4, 4)
            $pPanel.Controls.Add($pNameLbl)

            # OLD/NEW folder row (y=22)
            $pOldTag = New-Object System.Windows.Forms.Label
            $pOldTag.Text = 'OLD:'; $pOldTag.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
            $pOldTag.ForeColor = $cOldTag; $pOldTag.AutoSize = $true
            $pOldTag.Location  = New-Object System.Drawing.Point(16, 22)
            $pPanel.Controls.Add($pOldTag)

            $pOldVal = New-Object System.Windows.Forms.Label
            $pOldVal.Text = $pd.Name; $pOldVal.Font = New-Object System.Drawing.Font('Segoe UI', 8)
            $pOldVal.ForeColor = $cOldText; $pOldVal.AutoSize = $true
            $pOldVal.Location  = New-Object System.Drawing.Point(46, 22)
            $pPanel.Controls.Add($pOldVal)

            $pNewLbl = New-Object System.Windows.Forms.Label
            $pNewLbl.Text = 'NEW:  (enter fields to preview)'
            $pNewLbl.Font = New-Object System.Drawing.Font('Segoe UI', 8)
            $pNewLbl.ForeColor = $cNewText; $pNewLbl.AutoSize = $true
            $pNewLbl.Location  = New-Object System.Drawing.Point(300, 22)
            $pPanel.Controls.Add($pNewLbl)

            # Column header labels (y=42)
            $lay = script:InputLayout $innerW

            $lChar = New-Object System.Windows.Forms.Label
            $lChar.Text = 'Character  *required'; $lChar.Font = New-Object System.Drawing.Font('Segoe UI', 7)
            $lChar.ForeColor = $cLabel; $lChar.AutoSize = $true
            $lChar.Location  = New-Object System.Drawing.Point($lay.xChar, 42)
            $pPanel.Controls.Add($lChar)

            $lAdj = New-Object System.Windows.Forms.Label
            $lAdj.Text = '(Adj.)  - optional'; $lAdj.Font = New-Object System.Drawing.Font('Segoe UI', 7)
            $lAdj.ForeColor = $cLabel; $lAdj.AutoSize = $true
            $lAdj.Location  = New-Object System.Drawing.Point($lay.xAdj, 42)
            $pPanel.Controls.Add($lAdj)

            $lTheme = New-Object System.Windows.Forms.Label
            $lTheme.Text = 'Theme'; $lTheme.Font = New-Object System.Drawing.Font('Segoe UI', 7)
            $lTheme.ForeColor = $cLabel; $lTheme.AutoSize = $true
            $lTheme.Location  = New-Object System.Drawing.Point($lay.xTheme, 42)
            $pPanel.Controls.Add($lTheme)

            # Input controls (y=56)
            $tbChar = New-Object System.Windows.Forms.TextBox
            $tbChar.BackColor = $cInput; $tbChar.ForeColor = $cText
            $tbChar.BorderStyle = 'FixedSingle'; $tbChar.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $tbChar.Size = New-Object System.Drawing.Size($lay.wChar, 24)
            $tbChar.Location = New-Object System.Drawing.Point($lay.xChar, 56)
            $tbChar.Text = $defaultChar; $tbChar.Tag = $ppKey
            $pPanel.Controls.Add($tbChar)

            $sep1 = New-Object System.Windows.Forms.Label
            $sep1.Text = '_'; $sep1.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
            $sep1.ForeColor = $cMuted; $sep1.AutoSize = $true
            $sep1.Location  = New-Object System.Drawing.Point($lay.xSep1, 57)
            $pPanel.Controls.Add($sep1)

            $tbAdj = New-Object System.Windows.Forms.TextBox
            $tbAdj.BackColor = $cInput; $tbAdj.ForeColor = $cText
            $tbAdj.BorderStyle = 'FixedSingle'; $tbAdj.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $tbAdj.Size = New-Object System.Drawing.Size($lay.wAdj, 24)
            $tbAdj.Location = New-Object System.Drawing.Point($lay.xAdj, 56)
            $tbAdj.Text = $defaultAdj; $tbAdj.Tag = $ppKey
            $pPanel.Controls.Add($tbAdj)

            $sep2 = New-Object System.Windows.Forms.Label
            $sep2.Text = '_'; $sep2.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
            $sep2.ForeColor = $cMuted; $sep2.AutoSize = $true
            $sep2.Location  = New-Object System.Drawing.Point($lay.xSep2, 57)
            $pPanel.Controls.Add($sep2)

            $tbTheme = New-Object System.Windows.Forms.Label
            $tbTheme.Text = '  ' + (script:NoSpaces $gpRow.NewBox.Text.Trim())
            $tbTheme.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $tbTheme.BackColor = $cThemeBG; $tbTheme.ForeColor = $cThemeFG
            $tbTheme.BorderStyle = 'FixedSingle'
            $tbTheme.Size = New-Object System.Drawing.Size($lay.wTheme, 24)
            $tbTheme.Location = New-Object System.Drawing.Point($lay.xTheme, 56)
            $tbTheme.TextAlign = 'MiddleLeft'
            $pPanel.Controls.Add($tbTheme)
            $gpRow.ThemeLabels.Add($tbTheme)

            # Divider before file list (y=85)
            $divider = New-Object System.Windows.Forms.Panel
            $divider.BackColor = $cBorder
            $divider.Size = New-Object System.Drawing.Size(($innerW - 16), 1)
            $divider.Location = New-Object System.Drawing.Point(8, 85)
            $pPanel.Controls.Add($divider)

            # File rows (y=88, each 38px)
            $filePreviews = [System.Collections.Generic.List[object]]::new()
            $fy = 88

            foreach ($fi in $pd.Files) {
                # Alternating row background
                $rowBg = if (($filePreviews.Count % 2) -eq 0) { $cBG2 } else { $cFileSep }

                $fRow = New-Object System.Windows.Forms.Panel
                $fRow.BackColor = $rowBg
                $fRow.Size = New-Object System.Drawing.Size(($innerW - 8), 36)
                $fRow.Location = New-Object System.Drawing.Point(4, $fy)
                $pPanel.Controls.Add($fRow)

                $halfW = [int](($innerW - 8 - 60) / 2)

                # Suffix badge
                $suffixBadge = New-Object System.Windows.Forms.Label
                $suffixBadge.Text = $fi.Suffix
                $suffixBadge.Font = New-Object System.Drawing.Font('Consolas', 7, [System.Drawing.FontStyle]::Bold)
                $suffixBadge.ForeColor = $cAmber
                $suffixBadge.BackColor = clr "#1A1600"
                $suffixBadge.AutoSize = $false; $suffixBadge.Size = New-Object System.Drawing.Size(52, 16)
                $suffixBadge.Location = New-Object System.Drawing.Point(4, 10)
                $suffixBadge.TextAlign = 'MiddleCenter'
                $suffixBadge.BorderStyle = 'FixedSingle'
                $fRow.Controls.Add($suffixBadge)

                # OLD name
                $fOldLbl = New-Object System.Windows.Forms.Label
                $fOldLbl.Text = $fi.Name
                $fOldLbl.Font = New-Object System.Drawing.Font('Consolas', 8)
                $fOldLbl.ForeColor = $cOldText; $fOldLbl.AutoSize = $false
                $fOldLbl.Size = New-Object System.Drawing.Size($halfW, 36)
                $fOldLbl.Location = New-Object System.Drawing.Point(60, 0)
                $fOldLbl.TextAlign = 'MiddleLeft'
                $fRow.Controls.Add($fOldLbl)

                # Arrow
                $fArrow = New-Object System.Windows.Forms.Label
                $fArrow.Text = '->'
                $fArrow.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
                $fArrow.ForeColor = $cMuted; $fArrow.AutoSize = $true
                $fArrow.Location = New-Object System.Drawing.Point(($halfW + 62), 10)
                $fRow.Controls.Add($fArrow)

                # NEW name
                $fNewLbl = New-Object System.Windows.Forms.Label
                $fNewLbl.Text = '(enter Character)'
                $fNewLbl.Font = New-Object System.Drawing.Font('Consolas', 8)
                $fNewLbl.ForeColor = $cNewText; $fNewLbl.AutoSize = $false
                $fNewLbl.Size = New-Object System.Drawing.Size($halfW, 36)
                $fNewLbl.Location = New-Object System.Drawing.Point(($halfW + 82), 0)
                $fNewLbl.TextAlign = 'MiddleLeft'
                $fRow.Controls.Add($fNewLbl)

                $filePreviews.Add([PSCustomObject]@{
                    Path      = $fi.Path
                    Suffix    = $fi.Suffix
                    Extension = $fi.Extension
                    NewLabel  = $fNewLbl
                    OldLabel  = $fOldLbl
                    Arrow     = $fArrow
                    RowPanel  = $fRow
                })

                $fy += 38
            }

            # Register parent row
            $pRow = @{
                Char         = $tbChar
                Adj          = $tbAdj
                ThemeLabel   = $tbTheme
                NewLabel     = $pNewLbl
                GpKey        = $gpKey
                OldName      = $pd.Name
                FilePreviews = $filePreviews
            }
            $script:parentRows[$ppKey] = $pRow

            $script:resizeList.Add(@{
                Type    = 'parent'
                Panel   = $pPanel;   Divider = $divider
                Char    = $tbChar;   Adj     = $tbAdj;   Theme = $tbTheme
                Sep1    = $sep1;     Sep2    = $sep2
                LChar   = $lChar;    LAdj    = $lAdj;    LTheme = $lTheme
                Files   = $filePreviews
            })

            # TextChanged: strip non-alphanumeric, propagate
            $tbChar.Add_TextChanged({
                param($s, $e)
                $cur = $s.SelectionStart; $old = $s.Text; $n = $old -replace '[^a-zA-Z0-9]', ''
                if ($old -ne $n) { $s.Text=$n; $s.SelectionStart=[Math]::Max(0,$cur-($old.Length-$n.Length)); $s.SelectionLength=0 }
                script:UpdateFromParent $s.Tag
            })
            $tbAdj.Add_TextChanged({
                param($s, $e)
                $cur = $s.SelectionStart; $old = $s.Text; $n = $old -replace '[^a-zA-Z0-9]', ''
                if ($old -ne $n) { $s.Text=$n; $s.SelectionStart=[Math]::Max(0,$cur-($old.Length-$n.Length)); $s.SelectionLength=0 }
                script:UpdateFromParent $s.Tag
            })

            $scroll.Controls.Add($pPanel)
            $y += ($pBlockH + 4)
        }
        $y += 8
    }

    $scroll.AutoScrollMinSize = New-Object System.Drawing.Size(0, ($y + 10))
    $scroll.ResumeLayout()

    # Fire initial previews
    foreach ($ppKey in $script:parentRows.Keys) { script:UpdateFromParent $ppKey }
    script:UpdateRenameButton
}

function script:RefreshStatus {
    if ($script:db.Count -eq 0) {
        $lblStatus.Text = 'No results.  Browse a folder or drag files onto the window to start.'
        $lblStatus.ForeColor = $cMuted; return
    }
    $fc = 0; $pc = 0
    foreach ($gk in $script:db.Keys) {
        $pc += $script:db[$gk].Parents.Count
        foreach ($pk in $script:db[$gk].Parents.Keys) { $fc += $script:db[$gk].Parents[$pk].Files.Count }
    }
    $gc = $script:db.Count
    $lblStatus.Text = "Found  $fc files  in  $pc folder(s)  across  $gc grandparent group(s)."
    $lblStatus.ForeColor = $cGreen
}

# --- Fonts -------------------------------------------------------------------
$fntTitle  = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
$fntSub    = New-Object System.Drawing.Font('Segoe UI',  9)
$fntBtn    = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$fntStatus = New-Object System.Drawing.Font('Segoe UI',  9)
$fntDrop   = New-Object System.Drawing.Font('Segoe UI', 11)

# --- Form --------------------------------------------------------------------
$form                = New-Object System.Windows.Forms.Form
$form.Text           = '3MF File Finder'
$form.ClientSize     = New-Object System.Drawing.Size(860, 720)
$form.MinimumSize    = New-Object System.Drawing.Size(700, 560)
$form.BackColor      = $cBG; $form.ForeColor = $cText
$form.Font           = $fntSub; $form.StartPosition = 'CenterScreen'
$form.AllowDrop      = $true

$lblTitle            = New-Object System.Windows.Forms.Label
$lblTitle.Text       = '3MF File Finder'; $lblTitle.Font = $fntTitle
$lblTitle.ForeColor  = $cText; $lblTitle.AutoSize = $true
$lblTitle.Location   = New-Object System.Drawing.Point(24, 20)
$form.Controls.Add($lblTitle)

$lblSub              = New-Object System.Windows.Forms.Label
$lblSub.Text         = 'Finds all files in qualifying folders (anchored by Full.3mf) and renames them with a shared Character/Adj/Theme'
$lblSub.Font         = $fntSub; $lblSub.ForeColor = $cMuted; $lblSub.AutoSize = $true
$lblSub.Location     = New-Object System.Drawing.Point(26, 52)
$form.Controls.Add($lblSub)

$sep                 = New-Object System.Windows.Forms.Panel
$sep.Size            = New-Object System.Drawing.Size(812, 1); $sep.Location = New-Object System.Drawing.Point(24, 75)
$sep.BackColor       = $cBorder
$form.Controls.Add($sep)

# --- Drop Zone ---------------------------------------------------------------
$dropPanel              = New-Object System.Windows.Forms.Panel
$dropPanel.Size         = New-Object System.Drawing.Size(812, 128)
$dropPanel.Location     = New-Object System.Drawing.Point(24, 84)
$dropPanel.BackColor    = $cBG2; $dropPanel.AllowDrop = $true

$dropPanel.Add_Paint({
    param($s, $e)
    $e.Graphics.Clear($s.BackColor)
    $pen = New-Object System.Drawing.Pen($cBorder, 1.5)
    $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
    $e.Graphics.DrawRectangle($pen, 1, 1, $s.Width - 3, $s.Height - 3)
    $pen.Dispose()
})
$form.Controls.Add($dropPanel)

$lblDrop              = New-Object System.Windows.Forms.Label
$lblDrop.Text         = 'Drop grandparent folder here (or a Full.3mf file to target its folder)'
$lblDrop.Font         = $fntDrop; $lblDrop.ForeColor = $cMuted; $lblDrop.AutoSize = $false
$lblDrop.Size         = New-Object System.Drawing.Size(812, 34); $lblDrop.TextAlign = 'MiddleCenter'
$lblDrop.Location     = New-Object System.Drawing.Point(0, 8); $lblDrop.AllowDrop = $true
$dropPanel.Controls.Add($lblDrop)

function New-FlatButton($text, $bg, $fg, $borderColor, $w) {
    if (-not $w) { $w = 162 }
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text; $b.Size = New-Object System.Drawing.Size($w, 34); $b.FlatStyle = 'Flat'
    $b.BackColor = $bg; $b.ForeColor = $fg; $b.Font = $fntBtn
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize = if ($borderColor) { 1 } else { 0 }
    if ($borderColor) { $b.FlatAppearance.BorderColor = $borderColor }
    $b.FlatAppearance.MouseOverBackColor = $b.BackColor
    $b.FlatAppearance.MouseDownBackColor = $b.BackColor
    return $b
}

$btnBrowseFolder = New-FlatButton '  Browse Folder' $cAmber $cBG  $null    162
$btnBrowseFile   = New-FlatButton '  Browse File'   $cBG3  $cText $cBorder 162
$btnClear        = New-FlatButton '  Clear All'     $cBG3  $cRed  $cBorder 120
$btnRename       = New-FlatButton '  RENAME'        $cBtnOff $cBtnOffT $cBorder 130

function script:PositionButtons {
    $w = $dropPanel.Width; $total = 162 + 162 + 120 + 130; $gaps = 3
    $sp = [int](($w - $total) / ($gaps + 1))
    $x = $sp
    $btnBrowseFolder.Left = $x; $x += 162 + $sp
    $btnBrowseFile.Left   = $x; $x += 162 + $sp
    $btnClear.Left        = $x; $x += 120 + $sp
    $btnRename.Left       = $x
}

foreach ($b in @($btnBrowseFolder, $btnBrowseFile, $btnClear, $btnRename)) {
    $b.Top = 84; $dropPanel.Controls.Add($b)
}
script:PositionButtons
$btnRename.Enabled = $false; script:StyleRenameBtn $false

$lblStatus            = New-Object System.Windows.Forms.Label
$lblStatus.Text       = 'No results.  Browse a folder or drag files onto the window to start.'
$lblStatus.Font       = $fntStatus; $lblStatus.ForeColor = $cMuted; $lblStatus.AutoSize = $true
$lblStatus.Location   = New-Object System.Drawing.Point(26, 222)
$form.Controls.Add($lblStatus)

$scroll               = New-Object System.Windows.Forms.Panel
$scroll.AutoScroll    = $true; $scroll.BackColor = $cBG; $scroll.BorderStyle = 'None'
$scroll.Location      = New-Object System.Drawing.Point(24, 246)
$scroll.Size          = New-Object System.Drawing.Size(812, 458)
$scroll.Anchor        = [System.Windows.Forms.AnchorStyles]'Top, Bottom, Left, Right'
$form.Controls.Add($scroll)

# --- Resize ------------------------------------------------------------------
$form.Add_Resize({
    $w = $form.ClientSize.Width - 48
    $sep.Width = $w; $dropPanel.Width = $w; $lblDrop.Width = $w
    script:PositionButtons
    $dropPanel.Refresh()
    $scroll.Width  = $w
    $scroll.Height = $form.ClientSize.Height - 268
    script:ApplyResize
})

# --- Drag and Drop -----------------------------------------------------------
$onDragEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        $dropPanel.BackColor = $cDropHL; $dropPanel.Refresh()
    } else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
}
$onDragLeave = { $dropPanel.BackColor = $cBG2; $dropPanel.Refresh() }
$onDragDrop  = {
    param($s, $e)
    $dropPanel.BackColor = $cBG2; $dropPanel.Refresh()
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        foreach ($p in $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) { script:ScanPath $p }
        script:RebuildPanel; script:RefreshStatus
    }
}
foreach ($ctrl in @($form, $dropPanel, $lblDrop)) {
    $ctrl.AllowDrop = $true
    $ctrl.Add_DragEnter($onDragEnter); $ctrl.Add_DragLeave($onDragLeave); $ctrl.Add_DragDrop($onDragDrop)
}

# --- Buttons -----------------------------------------------------------------
$btnBrowseFolder.Add_Click({
    $path = [ModernFolderPicker]::Pick($form.Handle, 'Select a folder to search for Full.3mf files')
    if ($path) { script:ScanPath $path; script:RebuildPanel; script:RefreshStatus }
})
$btnBrowseFile.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = 'Select Full.3mf file(s)'; $ofd.Filter = '3MF Files (*.3mf)|*.3mf|All Files (*.*)|*.*'
    $ofd.Multiselect = $true; $ofd.CheckPathExists = $true
    if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($f in $ofd.FileNames) { script:ScanPath $f }
        script:RebuildPanel; script:RefreshStatus
    }
})
$btnClear.Add_Click({
    $script:db = [System.Collections.Specialized.OrderedDictionary]::new()
    $script:gpRows = @{}; $script:parentRows = @{}
    script:RebuildPanel; script:RefreshStatus
})
$btnRename.Add_Click({ script:DoRename })

# --- Initial render ----------------------------------------------------------
script:RebuildPanel
script:RefreshStatus
[System.Windows.Forms.Application]::Run($form)