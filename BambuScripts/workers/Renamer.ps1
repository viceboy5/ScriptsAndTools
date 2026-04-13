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

. (Join-Path $PSScriptRoot "..\libraries\NamesLibrary.ps1")
$script:validPrefixes = $script:PrinterPrefixes
$script:GpTags        = $script:Tags

# --- Data stores -------------------------------------------------------------
$script:db         = [System.Collections.Specialized.OrderedDictionary]::new()
$script:gpRows     = @{}
$script:parentRows = @{}
$script:resizeList = [System.Collections.Generic.List[hashtable]]::new()

# --- Helpers -----------------------------------------------------------------
function script:NoSpaces([string]$s) { $s -replace '[^a-zA-Z0-9]', '' }

# Enforces CamelCase/PascalCase and strips all non-alphanumeric chars
function script:ToPascalCase([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $parts = $s -split '[_\-\s.]+' | Where-Object { $_ -ne '' }
    $result = ''
    foreach ($p in $parts) {
        if ($p.Length -gt 0) {
            $result += $p.Substring(0,1).ToUpper() + $p.Substring(1)
        }
    }
    return $result -replace '[^a-zA-Z0-9]', ''
}

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

function script:ApplyResize {
    # PERFORMANCE FIX: Freeze UI layout updates while recalculating widths
    $scroll.SuspendLayout()

    $w = [Math]::Max(620, $scroll.ClientSize.Width - 8)
    $lay = script:InputLayout $w

    foreach ($entry in $script:resizeList) {
        if ($entry.Type -eq 'gp') {
            $entry.Panel.Width  = $w
            $entry.NewBox.Width = 140
            $entry.NewLbl.Location = New-Object System.Drawing.Point(520, 36)
            if ($entry.BtnDel) { $entry.BtnDel.Location = New-Object System.Drawing.Point(($w - 110), 12) }
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

            if ($entry.BtnDel) { $entry.BtnDel.Location = New-Object System.Drawing.Point(($w - 110), 10) }

            $halfW = [int](($w - 8 - 140) / 2)
            foreach ($fi in $entry.Files) {
                $fi.RowPanel.Width    = $w - 8
                $fi.OldLabel.Width    = $halfW
                $fi.Arrow.Location    = New-Object System.Drawing.Point((60 + $halfW + 5), $fi.Arrow.Location.Y)
                $fi.NewLabel.Location = New-Object System.Drawing.Point((60 + $halfW + 35), 0)
                $fi.NewLabel.Width    = ($w - 8) - (60 + $halfW + 35) - 40
                if ($fi.BtnDel) { $fi.BtnDel.Location = New-Object System.Drawing.Point(($w - 8 - 35), 15) }
            }
        }
    }

    # PERFORMANCE FIX: Release the layout freeze and draw exactly once
    $scroll.ResumeLayout()
}

function script:IsMatch([string]$name) { $name -imatch 'full\.3mf' }

# --- ParseFile: detect suffix and extension ----------------------------------
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

    if ($ext -ieq '.png') {
        return @{ Suffix = ''; Extension = $ext; Stem = $stem; Parts = $parts }
    }

    $suffix = if ($parts.Count -gt 0) { $parts[-1] } else { $stem }
    return @{ Suffix = $suffix; Extension = $ext; Stem = $stem; Parts = $parts }
}

# --- SmartFill: ruthless array filtering for Character and Adjective -----------
function script:SmartFill([string]$ppKey, [string]$gpTheme) {
    $pd = $null
    foreach ($gk in $script:db.Keys) {
        if ($script:db[$gk].Parents.Contains($ppKey)) { $pd = $script:db[$gk].Parents[$ppKey]; break }
    }
    if (-not $pd) { return @{ Char = ''; Adj = '' } }

    $anchor = $null
    foreach ($fi in $pd.Files) {
        if ($fi.Suffix -ieq 'Full' -and $fi.Extension -ieq '.3mf') { $anchor = $fi; break }
    }
    if (-not $anchor) { return @{ Char = ''; Adj = '' } }

    $stem = $anchor.Stem

    # 1. Brutally strip any valid prefix from ANYWHERE in the filename
    foreach ($p in $script:validPrefixes) {
        $stem = $stem -ireplace "(?i)(^|[_\-\s.])$p([_\-\s.]|$)", "`$1"
    }

    # 2. Strip the "_Full" anchor from the end
    $stem = $stem -ireplace "(?i)[_\-\s.]*Full$", ""

    # 3. Strip the Theme from the string if it exists — check current GP theme
    #    first, then fall back to the full known-themes list so that any theme
    #    name in $script:GpThemes is never picked up as part of the Adjective.
    $themesToStrip = [System.Collections.Generic.List[string]]::new()
    if ($gpTheme -ne '') { $themesToStrip.Add($gpTheme) }
    foreach ($t in $script:GpThemes) {
        $tClean = script:ToPascalCase $t
        if ($tClean -ne '' -and $tClean -ine $gpTheme) { $themesToStrip.Add($tClean) }
    }
    foreach ($t in $themesToStrip) {
        $escapedTheme = [regex]::Escape($t)
        $stem = $stem -ireplace "(?i)[_\-\s.]*$escapedTheme\b", ""
    }

    # 4. Break the remaining string into chunks at every underscore, space, hyphen, or period
    $chunks = $stem -split '[_\-\s.]+' | Where-Object { $_ -ne '' }

    $char = ''
    $adj  = ''

    # 5. Assign the first chunk to Character, and string all remaining chunks into Adjective
    if ($chunks.Count -ge 1) {
        $char = script:ToPascalCase $chunks
    }
    if ($chunks.Count -ge 2) {
        $adj = script:ToPascalCase ($chunks[1..($chunks.Count - 1)] -join ' ')
    }

    # 6. SELF-HEALING: Strip the duplicated Adjective from the Character name
    if ($adj -ne '' -and $char.Length -gt $adj.Length) {
        if ($char.EndsWith($adj, [System.StringComparison]::OrdinalIgnoreCase)) {
            $char = $char.Substring(0, $char.Length - $adj.Length)
        }
    }

    return @{ Char = $char; Adj = $adj }
}

function script:RegisterParent([string]$parentPath, $gpDir) {
    $gpKey  = if ($gpDir) { $gpDir.FullName } else { '__ROOT__' }
    $gpName = if ($gpDir) { $gpDir.Name     } else { '(root)' }

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

function script:ScanPath([string]$rawPath) {
    $path = $rawPath.Trim('"').Trim()
    if (Test-Path -LiteralPath $path -PathType Container) {
        try {
            Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { script:IsMatch $_.Name } |
                ForEach-Object { script:RegisterParent $_.Directory.FullName $_.Directory.Parent }
        } catch { }
    } elseif (Test-Path -LiteralPath $path -PathType Leaf) {
        if (script:IsMatch ([System.IO.Path]::GetFileName($path))) {
            $fi = [System.IO.FileInfo]::new($path)
            script:RegisterParent $fi.Directory.FullName $fi.Directory.Parent
        }
    }
}

# --- Name computation --------------------------------------------------------
function script:ComputeFileNew([string]$ppKey, $fp) {
    $pRow  = $script:parentRows[$ppKey]; if (-not $pRow) { return '' }
    $gpRow = $script:gpRows[$pRow.GpKey]

    $prefix = if ($gpRow -and $gpRow.PrefixDrop.Text -ne '') { script:ToPascalCase $gpRow.PrefixDrop.Text } else { '' }
    $theme  = if ($gpRow) { script:ToPascalCase $gpRow.NewBox.Text } else { '' }
    $char   = script:ToPascalCase $pRow.Char.Text
    $adj    = script:ToPascalCase $pRow.Adj.Text
    if ($char -eq '') { return '' }

    $suffix = if ($fp.SuffixBox) { script:ToPascalCase $fp.SuffixBox.Text } else { $fp.Suffix }
    $ext    = $fp.Extension

    $parts  = @()
    if ($prefix -ne '') { $parts += $prefix }
    $parts += $char
    if ($adj    -ne '') { $parts += $adj    }
    if ($theme  -ne '') { $parts += $theme  }
    if ($suffix -ne '') { $parts += $suffix }
    return ($parts -join '_') + $ext
}

function script:ComputeParentNew([string]$ppKey) {
    $pRow  = $script:parentRows[$ppKey]; if (-not $pRow) { return '' }
    $gpRow = $script:gpRows[$pRow.GpKey]

    $prefix = if ($gpRow -and $gpRow.PrefixDrop.Text -ne '') { script:ToPascalCase $gpRow.PrefixDrop.Text } else { '' }
    $theme  = if ($gpRow) { script:ToPascalCase $gpRow.NewBox.Text } else { '' }
    $char   = script:ToPascalCase $pRow.Char.Text
    $adj    = script:ToPascalCase $pRow.Adj.Text

    $parts = [System.Collections.Generic.List[string]]::new()
    if ($prefix -ne '') { $parts.Add($prefix) }
    if ($char  -ne '')  { $parts.Add($char)  }
    if ($adj   -ne '')  { $parts.Add($adj)   }
    if ($theme -ne '')  { $parts.Add($theme) }
    if ($parts.Count -eq 0) { return '' }
    return $parts -join '_'
}

# --- Live update chain -------------------------------------------------------
function script:RefreshGpPreview([string]$gpKey) {
    $gpRow = $script:gpRows[$gpKey]; if (-not $gpRow) { return }
    $newTheme = script:ToPascalCase $gpRow.NewBox.Text
    $prefix   = if ($gpRow.PrefixDrop.Text -ne '') { script:ToPascalCase $gpRow.PrefixDrop.Text } else { '' }

    $gpVal = if ($prefix -ne '' -and $newTheme -ne '') { "${prefix}_${newTheme}" } elseif ($newTheme) { $newTheme } else { '' }
    if ($gpRow.NewLbl) {
        $gpRow.NewLbl.Text = if ($gpVal) { "PREVIEW:  $gpVal" } else { "PREVIEW:  (enter theme)" }
    }
}

function script:UpdateFromGp([string]$gpKey) {
    $gpRow = $script:gpRows[$gpKey]; if (-not $gpRow) { return }
    $newTheme = script:ToPascalCase $gpRow.NewBox.Text
    foreach ($lbl in $gpRow.ThemeLabels) { $lbl.Text = "  $newTheme" }
    foreach ($ppKey in $gpRow.ParentKeys) {
        script:UpdateFromParent $ppKey
    }
    script:RefreshGpPreview $gpKey
    script:UpdateRenameButton
}

function script:UpdateFromParent([string]$ppKey) {
    $pRow = $script:parentRows[$ppKey]; if (-not $pRow) { return }
    $v = script:ComputeParentNew $ppKey
    $pRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { 'NEW:  (enter fields to preview)' }
    foreach ($fp in $pRow.FilePreviews) {
        $fv = script:ComputeFileNew $ppKey $fp
        $fp.NewLabel.Text = if ($fv) { $fv } else { '(enter Character)' }
    }
    script:RefreshGpPreview $pRow.GpKey
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
    if ($script:db.Count -eq 0) { $btnRename.Enabled = $false; script:StyleRenameBtn $false; $btnRename.Text = "  RENAME"; return }

    $ok = $true

    # Check for empty mandatory boxes
    foreach ($gk in $script:gpRows.Keys) {
        if ((script:ToPascalCase $script:gpRows[$gk].NewBox.Text) -eq '') { $ok = $false; break }
    }
    if ($ok) {
        foreach ($ppKey in $script:parentRows.Keys) {
            if ((script:ToPascalCase $script:parentRows[$ppKey].Char.Text) -eq '') { $ok = $false; break }
        }
    }

    # Map all target names to check for internal UI collisions
    $targetCounts = @{}
    foreach ($ppKey in $script:parentRows.Keys) {
        $pRow = $script:parentRows[$ppKey]
        foreach ($fp in $pRow.FilePreviews) {
            $newName = script:ComputeFileNew $ppKey $fp
            if ($newName -ne '') {
                $targetPath = "$ppKey\$newName"
                if (-not $targetCounts.ContainsKey($targetPath)) { $targetCounts[$targetPath] = 0 }
                $targetCounts[$targetPath]++
            }
        }
    }

    # Update colors based on collisions
    $hasCollision = $false
    foreach ($ppKey in $script:parentRows.Keys) {
        $pRow = $script:parentRows[$ppKey]
        foreach ($fp in $pRow.FilePreviews) {
            $newName = script:ComputeFileNew $ppKey $fp
            if ($newName -ne '') {
                $targetPath = "$ppKey\$newName"
                if ($targetCounts[$targetPath] -gt 1) {
                    $fp.NewLabel.ForeColor = clr "#D95F5F" # Red
                    $hasCollision = $true
                } else {
                    $fp.NewLabel.ForeColor = clr "#4CAF72" # Green
                }
            }
        }
    }

    if ($hasCollision) {
        $ok = $false
        $btnRename.Text = "  NAME CONFLICT"
    } else {
        $btnRename.Text = "  RENAME"
    }

    $btnRename.Enabled = $ok
    script:StyleRenameBtn $ok
}


# --- Perform renames ---------------------------------------------------------
function script:DoRename {
    $errors  = [System.Collections.Generic.List[string]]::new()
    $ops     = [System.Collections.Generic.List[object]]::new()
    $renamed = 0

    # Build List of File Ops
    foreach ($ppKey in $script:parentRows.Keys) {
        $pRow = $script:parentRows[$ppKey]
        foreach ($fp in $pRow.FilePreviews) {
            $newName = script:ComputeFileNew $ppKey $fp
            if ($newName -eq '' -or -not (Test-Path -LiteralPath $fp.Path -PathType Leaf)) { continue }
            $dir     = [System.IO.Path]::GetDirectoryName($fp.Path)
            $newPath = [System.IO.Path]::Combine($dir, $newName)
            if ($fp.Path -ne $newPath) {
                $ops.Add([PSCustomObject]@{ Type = 'file'; OldPath = $fp.Path; NewPath = $newPath })
            }
        }
    }

    # Build List of Parent Folder Ops
    foreach ($ppKey in $script:parentRows.Keys) {
        $newName = script:ComputeParentNew $ppKey
        if ($newName -eq '' -or -not (Test-Path -LiteralPath $ppKey -PathType Container)) { continue }
        $parent  = [System.IO.Path]::GetDirectoryName($ppKey)
        $newPath = [System.IO.Path]::Combine($parent, $newName)
        if ($ppKey -ne $newPath) {
            $ops.Add([PSCustomObject]@{ Type = 'parent'; OldPath = $ppKey; NewPath = $newPath })
        }
    }

    # Build List of Theme Folder Ops
    foreach ($gk in $script:gpRows.Keys) {
        $gpRow = $script:gpRows[$gk]
        if ($gpRow.SkipChk -and $gpRow.SkipChk.Checked) { continue }

        $newTheme = script:ToPascalCase $gpRow.NewBox.Text
        $prefix   = if ($gpRow.PrefixDrop.Text -ne '') { script:ToPascalCase $gpRow.PrefixDrop.Text } else { '' }

        $newName = if ($prefix -ne '' -and $newTheme -ne '') { "${prefix}_${newTheme}" } else { $newTheme }

        if ($newName -eq '' -or $gk -eq '__ROOT__' -or -not (Test-Path -LiteralPath $gk -PathType Container)) { continue }
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

    $summary = "About to rename $($ops.Count) item(s). Overwrites are ENABLED.`n`n"
    $maxShow = [Math]::Min($ops.Count, 14)
    for ($i = 0; $i -lt $maxShow; $i++) {
        $o = $ops[$i]
        $summary += [System.IO.Path]::GetFileName($o.OldPath) + "`n  -> " + [System.IO.Path]::GetFileName($o.NewPath) + "`n"
    }
    if ($ops.Count -gt 14) { $summary += "... and $($ops.Count - 14) more`n" }
    $summary += "`nProceed?"

    if ([System.Windows.Forms.MessageBox]::Show($summary, 'Confirm Rename', 'YesNo', 'Question') -ne 'Yes') { return }

    # CRITICAL: Gentle Swap execution phase to protect Synology Cloud links
    foreach ($typeOrder in @('file', 'parent', 'gp')) {
        foreach ($op in $ops) {
            if ($op.Type -ne $typeOrder) { continue }
            try {
                $isCaseChangeOnly = ($op.OldPath -ieq $op.NewPath) -and ($op.OldPath -cne $op.NewPath)

                if ($op.Type -eq 'file') {
                    if (Test-Path -LiteralPath $op.OldPath -PathType Leaf) {
                        if ($isCaseChangeOnly) {
                            $tempPath = $op.OldPath + "_TMP"
                            [System.IO.File]::Move($op.OldPath, $tempPath)
                            [System.IO.File]::Move($tempPath, $op.NewPath)
                        } else {
                            if (Test-Path -LiteralPath $op.NewPath -PathType Leaf) {
                                [System.IO.File]::Delete($op.NewPath)
                            }
                            [System.IO.File]::Move($op.OldPath, $op.NewPath)
                        }
                        $renamed++
                    }
                } else {
                    if (Test-Path -LiteralPath $op.OldPath -PathType Container) {
                        if ($isCaseChangeOnly) {
                            $tempPath = $op.OldPath + "_TMP"
                            [System.IO.Directory]::Move($op.OldPath, $tempPath)
                            [System.IO.Directory]::Move($tempPath, $op.NewPath)
                        } else {
                            if (Test-Path -LiteralPath $op.NewPath -PathType Container) {
                                Remove-Item -LiteralPath $op.NewPath -Recurse -Force -ErrorAction Stop
                            }
                            [System.IO.Directory]::Move($op.OldPath, $op.NewPath)
                        }
                        $renamed++
                    }
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
    $savedChar = @{}; $savedAdj = @{}; $savedGpNew = @{}; $savedSuffix = @{}; $savedPrefixGp = @{}; $savedSkipGp = @{}

    foreach ($ppKey in $script:parentRows.Keys) {
        $r = $script:parentRows[$ppKey]
        $savedChar[$ppKey] = $r.Char.Text
        $savedAdj[$ppKey]  = $r.Adj.Text
        foreach ($fp in $r.FilePreviews) {
            if ($fp.SuffixBox) { $savedSuffix[$fp.Path] = $fp.SuffixBox.Text }
        }
    }
    foreach ($gk in $script:gpRows.Keys) {
        $savedGpNew[$gk]  = $script:gpRows[$gk].NewBox.Text
        if ($script:gpRows[$gk].PrefixDrop) { $savedPrefixGp[$gk] = $script:gpRows[$gk].PrefixDrop.Text }
        if ($script:gpRows[$gk].SkipChk) { $savedSkipGp[$gk] = $script:gpRows[$gk].SkipChk.Checked }
    }

    # PERFORMANCE FIX: Grab exact scroll position and reset to 0,0
    $savedScroll = $scroll.AutoScrollPosition
    $scroll.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)

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

        # Strip all leading qualifier tokens (printer prefix + tags) from the GP folder name
        # to isolate the bare theme. Handles any combination, e.g.:
        #   "P2S_KC_Licensing" -> prefix "P2S", theme "Licensing"
        #   "KC_Licensing"     -> prefix "",    theme "Licensing"
        $gpTheme = $gd.Name
        $gpPrefix = ''

        $gpTokens = [System.Collections.Generic.List[string]]($gpTheme -split '[_\-\s.]+' | Where-Object { $_ -ne '' })
        while ($gpTokens.Count -gt 1) {
            $head = $gpTokens[0]
            if ($script:validPrefixes -icontains $head) {
                if ($gpPrefix -eq '') { $gpPrefix = $head }
                $gpTokens.RemoveAt(0)
            } elseif ($script:GpTags -icontains $head) {
                $gpTokens.RemoveAt(0)
            } else { break }
        }
        $gpTheme = $gpTokens -join '_'

        # Enforce alphanumeric immediately on load for the theme
        $gpTheme = script:ToPascalCase $gpTheme

        $gpRow = @{
            PrefixDrop  = $null
            NewBox      = $null
            NewLbl      = $null
            SkipChk     = $null
            ThemeLabels = [System.Collections.Generic.List[System.Windows.Forms.Label]]::new()
            ParentKeys  = [System.Collections.Generic.List[string]]::new()
            OldName     = $gd.Name
        }
        $script:gpRows[$gpKey] = $gpRow

        # -- Grandparent header -----------------------------------------------
        $gpPanel           = New-Object System.Windows.Forms.Panel
        $gpPanel.BackColor = $cBG3
        $gpPanel.Size      = New-Object System.Drawing.Size($innerW, 74)
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

        $gpPrefixTag = New-Object System.Windows.Forms.Label
        $gpPrefixTag.Text = 'PREFIX:'; $gpPrefixTag.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $gpPrefixTag.ForeColor = $cBlue; $gpPrefixTag.AutoSize = $true
        $gpPrefixTag.Location  = New-Object System.Drawing.Point(180, 36)
        $gpPanel.Controls.Add($gpPrefixTag)

        $gpPrefixDrop = New-Object System.Windows.Forms.ComboBox
        $gpPrefixDrop.DropDownStyle = 'DropDownList'
        $gpPrefixDrop.Items.AddRange(@('', 'X1C', 'P2S', 'H2S'))
        $gpPrefixDrop.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpPrefixDrop.BackColor = $cInput; $gpPrefixDrop.ForeColor = $cBlue
        $gpPrefixDrop.Size = New-Object System.Drawing.Size(60, 20)
        $gpPrefixDrop.Location = New-Object System.Drawing.Point(230, 33)
        $gpPrefixDrop.Tag = $gpKey
        $gpPrefixDrop.Text = if ($savedPrefixGp.ContainsKey($gpKey)) { $savedPrefixGp[$gpKey] } else { $gpPrefix }
        $gpPanel.Controls.Add($gpPrefixDrop)
        $gpRow.PrefixDrop = $gpPrefixDrop

        $gpNewTag = New-Object System.Windows.Forms.Label
        $gpNewTag.Text = 'THEME:'; $gpNewTag.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
        $gpNewTag.ForeColor = $cNewTag; $gpNewTag.AutoSize = $true
        $gpNewTag.Location  = New-Object System.Drawing.Point(300, 36)
        $gpPanel.Controls.Add($gpNewTag)

        $gpNewBox                  = New-Object System.Windows.Forms.TextBox
        $gpNewBox.Text             = if ($savedGpNew.ContainsKey($gpKey)) { script:ToPascalCase $savedGpNew[$gpKey] } else { $gpTheme }
        $gpNewBox.Font             = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpNewBox.BackColor        = $cInput; $gpNewBox.ForeColor = $cNewText
        $gpNewBox.BorderStyle      = 'FixedSingle'
        $gpNewBox.Size             = New-Object System.Drawing.Size(140, 20)
        $gpNewBox.Location         = New-Object System.Drawing.Point(345, 33)
        $gpNewBox.Tag              = $gpKey
        $gpPanel.Controls.Add($gpNewBox)
        $gpRow.NewBox = $gpNewBox

        $gpNewLbl = New-Object System.Windows.Forms.Label
        $gpNewLbl.Text = "PREVIEW:  (enter theme)"
        $gpNewLbl.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpNewLbl.ForeColor = $cNewText; $gpNewLbl.AutoSize = $true
        $gpNewLbl.Location = New-Object System.Drawing.Point(520, 36)
        $gpPanel.Controls.Add($gpNewLbl)
        $gpRow.NewLbl = $gpNewLbl

        $gpSkipChk                  = New-Object System.Windows.Forms.CheckBox
        $gpSkipChk.Text             = "Don't rename this folder  (use Theme name for files/subfolders only)"
        $gpSkipChk.Font             = New-Object System.Drawing.Font('Segoe UI', 8)
        $gpSkipChk.ForeColor        = $cMuted
        $gpSkipChk.AutoSize         = $true
        $gpSkipChk.Location         = New-Object System.Drawing.Point(10, 54)
        $gpSkipChk.FlatStyle        = 'Flat'
        $gpSkipChk.BackColor        = $cBG3
        if ($savedSkipGp.ContainsKey($gpKey)) { $gpSkipChk.Checked = $savedSkipGp[$gpKey] }
        $gpPanel.Controls.Add($gpSkipChk)
        $gpRow.SkipChk = $gpSkipChk

        # --- GRANDPARENT REMOVE BUTTON ---
        $btnDelGp = New-Object System.Windows.Forms.Button
        $btnDelGp.Text = "Remove Theme"
        $btnDelGp.BackColor = clr "#D95F5F"
        $btnDelGp.ForeColor = clr "#FFFFFF"
        $btnDelGp.FlatStyle = 'Flat'
        $btnDelGp.FlatAppearance.BorderSize = 0
        $btnDelGp.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $btnDelGp.Size = New-Object System.Drawing.Size(100, 26)
        $btnDelGp.Location = New-Object System.Drawing.Point(($innerW - 110), 12)
        $btnDelGp.Anchor = 'Top, Right'
        $gpPanel.Controls.Add($btnDelGp)

        $btnDelGp.Tag = @{ GpKey = $gpKey }
        $btnDelGp.Add_Click({
            $t = $this.Tag
            $script:db.Remove($t.GpKey)
            script:RebuildPanel
            script:RefreshStatus
        })

        $script:resizeList.Add(@{ Type="gp"; Panel=$gpPanel; NewBox=$gpNewBox; NewLbl=$gpNewLbl; PrefixDrop=$gpPrefixDrop; BtnDel=$btnDelGp })

        $gpNewBox.Add_TextChanged({
            param($s, $e)
            $old = $s.Text
            $n = script:ToPascalCase $old
            if ($old -ne $n) {
                $s.Text = $n
                $s.SelectionStart = $s.Text.Length
                $s.SelectionLength = 0
            }
            script:UpdateFromGp $s.Tag
        })

        $gpPrefixDrop.Add_SelectedIndexChanged({
            param($s, $e)
            script:UpdateFromGp $s.Tag
        })

        $scroll.Controls.Add($gpPanel)
        $y += 78

        foreach ($ppKey in $gd.Parents.Keys) {
            $pd = $gd.Parents[$ppKey]
            $gpRow.ParentKeys.Add($ppKey)

            $fill         = script:SmartFill $ppKey $gpTheme
            $defaultChar  = if ($savedChar.ContainsKey($ppKey)) { script:ToPascalCase $savedChar[$ppKey] } else { $fill.Char }
            $defaultAdj   = if ($savedAdj.ContainsKey($ppKey))  { script:ToPascalCase $savedAdj[$ppKey]  } else { $fill.Adj  }

            $fileCount = $pd.Files.Count
            $pBlockH = 97 + ($fileCount * 54)

            # -- Parent block -------------------------------------------------
            $pPanel           = New-Object System.Windows.Forms.Panel
            $pPanel.BackColor = $cBG2
            $pPanel.Size      = New-Object System.Drawing.Size($innerW, $pBlockH)
            $pPanel.Location  = New-Object System.Drawing.Point(4, $y)

            $pNameLbl = New-Object System.Windows.Forms.Label
            $pNameLbl.Text = "    $($pd.Name)"
            $pNameLbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $pNameLbl.ForeColor = $cBlue; $pNameLbl.AutoSize = $true
            $pNameLbl.Location  = New-Object System.Drawing.Point(4, 4)
            $pPanel.Controls.Add($pNameLbl)

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
            $tbTheme.Text = '  ' + (script:ToPascalCase $gpRow.NewBox.Text)
            $tbTheme.Font = New-Object System.Drawing.Font('Segoe UI', 9)
            $tbTheme.BackColor = $cThemeBG; $tbTheme.ForeColor = $cThemeFG
            $tbTheme.BorderStyle = 'FixedSingle'
            $tbTheme.Size = New-Object System.Drawing.Size($lay.wTheme, 24)
            $tbTheme.Location = New-Object System.Drawing.Point($lay.xTheme, 56)
            $tbTheme.TextAlign = 'MiddleLeft'
            $pPanel.Controls.Add($tbTheme)
            $gpRow.ThemeLabels.Add($tbTheme)

            # --- PARENT REMOVE BUTTON ---
            $btnDelParent = New-Object System.Windows.Forms.Button
            $btnDelParent.Text = "Remove Folder"
            $btnDelParent.BackColor = clr "#D95F5F"
            $btnDelParent.ForeColor = clr "#FFFFFF"
            $btnDelParent.FlatStyle = 'Flat'
            $btnDelParent.FlatAppearance.BorderSize = 0
            $btnDelParent.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
            $btnDelParent.Size = New-Object System.Drawing.Size(100, 24)
            $btnDelParent.Location = New-Object System.Drawing.Point(($innerW - 110), 10)
            $btnDelParent.Anchor = 'Top, Right'
            $pPanel.Controls.Add($btnDelParent)

            $btnDelParent.Tag = @{ GpKey = $gpKey; ParentKey = $ppKey }
            $btnDelParent.Add_Click({
                $t = $this.Tag
                $script:db[$t.GpKey].Parents.Remove($t.ParentKey)
                if ($script:db[$t.GpKey].Parents.Count -eq 0) {
                    $script:db.Remove($t.GpKey)
                }
                script:RebuildPanel
                script:RefreshStatus
            })

            $divider = New-Object System.Windows.Forms.Panel
            $divider.BackColor = $cBorder
            $divider.Size = New-Object System.Drawing.Size(($innerW - 16), 1)
            $divider.Location = New-Object System.Drawing.Point(8, 85)
            $pPanel.Controls.Add($divider)

            $filePreviews = [System.Collections.Generic.List[object]]::new()
            $fy = 88

            foreach ($fi in $pd.Files) {
                # Suffix detection
                $fSuffix = $fi.Suffix # Default mapped from ParseFile
                $fParts  = $fi.Stem -split '_'
                $themeIdx = -1
                for ($i = $fParts.Count - 1; $i -ge 0; $i--) {
                    if ($fParts[$i] -ieq $gpTheme) { $themeIdx = $i; break }
                }
                if ($themeIdx -gt 0 -and $themeIdx -lt $fParts.Count - 1) {
                    $fSuffix = $fParts[($themeIdx + 1)..($fParts.Count - 1)] -join '_'
                }

                $rowBg = if (($filePreviews.Count % 2) -eq 0) { $cBG2 } else { $cFileSep }

                $fRow = New-Object System.Windows.Forms.Panel
                $fRow.BackColor = $rowBg
                $fRow.Size = New-Object System.Drawing.Size(($innerW - 8), 52)
                $fRow.Location = New-Object System.Drawing.Point(4, $fy)
                $pPanel.Controls.Add($fRow)

                $halfW = [int](($innerW - 8 - 140) / 2)

                # Suffix Badge & Box
                $suffixBadge = New-Object System.Windows.Forms.Label
                $suffixBadge.Text = if ($fi.Suffix -ne '') { $fi.Suffix } else { '(none)' }
                $suffixBadge.Font = New-Object System.Drawing.Font('Consolas', 7, [System.Drawing.FontStyle]::Bold)
                $suffixBadge.ForeColor = $cAmber
                $suffixBadge.BackColor = clr "#1A1600"
                $suffixBadge.AutoSize = $false; $suffixBadge.Size = New-Object System.Drawing.Size(52, 16)
                $suffixBadge.Location = New-Object System.Drawing.Point(4, 4)
                $suffixBadge.TextAlign = 'MiddleCenter'
                $suffixBadge.BorderStyle = 'FixedSingle'
                $fRow.Controls.Add($suffixBadge)

                $suffixBox = New-Object System.Windows.Forms.TextBox
                $suffixBox.Font = New-Object System.Drawing.Font('Consolas', 7)
                $suffixBox.BackColor = $cInput; $suffixBox.ForeColor = $cAmber
                $suffixBox.BorderStyle = 'FixedSingle'
                $suffixBox.Size = New-Object System.Drawing.Size(52, 18)
                $suffixBox.Location = New-Object System.Drawing.Point(4, 24)
                $suffixBox.TextAlign = 'Center'
                $suffixBox.Text = if ($savedSuffix.ContainsKey($fi.Path)) { script:ToPascalCase $savedSuffix[$fi.Path] } else { script:ToPascalCase $fSuffix }
                $suffixBox.Tag = $ppKey
                $fRow.Controls.Add($suffixBox)

                # OLD name
                $fOldLbl = New-Object System.Windows.Forms.Label
                $fOldLbl.Text = $fi.Name
                $fOldLbl.Font = New-Object System.Drawing.Font('Consolas', 8)
                $fOldLbl.ForeColor = $cOldText; $fOldLbl.AutoSize = $false
                $fOldLbl.Size = New-Object System.Drawing.Size($halfW, 52)
                $fOldLbl.Location = New-Object System.Drawing.Point(60, 0)
                $fOldLbl.TextAlign = 'MiddleLeft'
                $fRow.Controls.Add($fOldLbl)

                # Arrow
                $fArrow = New-Object System.Windows.Forms.Label
                $fArrow.Text = '->'
                $fArrow.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
                $fArrow.ForeColor = $cMuted; $fArrow.AutoSize = $true
                $fArrow.Location = New-Object System.Drawing.Point((60 + $halfW + 5), 19)
                $fRow.Controls.Add($fArrow)

                # NEW name
                $fNewLbl = New-Object System.Windows.Forms.Label
                $fNewLbl.Text = '(enter Character)'
                $fNewLbl.Font = New-Object System.Drawing.Font('Consolas', 8)
                $fNewLbl.ForeColor = $cNewText; $fNewLbl.AutoSize = $false
                $fNewLbl.Size = New-Object System.Drawing.Size((($innerW - 8) - (60 + $halfW + 35) - 40), 52)
                $fNewLbl.Location = New-Object System.Drawing.Point((60 + $halfW + 35), 0)
                $fNewLbl.TextAlign = 'MiddleLeft'
                $fRow.Controls.Add($fNewLbl)

                # --- FILE DELETE BUTTON ---
                $btnDel = New-Object System.Windows.Forms.Button
                $btnDel.Text = "X"
                $btnDel.BackColor = clr "#D95F5F"
                $btnDel.ForeColor = clr "#FFFFFF"
                $btnDel.FlatStyle = 'Flat'
                $btnDel.FlatAppearance.BorderSize = 0
                $btnDel.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
                $btnDel.Size = New-Object System.Drawing.Size(20, 20)
                $btnDel.Location = New-Object System.Drawing.Point(($innerW - 8 - 35), 15)
                $btnDel.Anchor = 'Top, Right'
                $fRow.Controls.Add($btnDel)
                $btnDel.BringToFront()

                $btnDel.Tag = @{ Path = $fi.Path; ParentKey = $ppKey; GpKey = $gpKey }
                $btnDel.Add_Click({
                    $t = $this.Tag
                    $res = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to permanently delete:`n$($t.Path | Split-Path -Leaf)?", "Confirm File Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($res -eq 'Yes') {
                        try {
                            Remove-Item -LiteralPath $t.Path -Force -ErrorAction Stop
                            $parentData = $script:db[$t.GpKey].Parents[$t.ParentKey]
                            $fileObj = $parentData.Files | Where-Object { $_.Path -eq $t.Path }
                            if ($fileObj) { $parentData.Files.Remove($fileObj) | Out-Null }
                            script:RebuildPanel
                            script:RefreshStatus
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("Failed to delete file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                        }
                    }
                })

                $filePreviews.Add([PSCustomObject]@{
                    Path      = $fi.Path
                    Suffix    = $fi.Suffix
                    SuffixBox = $suffixBox
                    Extension = $fi.Extension
                    NewLabel  = $fNewLbl
                    OldLabel  = $fOldLbl
                    Arrow     = $fArrow
                    RowPanel  = $fRow
                    BtnDel    = $btnDel
                })

                $suffixBox.Add_TextChanged({
                    param($s, $e)
                    $old = $s.Text
                    $n = script:ToPascalCase $old
                    if ($old -ne $n) {
                        $s.Text=$n
                        $s.SelectionStart = $s.Text.Length
                        $s.SelectionLength=0
                    }
                    script:UpdateFromParent $s.Tag
                })

                $fy += 54
            }

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
                BtnDel  = $btnDelParent
            })

            $tbChar.Add_TextChanged({
                param($s, $e)
                $old = $s.Text
                $n = script:ToPascalCase $old
                if ($old -ne $n) {
                    $s.Text=$n
                    $s.SelectionStart = $s.Text.Length
                    $s.SelectionLength=0
                }
                script:UpdateFromParent $s.Tag
            })
            $tbAdj.Add_TextChanged({
                param($s, $e)
                $old = $s.Text
                $n = script:ToPascalCase $old
                if ($old -ne $n) {
                    $s.Text=$n
                    $s.SelectionStart = $s.Text.Length
                    $s.SelectionLength=0
                }
                script:UpdateFromParent $s.Tag
            })

            $scroll.Controls.Add($pPanel)
            $y += ($pBlockH + 4)
        }
        $y += 8
    }

    $scroll.AutoScrollMinSize = New-Object System.Drawing.Size(0, ($y + 10))
    $scroll.ResumeLayout()

    # PERFORMANCE FIX: Restore exact scroll position
    $scroll.AutoScrollPosition = New-Object System.Drawing.Point([Math]::Abs($savedScroll.X), [Math]::Abs($savedScroll.Y))

    foreach ($gpKey in $script:gpRows.Keys) { script:RefreshGpPreview $gpKey }
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
    # PERFORMANCE FIX: Suspend the entire form layout during resize dragging
    $form.SuspendLayout()

    $w = $form.ClientSize.Width - 48
    $sep.Width = $w; $dropPanel.Width = $w; $lblDrop.Width = $w
    script:PositionButtons

    # PERFORMANCE FIX: Invalidate is much cheaper than Refresh
    $dropPanel.Invalidate()

    $scroll.Width  = $w
    $scroll.Height = $form.ClientSize.Height - 268
    script:ApplyResize

    # PERFORMANCE FIX: Resume layout and draw exactly once
    $form.ResumeLayout()
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