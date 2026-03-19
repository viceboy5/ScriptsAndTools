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

# --- Data stores -------------------------------------------------------------
$script:db         = [System.Collections.Specialized.OrderedDictionary]::new()
$script:fileRows   = @{}
$script:gpRows     = @{}
$script:parentRows = @{}

# --- Extract default Character from filename ---------------------------------
# Grabs everything before the last "_Full" (case-insensitive) in the stem
# Returns @{ Char = "..."; Adj = "..." } pre-filled intelligently.
#
# Already-renamed pattern: Character[_Adj]_Theme_Full.3mf
# After stripping _Theme_Full the prefix will be CamelCase (no underscores in
# Character or Adj because the script enforces that rule), so:
#   - 1 segment  -> Char=segment,         Adj=""
#   - 2 segments -> Char=segments[0],     Adj=segments[1]
#
# Original (un-renamed) files: extract everything before _Full as Character.
function script:SmartFill([string]$filename, [string]$gpName) {
    $stem    = [System.IO.Path]::GetFileNameWithoutExtension($filename)
    $escaped = [regex]::Escape($gpName)

    # Already renamed: ends with _Theme_Full (case-insensitive)
    if ($gpName -ne "" -and $stem -imatch "^(.+)_${escaped}_Full$") {
        $prefix   = $Matches[1]
        $segments = $prefix -split '_'
        if ($segments.Count -ge 2) {
            return @{ Char = $segments[0]; Adj = ($segments[1..($segments.Count-1)] -join '_') }
        } else {
            return @{ Char = $prefix; Adj = "" }
        }
    }

    # Original file: grab everything before the first _Full
    if ($stem -imatch '^(.+?)_Full') {
        return @{ Char = $Matches[1]; Adj = "" }
    }

    # Fallback
    return @{ Char = $stem; Adj = "" }
}

# --- Layout math -------------------------------------------------------------
# Compute responsive x/width for the three input boxes given panel inner width
function script:InputLayout([int]$panelW) {
    $leftMargin  = 12
    $rightMargin = 12
    $sepW        = 16   # width of each "_" separator label
    $gap         = 4    # gap between field and separator
    $avail = $panelW - $leftMargin - $rightMargin - ($sepW * 2) - ($gap * 4)
    $wChar  = [int]($avail * 0.36)
    $wAdj   = [int]($avail * 0.28)
    $wTheme = $avail - $wChar - $wAdj

    $xChar  = $leftMargin
    $xSep1  = $xChar + $wChar + $gap
    $xAdj   = $xSep1 + $sepW + $gap
    $xSep2  = $xAdj + $wAdj + $gap
    $xTheme = $xSep2 + $sepW + $gap

    return @{
        xChar  = $xChar;  wChar  = $wChar
        xSep1  = $xSep1
        xAdj   = $xAdj;   wAdj   = $wAdj
        xSep2  = $xSep2
        xTheme = $xTheme; wTheme = $wTheme
    }
}

# --- Utility: no spaces ------------------------------------------------------
function script:NoSpaces([string]$s) { $s -replace '[^a-zA-Z0-9]', '' }

# --- Scan helpers ------------------------------------------------------------
function script:IsMatch([string]$name) { $name -imatch 'full\.3mf' }

function script:AddFile([string]$filePath) {
    try {
        $fi     = [System.IO.FileInfo]::new($filePath)
        $par    = $fi.Directory
        $gp     = $par.Parent
        $gpKey  = if ($gp) { $gp.FullName } else { "__ROOT__" }
        $gpName = if ($gp) { $gp.Name     } else { "(root)" }
        if (-not $script:db.Contains($gpKey)) {
            $script:db[$gpKey] = [PSCustomObject]@{
                Name    = $gpName
                Path    = $gpKey
                Parents = [System.Collections.Specialized.OrderedDictionary]::new()
            }
        }
        $gd = $script:db[$gpKey]
        if (-not $gd.Parents.Contains($par.FullName)) {
            $gd.Parents[$par.FullName] = [PSCustomObject]@{
                Name  = $par.Name
                Path  = $par.FullName
                Files = [System.Collections.Generic.List[string]]::new()
            }
        }
        $pd = $gd.Parents[$par.FullName]
        if (-not ($pd.Files -contains $fi.FullName)) { $pd.Files.Add($fi.FullName) }
    } catch { }
}

function script:ScanPath([string]$rawPath) {
    $path = $rawPath.Trim('"').Trim()
    if ([System.IO.Directory]::Exists($path)) {
        try {
            Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { script:IsMatch $_.Name } |
                ForEach-Object { script:AddFile $_.FullName }
        } catch { }
    } elseif ([System.IO.File]::Exists($path)) {
        if (script:IsMatch ([System.IO.Path]::GetFileName($path))) {
            script:AddFile $path
        }
    }
}

# --- Name computation --------------------------------------------------------
function script:ComputeFileNew([string]$fp) {
    $fRow  = $script:fileRows[$fp]; if (-not $fRow) { return "" }
    $gpRow = $script:gpRows[$fRow.GpKey]
    $theme = if ($gpRow) { script:NoSpaces $gpRow.NewBox.Text.Trim() } else { "" }
    $char  = script:NoSpaces $fRow.Char.Text.Trim()
    $adj   = script:NoSpaces $fRow.Adj.Text.Trim()
    if ($char -eq "") { return "" }
    $parts = @($char)
    if ($adj   -ne "") { $parts += $adj }
    if ($theme -ne "") { $parts += $theme }
    return ($parts -join "_") + "_Full.3mf"
}

function script:ComputeParentNew([string]$ppKey) {
    $pRow  = $script:parentRows[$ppKey]; if (-not $pRow) { return "" }
    $gpRow = $script:gpRows[$pRow.GpKey]
    $theme = if ($gpRow) { script:NoSpaces $gpRow.NewBox.Text.Trim() } else { "" }
    $firstFp = if ($pRow.FilePaths.Count -gt 0) { $pRow.FilePaths[0] } else { $null }
    $char = ""; $adj = ""
    if ($firstFp -and $script:fileRows.ContainsKey($firstFp)) {
        $fRow = $script:fileRows[$firstFp]
        $char = script:NoSpaces $fRow.Char.Text.Trim()
        $adj  = script:NoSpaces $fRow.Adj.Text.Trim()
    }
    # Order: Character_Adj_Theme  (same as file, without _Full.3mf)
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($char  -ne "") { $parts.Add($char)  }
    if ($adj   -ne "") { $parts.Add($adj)   }
    if ($theme -ne "") { $parts.Add($theme) }
    if ($parts.Count -eq 0) { return "" }
    return $parts -join "_"
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
            $pRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { "NEW:  (enter fields to preview)" }
        }
    }
    foreach ($fp in $gpRow.FilePaths) {
        $fRow = $script:fileRows[$fp]
        if ($fRow) {
            $v = script:ComputeFileNew $fp
            $fRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { "NEW:  (enter Character to preview)" }
        }
    }
    script:UpdateRenameButton
}

function script:UpdateFromFile([string]$fp) {
    $fRow = $script:fileRows[$fp]; if (-not $fRow) { return }
    $v = script:ComputeFileNew $fp
    $fRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { "NEW:  (enter Character to preview)" }
    $pRow = $script:parentRows[$fRow.ParentKey]
    if ($pRow) {
        $v2 = script:ComputeParentNew $fRow.ParentKey
        $pRow.NewLabel.Text = if ($v2) { "NEW:  $v2" } else { "NEW:  (enter fields to preview)" }
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
        $btnRename.FlatAppearance.MouseOverBackColor = clr "#3A7050"
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
    if ($script:db.Count -eq 0) {
        $btnRename.Enabled = $false; script:StyleRenameBtn $false; return
    }
    $ok = $true
    foreach ($gk in $script:gpRows.Keys) {
        if ((script:NoSpaces $script:gpRows[$gk].NewBox.Text.Trim()) -eq "") { $ok = $false; break }
    }
    if ($ok) {
        foreach ($fp in $script:fileRows.Keys) {
            if ((script:NoSpaces $script:fileRows[$fp].Char.Text.Trim()) -eq "") { $ok = $false; break }
        }
    }
    $btnRename.Enabled = $ok
    script:StyleRenameBtn $ok
}

# --- Perform renames ---------------------------------------------------------
function script:DoRename {
    $errors  = [System.Collections.Generic.List[string]]::new()
    $ops     = [System.Collections.Generic.List[object]]::new()
    $renamed = 0

    foreach ($fp in $script:fileRows.Keys) {
        $newName = script:ComputeFileNew $fp
        if ($newName -eq "" -or -not [System.IO.File]::Exists($fp)) { continue }
        $dir     = [System.IO.Path]::GetDirectoryName($fp)
        $newPath = [System.IO.Path]::Combine($dir, $newName)
        if ($fp -ne $newPath) { $ops.Add([PSCustomObject]@{ Type = "file"; OldPath = $fp; NewPath = $newPath }) }
    }
    foreach ($ppKey in $script:parentRows.Keys) {
        $newName = script:ComputeParentNew $ppKey
        if ($newName -eq "" -or -not [System.IO.Directory]::Exists($ppKey)) { continue }
        $parent  = [System.IO.Path]::GetDirectoryName($ppKey)
        $newPath = [System.IO.Path]::Combine($parent, $newName)
        if ($ppKey -ne $newPath) { $ops.Add([PSCustomObject]@{ Type = "parent"; OldPath = $ppKey; NewPath = $newPath }) }
    }
    foreach ($gk in $script:gpRows.Keys) {
        $newName = script:NoSpaces $script:gpRows[$gk].NewBox.Text.Trim()
        if ($newName -eq "" -or $gk -eq "__ROOT__" -or -not [System.IO.Directory]::Exists($gk)) { continue }
        $parent  = [System.IO.Path]::GetDirectoryName($gk)
        $newPath = [System.IO.Path]::Combine($parent, $newName)
        if ($gk -ne $newPath) { $ops.Add([PSCustomObject]@{ Type = "gp"; OldPath = $gk; NewPath = $newPath }) }
    }

    if ($ops.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Nothing to rename - all names are already up to date.", "No Changes", "OK", "Information") | Out-Null
        return
    }

    $summary = "About to perform $($ops.Count) rename operation(s).`n`n"
    $maxShow = [Math]::Min($ops.Count, 12)
    for ($i = 0; $i -lt $maxShow; $i++) {
        $o = $ops[$i]
        $summary += [System.IO.Path]::GetFileName($o.OldPath) + "`n  -> " + [System.IO.Path]::GetFileName($o.NewPath) + "`n"
    }
    if ($ops.Count -gt 12) { $summary += "... and $($ops.Count - 12) more`n" }
    $summary += "`nProceed?"
    $res = [System.Windows.Forms.MessageBox]::Show($summary, "Confirm Rename", "YesNo", "Question")
    if ($res -ne "Yes") { return }

    foreach ($typeOrder in @("file", "parent", "gp")) {
        foreach ($op in $ops) {
            if ($op.Type -ne $typeOrder) { continue }
            $maxAttempts = if ($op.Type -eq "gp") { 3 } else { 1 }
            $success = $false
            $lastErr = ""
            for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                try {
                    if ($op.Type -eq "file") {
                        if ([System.IO.File]::Exists($op.OldPath)) {
                            [System.IO.File]::Move($op.OldPath, $op.NewPath)
                            $renamed++
                        }
                        $success = $true
                    } else {
                        if ([System.IO.Directory]::Exists($op.OldPath)) {
                            [System.IO.Directory]::Move($op.OldPath, $op.NewPath)
                            $renamed++
                        }
                        $success = $true
                    }
                    break
                } catch {
                    $lastErr = $_.Exception.Message
                    if ($attempt -lt $maxAttempts) { Start-Sleep -Milliseconds 400 }
                }
            }
            if (-not $success -and $lastErr -ne "") {
                $label = $op.Type.ToUpper()
                $oldName = [System.IO.Path]::GetFileName($op.OldPath)
                $newName = [System.IO.Path]::GetFileName($op.NewPath)
                $errors.Add("[$label]  $oldName  ->  $newName`n    $lastErr")
            }
        }
    }

    if ($errors.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Successfully renamed $renamed item(s).`n`nUse Clear All and re-scan to reload the updated names.", "Done", "OK", "Information") | Out-Null
    } else {
        $msg = "Renamed $renamed item(s) with $($errors.Count) error(s):`n`n" + ($errors -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, "Partial Success", "OK", "Warning") | Out-Null
    }
}

# --- Build panel -------------------------------------------------------------
function script:RebuildPanel {
    $savedAdj   = @{}; $savedChar = @{}; $savedGpNew = @{}
    foreach ($fp in $script:fileRows.Keys) {
        $r = $script:fileRows[$fp]
        $savedAdj[$fp]  = $r.Adj.Text
        $savedChar[$fp] = $r.Char.Text
    }
    foreach ($gk in $script:gpRows.Keys) { $savedGpNew[$gk] = $script:gpRows[$gk].NewBox.Text }

    $script:fileRows   = @{}
    $script:gpRows     = @{}
    $script:parentRows = @{}

    $scroll.SuspendLayout()
    $scroll.Controls.Clear()

    if ($script:db.Count -eq 0) {
        $lbl           = New-Object System.Windows.Forms.Label
        $lbl.Text      = "No results yet - browse or drop a folder above to begin."
        $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
        $lbl.ForeColor = $cMuted
        $lbl.AutoSize  = $true
        $lbl.Location  = New-Object System.Drawing.Point(20, 20)
        $scroll.Controls.Add($lbl)
        $scroll.ResumeLayout()
        script:UpdateRenameButton
        return
    }

    $innerW = [Math]::Max(600, $scroll.ClientSize.Width - 8)
    $y = 6

    foreach ($gpKey in $script:db.Keys) {
        $gd = $script:db[$gpKey]
        $totalFiles = 0
        foreach ($pk in $gd.Parents.Keys) { $totalFiles += $gd.Parents[$pk].Files.Count }
        $badge = if ($totalFiles -eq 1) { "1 file" } else { "$totalFiles files" }

        $gpRow = @{
            NewBox      = $null
            Panel       = $null
            ThemeLabels = [System.Collections.Generic.List[System.Windows.Forms.Label]]::new()
            ParentKeys  = [System.Collections.Generic.List[string]]::new()
            FilePaths   = [System.Collections.Generic.List[string]]::new()
            OldName     = $gd.Name
        }
        $script:gpRows[$gpKey] = $gpRow

        # -- Grandparent header (56px tall) -----------------------------------
        $gpPanel           = New-Object System.Windows.Forms.Panel
        $gpPanel.BackColor = $cBG3
        $gpPanel.Size      = New-Object System.Drawing.Size($innerW, 56)
        $gpPanel.Location  = New-Object System.Drawing.Point(4, $y)

        $gpNameLbl           = New-Object System.Windows.Forms.Label
        $gpNameLbl.Text      = "  $($gd.Name)"
        $gpNameLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
        $gpNameLbl.ForeColor = $cAmber
        $gpNameLbl.AutoSize  = $true
        $gpNameLbl.Location  = New-Object System.Drawing.Point(4, 6)
        $gpPanel.Controls.Add($gpNameLbl)

        $gpBadge           = New-Object System.Windows.Forms.Label
        $gpBadge.Text      = "[$badge]"
        $gpBadge.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
        $gpBadge.ForeColor = $cMuted
        $gpBadge.AutoSize  = $true
        $gpBadge.Location  = New-Object System.Drawing.Point(($gpNameLbl.PreferredWidth + 10), 8)
        $gpPanel.Controls.Add($gpBadge)

        # OLD/NEW row
        $gpOldTag           = New-Object System.Windows.Forms.Label
        $gpOldTag.Text      = "OLD:"
        $gpOldTag.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        $gpOldTag.ForeColor = $cOldTag; $gpOldTag.AutoSize = $true
        $gpOldTag.Location  = New-Object System.Drawing.Point(8, 36)
        $gpPanel.Controls.Add($gpOldTag)

        $gpOldVal           = New-Object System.Windows.Forms.Label
        $gpOldVal.Text      = $gd.Name
        $gpOldVal.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
        $gpOldVal.ForeColor = $cOldText; $gpOldVal.AutoSize = $true
        $gpOldVal.Location  = New-Object System.Drawing.Point(38, 36)
        $gpPanel.Controls.Add($gpOldVal)

        $gpNewTag           = New-Object System.Windows.Forms.Label
        $gpNewTag.Text      = "NEW:"
        $gpNewTag.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        $gpNewTag.ForeColor = $cNewTag; $gpNewTag.AutoSize = $true
        $gpNewTag.Location  = New-Object System.Drawing.Point(250, 36)
        $gpPanel.Controls.Add($gpNewTag)

        $gpNewBox                  = New-Object System.Windows.Forms.TextBox
        $gpNewBox.Text             = if ($savedGpNew.ContainsKey($gpKey)) { $savedGpNew[$gpKey] } else { $gd.Name }
        $gpNewBox.Font             = New-Object System.Drawing.Font("Segoe UI", 8)
        $gpNewBox.BackColor        = $cInput; $gpNewBox.ForeColor = $cNewText
        $gpNewBox.BorderStyle      = "FixedSingle"
        $gpNewBox.Size             = New-Object System.Drawing.Size(($innerW - 290), 20)
        $gpNewBox.Location         = New-Object System.Drawing.Point(282, 33)
        $gpNewBox.Tag              = $gpKey
        $gpPanel.Controls.Add($gpNewBox)
        $gpRow.NewBox  = $gpNewBox
        $gpRow.Panel  = $gpPanel

        $gpNewBox.Add_TextChanged({
            param($s, $e)
            $cur = $s.SelectionStart; $old = $s.Text; $new2 = $old -replace '[^a-zA-Z0-9]', ''
            if ($old -ne $new2) { $s.Text = $new2; $s.SelectionStart = [Math]::Max(0,$cur - ($old.Length - $new2.Length)); $s.SelectionLength = 0 }
            script:UpdateFromGp $s.Tag
        })

        $scroll.Controls.Add($gpPanel)
        $y += 60

        foreach ($ppKey in $gd.Parents.Keys) {
            $pd = $gd.Parents[$ppKey]

            $pRow = @{ NewLabel = $null; Panel = $null; GpKey = $gpKey; FilePaths = [System.Collections.Generic.List[string]]::new(); OldName = $pd.Name }
            $script:parentRows[$ppKey] = $pRow
            $gpRow.ParentKeys.Add($ppKey)

            # -- Parent header (46px) -----------------------------------------
            $pPanel           = New-Object System.Windows.Forms.Panel
            $pPanel.BackColor = $cBG2
            $pPanel.Size      = New-Object System.Drawing.Size($innerW, 46)
            $pPanel.Location  = New-Object System.Drawing.Point(4, $y)

            $pNameLbl           = New-Object System.Windows.Forms.Label
            $pNameLbl.Text      = "    $($pd.Name)"
            $pNameLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $pNameLbl.ForeColor = $cBlue; $pNameLbl.AutoSize = $true
            $pNameLbl.Location  = New-Object System.Drawing.Point(4, 4)
            $pPanel.Controls.Add($pNameLbl)

            $pOldTag            = New-Object System.Windows.Forms.Label
            $pOldTag.Text       = "OLD:"; $pOldTag.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
            $pOldTag.ForeColor  = $cOldTag; $pOldTag.AutoSize = $true
            $pOldTag.Location   = New-Object System.Drawing.Point(16, 26)
            $pPanel.Controls.Add($pOldTag)

            $pOldVal            = New-Object System.Windows.Forms.Label
            $pOldVal.Text       = $pd.Name; $pOldVal.Font = New-Object System.Drawing.Font("Segoe UI", 8)
            $pOldVal.ForeColor  = $cOldText; $pOldVal.AutoSize = $true
            $pOldVal.Location   = New-Object System.Drawing.Point(46, 26)
            $pPanel.Controls.Add($pOldVal)

            $pNewLbl            = New-Object System.Windows.Forms.Label
            $pNewLbl.Text       = "NEW:  (enter fields to preview)"
            $pNewLbl.Font       = New-Object System.Drawing.Font("Segoe UI", 8)
            $pNewLbl.ForeColor  = $cNewText; $pNewLbl.AutoSize = $true
            $pNewLbl.Location   = New-Object System.Drawing.Point(300, 26)
            $pPanel.Controls.Add($pNewLbl)

            $pRow.NewLabel = $pNewLbl
            $pRow.Panel    = $pPanel

            $scroll.Controls.Add($pPanel)
            $y += 50

            foreach ($fp in $pd.Files) {
                $fn = [System.IO.Path]::GetFileName($fp)
                $pRow.FilePaths.Add($fp)
                $gpRow.FilePaths.Add($fp)

                # Smart-fill Character (and Adj) from filename
                $fillDefaults = script:SmartFill $fn $gd.Name
                $defaultChar  = $fillDefaults.Char
                $defaultAdj   = $fillDefaults.Adj

                # -- File row (96px) ------------------------------------------
                $rowPanel           = New-Object System.Windows.Forms.Panel
                $rowPanel.BackColor = $cBG4
                $rowPanel.Size      = New-Object System.Drawing.Size($innerW, 96)
                $rowPanel.Location  = New-Object System.Drawing.Point(4, $y)

                # OLD / NEW preview labels
                $fOldTag = New-Object System.Windows.Forms.Label
                $fOldTag.Text = "OLD:"; $fOldTag.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
                $fOldTag.ForeColor = $cOldTag; $fOldTag.AutoSize = $true
                $fOldTag.Location = New-Object System.Drawing.Point(12, 6)
                $rowPanel.Controls.Add($fOldTag)

                $fOldVal = New-Object System.Windows.Forms.Label
                $fOldVal.Text = $fn; $fOldVal.Font = New-Object System.Drawing.Font("Consolas", 8)
                $fOldVal.ForeColor = $cOldText; $fOldVal.AutoSize = $true
                $fOldVal.Location = New-Object System.Drawing.Point(42, 7)
                $rowPanel.Controls.Add($fOldVal)

                $fNewTag = New-Object System.Windows.Forms.Label
                $fNewTag.Text = "NEW:"; $fNewTag.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
                $fNewTag.ForeColor = $cNewTag; $fNewTag.AutoSize = $true
                $fNewTag.Location = New-Object System.Drawing.Point(12, 22)
                $rowPanel.Controls.Add($fNewTag)

                $fNewLbl = New-Object System.Windows.Forms.Label
                $fNewLbl.Text = "NEW:  (enter Character to preview)"
                $fNewLbl.Font = New-Object System.Drawing.Font("Consolas", 8)
                $fNewLbl.ForeColor = $cNewText; $fNewLbl.AutoSize = $true
                $fNewLbl.Location = New-Object System.Drawing.Point(42, 22)
                $rowPanel.Controls.Add($fNewLbl)

                # Compute initial layout
                $lay = script:InputLayout $innerW

                # Column header labels (y=44)
                $lChar = New-Object System.Windows.Forms.Label
                $lChar.Text = "Character  *required"; $lChar.Font = New-Object System.Drawing.Font("Segoe UI", 7)
                $lChar.ForeColor = $cLabel; $lChar.AutoSize = $true
                $lChar.Location = New-Object System.Drawing.Point($lay.xChar, 44)
                $rowPanel.Controls.Add($lChar)

                $lAdj = New-Object System.Windows.Forms.Label
                $lAdj.Text = "(Adj.)  - optional"; $lAdj.Font = New-Object System.Drawing.Font("Segoe UI", 7)
                $lAdj.ForeColor = $cLabel; $lAdj.AutoSize = $true
                $lAdj.Location = New-Object System.Drawing.Point($lay.xAdj, 44)
                $rowPanel.Controls.Add($lAdj)

                $lTheme = New-Object System.Windows.Forms.Label
                $lTheme.Text = "Theme"; $lTheme.Font = New-Object System.Drawing.Font("Segoe UI", 7)
                $lTheme.ForeColor = $cLabel; $lTheme.AutoSize = $true
                $lTheme.Location = New-Object System.Drawing.Point($lay.xTheme, 44)
                $rowPanel.Controls.Add($lTheme)

                # Input row (y=58)
                $tbChar = New-Object System.Windows.Forms.TextBox
                $tbChar.BackColor = $cInput; $tbChar.ForeColor = $cText
                $tbChar.BorderStyle = "FixedSingle"; $tbChar.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $tbChar.Size = New-Object System.Drawing.Size($lay.wChar, 24)
                $tbChar.Location = New-Object System.Drawing.Point($lay.xChar, 58)
                $tbChar.Tag = $fp
                $rowPanel.Controls.Add($tbChar)

                $s1 = New-Object System.Windows.Forms.Label
                $s1.Text = "_"; $s1.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
                $s1.ForeColor = $cMuted; $s1.AutoSize = $true
                $s1.Location = New-Object System.Drawing.Point($lay.xSep1, 59)
                $rowPanel.Controls.Add($s1)

                $tbAdj = New-Object System.Windows.Forms.TextBox
                $tbAdj.BackColor = $cInput; $tbAdj.ForeColor = $cText
                $tbAdj.BorderStyle = "FixedSingle"; $tbAdj.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $tbAdj.Size = New-Object System.Drawing.Size($lay.wAdj, 24)
                $tbAdj.Location = New-Object System.Drawing.Point($lay.xAdj, 58)
                $tbAdj.Tag = $fp
                $rowPanel.Controls.Add($tbAdj)

                $s2 = New-Object System.Windows.Forms.Label
                $s2.Text = "_"; $s2.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
                $s2.ForeColor = $cMuted; $s2.AutoSize = $true
                $s2.Location = New-Object System.Drawing.Point($lay.xSep2, 59)
                $rowPanel.Controls.Add($s2)

                # Theme: Label styled as non-interactive box
                $tbTheme = New-Object System.Windows.Forms.Label
                $tbTheme.Text = "  " + (script:NoSpaces $gpRow.NewBox.Text.Trim())
                $tbTheme.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $tbTheme.BackColor = $cThemeBG; $tbTheme.ForeColor = $cThemeFG
                $tbTheme.BorderStyle = "FixedSingle"
                $tbTheme.Size = New-Object System.Drawing.Size($lay.wTheme, 24)
                $tbTheme.Location = New-Object System.Drawing.Point($lay.xTheme, 58)
                $tbTheme.TextAlign = "MiddleLeft"
                $rowPanel.Controls.Add($tbTheme)
                $gpRow.ThemeLabels.Add($tbTheme)

                # Register
                $script:fileRows[$fp] = @{
                    Adj        = $tbAdj;    Char       = $tbChar
                    ThemeLabel = $tbTheme;  NewLabel   = $fNewLbl
                    GpKey      = $gpKey;    ParentKey  = $ppKey
                    Panel      = $rowPanel
                    Sep1       = $s1;       Sep2       = $s2
                    LChar      = $lChar;    LAdj       = $lAdj;   LTheme = $lTheme
                }

                # Restore or default
                $tbAdj.Text  = if ($savedAdj.ContainsKey($fp))  { $savedAdj[$fp]  } else { $defaultAdj  }
                $tbChar.Text = if ($savedChar.ContainsKey($fp)) { $savedChar[$fp] } else { $defaultChar }

                # TextChanged -- strip any non-alphanumeric characters and propagate
                $tbChar.Add_TextChanged({
                    param($s, $e)
                    $cur = $s.SelectionStart; $old = $s.Text; $new2 = $old -replace '[^a-zA-Z0-9]', ''
                    if ($old -ne $new2) { $s.Text=$new2; $s.SelectionStart=[Math]::Max(0,$cur - ($old.Length - $new2.Length)); $s.SelectionLength=0 }
                    script:UpdateFromFile $s.Tag
                })
                $tbAdj.Add_TextChanged({
                    param($s, $e)
                    $cur = $s.SelectionStart; $old = $s.Text; $new2 = $old -replace '[^a-zA-Z0-9]', ''
                    if ($old -ne $new2) { $s.Text=$new2; $s.SelectionStart=[Math]::Max(0,$cur - ($old.Length - $new2.Length)); $s.SelectionLength=0 }
                    script:UpdateFromFile $s.Tag
                })

                $scroll.Controls.Add($rowPanel)
                $y += 100
            }
        }
        $y += 8
    }

    $scroll.AutoScrollMinSize = New-Object System.Drawing.Size(0, ($y + 10))
    $scroll.ResumeLayout()

    # Fire initial previews
    foreach ($fp in $script:fileRows.Keys)   { script:UpdateFromFile $fp }
    foreach ($gk in $script:gpRows.Keys) {
        foreach ($ppKey in $script:gpRows[$gk].ParentKeys) {
            $pRow = $script:parentRows[$ppKey]
            if ($pRow) {
                $v = script:ComputeParentNew $ppKey
                $pRow.NewLabel.Text = if ($v) { "NEW:  $v" } else { "NEW:  (enter fields to preview)" }
            }
        }
    }
    script:UpdateRenameButton
}

function script:RefreshStatus {
    if ($script:db.Count -eq 0) {
        $lblStatus.Text = "No results.  Browse a folder or drag files onto the window to start."
        $lblStatus.ForeColor = $cMuted; return
    }
    $fc = 0
    foreach ($gk in $script:db.Keys) { foreach ($pk in $script:db[$gk].Parents.Keys) { $fc += $script:db[$gk].Parents[$pk].Files.Count } }
    $gc = $script:db.Count
    $lblStatus.Text = "Found  $fc $(if($fc-eq 1){'file'}else{'files'})  across  $gc grandparent $(if($gc-eq 1){'group'}else{'groups'})."
    $lblStatus.ForeColor = $cGreen
}

# --- Fonts -------------------------------------------------------------------
$fntTitle  = New-Object System.Drawing.Font("Segoe UI Semibold", 17)
$fntSub    = New-Object System.Drawing.Font("Segoe UI",  9)
$fntBtn    = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$fntStatus = New-Object System.Drawing.Font("Segoe UI",  9)
$fntDrop   = New-Object System.Drawing.Font("Segoe UI", 11)

# --- Form --------------------------------------------------------------------
$form                = New-Object System.Windows.Forms.Form
$form.Text           = "3MF File Finder"
$form.ClientSize     = New-Object System.Drawing.Size(840, 720)
$form.MinimumSize    = New-Object System.Drawing.Size(700, 560)
$form.BackColor      = $cBG; $form.ForeColor = $cText
$form.Font           = $fntSub
$form.StartPosition  = "CenterScreen"
$form.AllowDrop      = $true

$lblTitle            = New-Object System.Windows.Forms.Label
$lblTitle.Text       = "3MF File Finder"; $lblTitle.Font = $fntTitle
$lblTitle.ForeColor  = $cText; $lblTitle.AutoSize = $true
$lblTitle.Location   = New-Object System.Drawing.Point(24, 20)
$form.Controls.Add($lblTitle)

$lblSub              = New-Object System.Windows.Forms.Label
$lblSub.Text         = "Recursively finds *Full.3mf* files and groups them by grandparent folder"
$lblSub.Font         = $fntSub; $lblSub.ForeColor = $cMuted; $lblSub.AutoSize = $true
$lblSub.Location     = New-Object System.Drawing.Point(26, 52)
$form.Controls.Add($lblSub)

$sep                 = New-Object System.Windows.Forms.Panel
$sep.Size            = New-Object System.Drawing.Size(792, 1)
$sep.Location        = New-Object System.Drawing.Point(24, 75)
$sep.BackColor       = $cBorder
$form.Controls.Add($sep)

# --- Drop Zone ---------------------------------------------------------------
$dropPanel              = New-Object System.Windows.Forms.Panel
$dropPanel.Size         = New-Object System.Drawing.Size(792, 128)
$dropPanel.Location     = New-Object System.Drawing.Point(24, 84)
$dropPanel.BackColor    = $cBG2
$dropPanel.AllowDrop    = $true

$dropPanel.Add_Paint({
    param($s, $e)
    $e.Graphics.Clear($s.BackColor)   # clear first to wipe old drawing artifacts
    $pen           = New-Object System.Drawing.Pen($cBorder, 1.5)
    $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
    $e.Graphics.DrawRectangle($pen, 1, 1, $s.Width - 3, $s.Height - 3)
    $pen.Dispose()
})
$form.Controls.Add($dropPanel)

$lblDrop              = New-Object System.Windows.Forms.Label
$lblDrop.Text         = "Drop folders or files here"
$lblDrop.Font         = $fntDrop; $lblDrop.ForeColor = $cMuted
$lblDrop.AutoSize     = $false
$lblDrop.Size         = New-Object System.Drawing.Size(792, 34)
$lblDrop.TextAlign    = "MiddleCenter"
$lblDrop.Location     = New-Object System.Drawing.Point(0, 8)
$lblDrop.AllowDrop    = $true
$dropPanel.Controls.Add($lblDrop)

function New-FlatButton($text, $bg, $fg, $borderColor, $w) {
    if (-not $w) { $w = 162 }
    $b                                   = New-Object System.Windows.Forms.Button
    $b.Text                              = $text
    $b.Size                              = New-Object System.Drawing.Size($w, 34)
    $b.FlatStyle                         = "Flat"
    $b.BackColor                         = $bg; $b.ForeColor = $fg; $b.Font = $fntBtn
    $b.Cursor                            = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize         = if ($borderColor) { 1 } else { 0 }
    if ($borderColor) { $b.FlatAppearance.BorderColor = $borderColor }
    $b.FlatAppearance.MouseOverBackColor = $b.BackColor
    $b.FlatAppearance.MouseDownBackColor = $b.BackColor
    return $b
}

$btnBrowseFolder = New-FlatButton "  Browse Folder" $cAmber $cBG  $null    162
$btnBrowseFile   = New-FlatButton "  Browse File"   $cBG3  $cText $cBorder 162
$btnClear        = New-FlatButton "  Clear All"     $cBG3  $cRed  $cBorder 120
$btnRename       = New-FlatButton "  RENAME"        $cBtnOff $cBtnOffT $cBorder 130

# Helper: evenly space 4 buttons in the drop panel
function script:PositionButtons {
    $w      = $dropPanel.Width
    $total  = 162 + 162 + 120 + 130   # button widths
    $gaps   = 3                         # gaps between 4 buttons
    $spacing = [int](($w - $total) / ($gaps + 1))
    $x = $spacing
    $btnBrowseFolder.Left = $x; $x += 162 + $spacing
    $btnBrowseFile.Left   = $x; $x += 162 + $spacing
    $btnClear.Left        = $x; $x += 120 + $spacing
    $btnRename.Left       = $x
}

$btnBrowseFolder.Top = 84; $btnBrowseFile.Top = 84; $btnClear.Top = 84; $btnRename.Top = 84
foreach ($b in @($btnBrowseFolder,$btnBrowseFile,$btnClear,$btnRename)) { $b.Top = 84; $dropPanel.Controls.Add($b) }
script:PositionButtons

$btnRename.Enabled = $false
script:StyleRenameBtn $false

$lblStatus            = New-Object System.Windows.Forms.Label
$lblStatus.Text       = "No results.  Browse a folder or drag files onto the window to start."
$lblStatus.Font       = $fntStatus; $lblStatus.ForeColor = $cMuted; $lblStatus.AutoSize = $true
$lblStatus.Location   = New-Object System.Drawing.Point(26, 222)
$form.Controls.Add($lblStatus)

$scroll               = New-Object System.Windows.Forms.Panel
$scroll.AutoScroll    = $true; $scroll.BackColor = $cBG; $scroll.BorderStyle = "None"
$scroll.Location      = New-Object System.Drawing.Point(24, 246)
$scroll.Size          = New-Object System.Drawing.Size(792, 458)
$scroll.Anchor        = [System.Windows.Forms.AnchorStyles]"Top, Bottom, Left, Right"
$form.Controls.Add($scroll)

# --- Responsive resize function (no closures -- direct property sets) --------
function script:ApplyResize {
    $innerW = [Math]::Max(600, $scroll.ClientSize.Width - 8)

    foreach ($gk in $script:gpRows.Keys) {
        $gr = $script:gpRows[$gk]
        if ($gr.Panel)  { $gr.Panel.Width  = $innerW }
        if ($gr.NewBox) { $gr.NewBox.Width = $innerW - 290 }
    }
    foreach ($ppKey in $script:parentRows.Keys) {
        $pr = $script:parentRows[$ppKey]
        if ($pr.Panel) { $pr.Panel.Width = $innerW }
    }
    $lay = script:InputLayout $innerW
    foreach ($fp in $script:fileRows.Keys) {
        $fr = $script:fileRows[$fp]
        if (-not $fr.Panel) { continue }
        $fr.Panel.Width = $innerW
        $fr.Char.Location  = New-Object System.Drawing.Point($lay.xChar,  58)
        $fr.Char.Width     = $lay.wChar
        $fr.Sep1.Location  = New-Object System.Drawing.Point($lay.xSep1, 59)
        $fr.Adj.Location   = New-Object System.Drawing.Point($lay.xAdj,  58)
        $fr.Adj.Width      = $lay.wAdj
        $fr.Sep2.Location  = New-Object System.Drawing.Point($lay.xSep2, 59)
        $fr.ThemeLabel.Location = New-Object System.Drawing.Point($lay.xTheme, 58)
        $fr.ThemeLabel.Width    = $lay.wTheme
        $fr.LChar.Location = New-Object System.Drawing.Point($lay.xChar,  44)
        $fr.LAdj.Location  = New-Object System.Drawing.Point($lay.xAdj,  44)
        $fr.LTheme.Location = New-Object System.Drawing.Point($lay.xTheme, 44)
    }
}

# --- Resize ------------------------------------------------------------------
$form.Add_Resize({
    $w = $form.ClientSize.Width - 48
    $sep.Width       = $w
    $dropPanel.Width = $w
    $lblDrop.Width   = $w
    script:PositionButtons
    $dropPanel.Invalidate()

    $scroll.Width  = $w
    $scroll.Height = $form.ClientSize.Height - 268

    script:ApplyResize
})

# --- Drag & Drop -------------------------------------------------------------
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
    $ctrl.Add_DragEnter($onDragEnter)
    $ctrl.Add_DragLeave($onDragLeave)
    $ctrl.Add_DragDrop($onDragDrop)
}

# --- Buttons -----------------------------------------------------------------
$btnBrowseFolder.Add_Click({
    $path = [ModernFolderPicker]::Pick($form.Handle, "Select a folder to search for Full.3mf files")
    if ($path) { script:ScanPath $path; script:RebuildPanel; script:RefreshStatus }
})
$btnBrowseFile.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select 3MF file(s)"; $ofd.Filter = "3MF Files (*.3mf)|*.3mf|All Files (*.*)|*.*"
    $ofd.Multiselect = $true; $ofd.CheckPathExists = $true
    if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($f in $ofd.FileNames) { script:ScanPath $f }
        script:RebuildPanel; script:RefreshStatus
    }
})
$btnClear.Add_Click({
    $script:db = [System.Collections.Specialized.OrderedDictionary]::new()
    $script:fileRows = @{}; $script:gpRows = @{}; $script:parentRows = @{}
    script:RebuildPanel; script:RefreshStatus
})
$btnRename.Add_Click({ script:DoRename })

# --- Initial render ----------------------------------------------------------
script:RebuildPanel
script:RefreshStatus

[System.Windows.Forms.Application]::Run($form)