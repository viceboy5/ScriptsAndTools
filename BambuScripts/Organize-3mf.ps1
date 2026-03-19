#Requires -Version 3.0
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Hide the PowerShell console window --------------------------------------
try {
    Add-Type -Name ConsoleHider -Namespace Win32 -MemberDefinition @'
        [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]   public static extern bool   ShowWindow(IntPtr hWnd, int nCmdShow);
'@
    [Win32.ConsoleHider]::ShowWindow([Win32.ConsoleHider]::GetConsoleWindow(), 0) | Out-Null
} catch { }

# --- Modern Windows Explorer-style Folder Picker (IFileOpenDialog COM) -------
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

# --- Color palette -----------------------------------------------------------
function clr($hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
$cBG      = clr "#111214"
$cBG2     = clr "#16171B"
$cBG3     = clr "#1C1D23"
$cBorder  = clr "#2A2C35"
$cText    = clr "#DDE0E8"
$cMuted   = clr "#5A5D6B"
$cAmber   = clr "#E8A135"
$cBlue    = clr "#6DAEE0"
$cRed     = clr "#D95F5F"
$cGreen   = clr "#4CAF72"
$cDropHL  = clr "#1E2010"

# --- Data store --------------------------------------------------------------
$script:db = [System.Collections.Specialized.OrderedDictionary]::new()

function script:IsMatch([string]$name) { $name -imatch 'full\.3mf' }

function script:AddFile([string]$filePath) {
    try {
        $fi  = [System.IO.FileInfo]::new($filePath)
        $par = $fi.Directory
        $gp  = $par.Parent

        $gpKey  = if ($gp)  { $gp.FullName } else { "__ROOT__" }
        $gpName = if ($gp)  { $gp.Name     } else { "(root)" }

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

        if (-not ($pd.Files -contains $fi.FullName)) {
            $pd.Files.Add($fi.FullName)
        }
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

function script:RebuildTree {
    $tree.BeginUpdate()
    $tree.Nodes.Clear()

    if ($script:db.Count -eq 0) {
        $n = New-Object System.Windows.Forms.TreeNode "  No results yet - browse or drop a folder above to begin."
        $n.ForeColor = $cMuted
        $tree.Nodes.Add($n) | Out-Null
        $tree.EndUpdate()
        return
    }

    foreach ($gpKey in $script:db.Keys) {
        $gd = $script:db[$gpKey]

        $totalFiles = 0
        foreach ($pk in $gd.Parents.Keys) { $totalFiles += $gd.Parents[$pk].Files.Count }
        $label = if ($totalFiles -eq 1) { "1 file" } else { "$totalFiles files" }

        $gpNode           = New-Object System.Windows.Forms.TreeNode
        $gpNode.Text      = "  $($gd.Name)    [$label]"
        $gpNode.ForeColor = $cAmber
        $gpNode.NodeFont  = New-Object System.Drawing.Font("Segoe UI Semibold", 10)

        foreach ($ppKey in $gd.Parents.Keys) {
            $pd = $gd.Parents[$ppKey]

            $pNode           = New-Object System.Windows.Forms.TreeNode
            $pNode.Text      = "  $($pd.Name)"
            $pNode.ForeColor = $cBlue
            $pNode.NodeFont  = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

            foreach ($fp in $pd.Files) {
                $fn              = [System.IO.Path]::GetFileName($fp)
                $fNode           = New-Object System.Windows.Forms.TreeNode
                $fNode.Text      = "      $fn"
                $fNode.ForeColor = $cText
                $fNode.NodeFont  = New-Object System.Drawing.Font("Consolas", 8.5)
                $pNode.Nodes.Add($fNode) | Out-Null
            }

            $pNode.Expand()
            $gpNode.Nodes.Add($pNode) | Out-Null
        }

        $gpNode.Expand()
        $tree.Nodes.Add($gpNode) | Out-Null
    }

    $tree.EndUpdate()
}

function script:RefreshStatus {
    if ($script:db.Count -eq 0) {
        $lblStatus.Text      = "No results.  Browse a folder or drag files onto the window to start."
        $lblStatus.ForeColor = $cMuted
        return
    }
    $fc = 0
    foreach ($gk in $script:db.Keys) {
        foreach ($pk in $script:db[$gk].Parents.Keys) {
            $fc += $script:db[$gk].Parents[$pk].Files.Count
        }
    }
    $gc = $script:db.Count
    $fs = if ($fc -eq 1) { "file" } else { "files" }
    $gs = if ($gc -eq 1) { "group" } else { "groups" }
    $lblStatus.Text      = "Found  $fc $fs  across  $gc grandparent $gs."
    $lblStatus.ForeColor = $cGreen
}

# --- Fonts -------------------------------------------------------------------
$fntTitle  = New-Object System.Drawing.Font("Segoe UI Semibold", 17)
$fntSub    = New-Object System.Drawing.Font("Segoe UI",           9)
$fntBtn    = New-Object System.Drawing.Font("Segoe UI Semibold",  9)
$fntStatus = New-Object System.Drawing.Font("Segoe UI",           9)
$fntDrop   = New-Object System.Drawing.Font("Segoe UI",          11)
$fntTree   = New-Object System.Drawing.Font("Segoe UI",           9)

# --- Build Form --------------------------------------------------------------
$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "3MF File Finder"
$form.ClientSize       = New-Object System.Drawing.Size(780, 680)
$form.MinimumSize      = New-Object System.Drawing.Size(620, 520)
$form.BackColor        = $cBG
$form.ForeColor        = $cText
$form.Font             = $fntSub
$form.StartPosition    = "CenterScreen"
$form.AllowDrop        = $true

# Title
$lblTitle              = New-Object System.Windows.Forms.Label
$lblTitle.Text         = "3MF File Finder"
$lblTitle.Font         = $fntTitle
$lblTitle.ForeColor    = $cText
$lblTitle.AutoSize     = $true
$lblTitle.Location     = New-Object System.Drawing.Point(24, 20)
$form.Controls.Add($lblTitle)

$lblSub                = New-Object System.Windows.Forms.Label
$lblSub.Text           = "Recursively finds *Full.3mf* files and groups them by grandparent folder"
$lblSub.Font           = $fntSub
$lblSub.ForeColor      = $cMuted
$lblSub.AutoSize       = $true
$lblSub.Location       = New-Object System.Drawing.Point(26, 52)
$form.Controls.Add($lblSub)

# Separator
$sep                   = New-Object System.Windows.Forms.Panel
$sep.Size              = New-Object System.Drawing.Size(732, 1)
$sep.Location          = New-Object System.Drawing.Point(24, 75)
$sep.BackColor         = $cBorder
$form.Controls.Add($sep)

# --- Drop Zone ---------------------------------------------------------------
$dropPanel             = New-Object System.Windows.Forms.Panel
$dropPanel.Size        = New-Object System.Drawing.Size(732, 118)
$dropPanel.Location    = New-Object System.Drawing.Point(24, 84)
$dropPanel.BackColor   = $cBG2
$dropPanel.AllowDrop   = $true

$dropPanel.Add_Paint({
    param($s, $e)
    $pen           = New-Object System.Drawing.Pen($cBorder, 1.5)
    $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
    $e.Graphics.DrawRectangle($pen, 1, 1, $s.Width - 3, $s.Height - 3)
    $pen.Dispose()
})
$form.Controls.Add($dropPanel)

$lblDrop               = New-Object System.Windows.Forms.Label
$lblDrop.Text          = "Drop folders or files here"
$lblDrop.Font          = $fntDrop
$lblDrop.ForeColor     = $cMuted
$lblDrop.AutoSize      = $false
$lblDrop.Size          = New-Object System.Drawing.Size(732, 38)
$lblDrop.TextAlign     = "MiddleCenter"
$lblDrop.Location      = New-Object System.Drawing.Point(0, 12)
$lblDrop.AllowDrop     = $true
$dropPanel.Controls.Add($lblDrop)

# Button helper
function New-FlatButton($text, $bg, $fg, $borderColor) {
    $b                                   = New-Object System.Windows.Forms.Button
    $b.Text                              = $text
    $b.Size                              = New-Object System.Drawing.Size(158, 34)
    $b.FlatStyle                         = "Flat"
    $b.BackColor                         = $bg
    $b.ForeColor                         = $fg
    $b.Font                              = $fntBtn
    $b.Cursor                            = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize         = if ($borderColor) { 1 } else { 0 }
    if ($borderColor) { $b.FlatAppearance.BorderColor = $borderColor }
    $b.FlatAppearance.MouseOverBackColor = $b.BackColor
    $b.FlatAppearance.MouseDownBackColor = $b.BackColor
    return $b
}

$btnBrowseFolder = New-FlatButton "  Browse Folder" $cAmber $cBG $null
$btnBrowseFile   = New-FlatButton "  Browse File"   $cBG3 $cText $cBorder
$btnClear        = New-FlatButton "  Clear All"     $cBG3 $cRed $cBorder

$btnBrowseFolder.Location = New-Object System.Drawing.Point(120, 72)
$btnBrowseFile.Location   = New-Object System.Drawing.Point(292, 72)
$btnClear.Location        = New-Object System.Drawing.Point(464, 72)
$dropPanel.Controls.Add($btnBrowseFolder)
$dropPanel.Controls.Add($btnBrowseFile)
$dropPanel.Controls.Add($btnClear)

# Status bar
$lblStatus             = New-Object System.Windows.Forms.Label
$lblStatus.Text        = "No results.  Browse a folder or drag files onto the window to start."
$lblStatus.Font        = $fntStatus
$lblStatus.ForeColor   = $cMuted
$lblStatus.AutoSize    = $true
$lblStatus.Location    = New-Object System.Drawing.Point(26, 212)
$form.Controls.Add($lblStatus)

# Results tree
$tree                  = New-Object System.Windows.Forms.TreeView
$tree.BackColor        = $cBG2
$tree.ForeColor        = $cText
$tree.BorderStyle      = "None"
$tree.Font             = $fntTree
$tree.Location         = New-Object System.Drawing.Point(24, 236)
$tree.Size             = New-Object System.Drawing.Size(732, 428)
$tree.Anchor           = [System.Windows.Forms.AnchorStyles]"Top, Bottom, Left, Right"
$tree.ShowLines        = $false
$tree.ShowRootLines    = $false
$tree.ShowPlusMinus    = $true
$tree.ItemHeight       = 28
$tree.Indent           = 18
$form.Controls.Add($tree)

# --- Resize handler ----------------------------------------------------------
$form.Add_Resize({
    $w = $form.ClientSize.Width - 48
    $sep.Width             = $w
    $dropPanel.Width       = $w
    $lblDrop.Width         = $w
    $btnBrowseFolder.Left  = [int](($w - 514) / 2)
    $btnBrowseFile.Left    = $btnBrowseFolder.Left + 172
    $btnClear.Left         = $btnBrowseFile.Left + 172
    $tree.Width            = $w
    $tree.Height           = $form.ClientSize.Height - 258
})

# --- Drag and Drop handlers --------------------------------------------------
$onDragEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect            = [System.Windows.Forms.DragDropEffects]::Copy
        $dropPanel.BackColor = $cDropHL
        $dropPanel.Refresh()
    } else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
}

$onDragLeave = {
    $dropPanel.BackColor = $cBG2
    $dropPanel.Refresh()
}

$onDragDrop = {
    param($s, $e)
    $dropPanel.BackColor = $cBG2
    $dropPanel.Refresh()
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        foreach ($p in $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) {
            script:ScanPath $p
        }
        script:RebuildTree
        script:RefreshStatus
    }
}

foreach ($ctrl in @($form, $dropPanel, $lblDrop)) {
    $ctrl.AllowDrop = $true
    $ctrl.Add_DragEnter($onDragEnter)
    $ctrl.Add_DragLeave($onDragLeave)
    $ctrl.Add_DragDrop($onDragDrop)
}

# --- Button handlers ---------------------------------------------------------
$btnBrowseFolder.Add_Click({
    $path = [ModernFolderPicker]::Pick($form.Handle, "Select a folder to search for Full.3mf files")
    if ($path) {
        script:ScanPath $path
        script:RebuildTree
        script:RefreshStatus
    }
})

$btnBrowseFile.Add_Click({
    $ofd                 = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title           = "Select 3MF file(s)"
    $ofd.Filter          = "3MF Files (*.3mf)|*.3mf|All Files (*.*)|*.*"
    $ofd.Multiselect     = $true
    $ofd.CheckPathExists = $true
    if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($f in $ofd.FileNames) { script:ScanPath $f }
        script:RebuildTree
        script:RefreshStatus
    }
})

$btnClear.Add_Click({
    $script:db = [System.Collections.Specialized.OrderedDictionary]::new()
    script:RebuildTree
    script:RefreshStatus
})

# --- Initial render ----------------------------------------------------------
script:RebuildTree
script:RefreshStatus

[System.Windows.Forms.Application]::Run($form)