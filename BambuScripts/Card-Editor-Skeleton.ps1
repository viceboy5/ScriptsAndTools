Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# --- 1. Load the Official Library ---
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

# --- 2. Parse File Helpers & Folder Picker ---
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

function ParseFile([string]$filename) {
    $compoundExts = @('.gcode.3mf', '.gcode.stl', '.gcode.step', '.f3d.3mf')
    $ext = $null
    foreach ($ce in $compoundExts) {
        if ($filename.ToLower().EndsWith($ce.ToLower())) { $ext = $filename.Substring($filename.Length - $ce.Length); break }
    }
    if (-not $ext) { $ext = [System.IO.Path]::GetExtension($filename) }
    if (-not $ext) { $ext = '' }
    $stem  = $filename.Substring(0, $filename.Length - $ext.Length)
    $parts = [string[]]( ($stem -split '[\s._-]+') | Where-Object { $_ -ne '' } )
    if ($ext -ieq '.png') { return @{ Suffix = ''; Extension = $ext; Stem = $stem; Parts = $parts } }
    $suffix = if ($parts.Count -gt 0) { $parts[-1] } else { $stem }
    return @{ Suffix = $suffix; Extension = $ext; Stem = $stem; Parts = $parts }
}

function SmartFill([string]$anchorName, [string]$gpName) {
    $stem = (ParseFile $anchorName).Stem
    $escaped = [regex]::Escape($gpName)
    if ($gpName -ne '' -and $stem -imatch "^(.+)_${escaped}_Full$") {
        $prefix = $Matches[1]
        $segments = $prefix -split '_'
        if ($segments.Count -ge 2) {
            $adj = ($segments[1..($segments.Count-1)]) -join ''
            return @{ Char = $segments[0]; Adj = $adj }
        }
        return @{ Char = $prefix; Adj = '' }
    }
    $parts = (ParseFile $anchorName).Parts
    if ($parts.Count -ge 2) { return @{ Char = ($parts[0..($parts.Count - 2)] -join ''); Adj = '' } }
    return @{ Char = ($parts -join ''); Adj = '' }
}

function clr($hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { return [System.Drawing.Color]::Gray }
    if ($hex.Length -ge 9) { $hex = "#" + $hex.Substring(1,6) }
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

$cBG      = clr "#000000"; $cUI      = clr "#16171B"
$cText    = clr "#FFFFFF"; $cMuted   = clr "#A0A0A0"
$cAccent  = clr "#FFD700"; $cInput   = clr "#1E2028"
$cGrayTxt = clr "#B4B4B4"; $cGreen   = clr "#4CAF72"
$cRed     = clr "#D95F5F"; $cBorder  = clr "#2A2C35"

# --- 3. Build Queue from Drag & Drop or Folder Picker ---
$script:queue = [System.Collections.Generic.List[string]]::new()

if ($args.Count -eq 0) {
    $pickedFolder = [ModernFolderPicker]::Pick([IntPtr]::Zero, "Select a Master Folder to scan for Full.3mf files")
    if (-not $pickedFolder) { exit }

    $found = Get-ChildItem -Path $pickedFolder -Filter "*Full.3mf" -Recurse -File
    foreach ($f in $found) {
        if (-not $script:queue.Contains($f.DirectoryName)) { $script:queue.Add($f.DirectoryName) }
    }
} else {
    foreach ($p in $args) {
        if (Test-Path $p -PathType Container) {
            $found = Get-ChildItem -Path $p -Filter "*Full.3mf" -Recurse -File
            foreach ($f in $found) {
                if (-not $script:queue.Contains($f.DirectoryName)) { $script:queue.Add($f.DirectoryName) }
            }
        } elseif ($p -match '(?i)Full\.3mf$') {
            $fi = Get-Item $p
            if (-not $script:queue.Contains($fi.DirectoryName)) { $script:queue.Add($fi.DirectoryName) }
        }
    }
}

if ($script:queue.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No Full.3mf files found in the target directory.", "Empty Queue") | Out-Null
    exit
}

$script:qIndex = 0
$script:tempWork = ""

# --- 4. Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Batch Pre-Flight Editor"
$form.ClientSize = New-Object System.Drawing.Size(1250, 750)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 600)
$form.BackColor = $cUI; $form.StartPosition = "CenterScreen"

$card = New-Object System.Windows.Forms.Panel
$card.BackColor = $cBG; $card.BorderStyle = 'FixedSingle'
$form.Controls.Add($card)

$pnlRight = New-Object System.Windows.Forms.Panel
$pnlRight.BackColor = $cUI; $pnlRight.BorderStyle = 'None'
$form.Controls.Add($pnlRight)

$script:scaleElements = @()
function Add-ScaledElement($ctrl, $bx, $by, $bw, $bh, $bf) {
    $script:scaleElements += [PSCustomObject]@{ Ctrl = $ctrl; X = $bx; Y = $by; W = $bw; H = $bh; Font = $bf }
    $ctrl.AutoSize = $false
    $card.Controls.Add($ctrl)
}

# --- Global Sync Function ---
function script:Update-RenamePreview {
    if ($null -eq $script:rpTBChar) { return }

    $ch = $script:rpTBChar.Text -replace '[^a-zA-Z0-9]', ''
    $ad = $script:rpTBAdj.Text -replace '[^a-zA-Z0-9]', ''
    $th = $script:rpTBTheme.Text -replace '[^a-zA-Z0-9]', ''

    # Update Canvas Display Labels directly
    $script:rpLblChar.Text = $ch
    $script:rpLblAdj.Text  = if ($ad) { "($ad)" } else { "" }

    # Build green list preview
    foreach ($r in $script:fileRows) {
        $sf = $r.SuffixBox.Text -replace '[^a-zA-Z0-9]', ''
        $parts = [System.Collections.Generic.List[string]]::new()
        if ($ch) { $parts.Add($ch) }
        if ($ad) { $parts.Add($ad) }
        if ($th) { $parts.Add($th) }
        if ($sf) { $parts.Add($sf) }
        $r.NewLbl.Text = ($parts -join '_') + $r.Ext
    }
}

# --- 5. State Machine: Load Item ---
function Load-QueueItem {
    if ($script:qIndex -ge $script:queue.Count) {
        [System.Windows.Forms.MessageBox]::Show("All items processed successfully!", "Queue Complete") | Out-Null
        $form.Close()
        return
    }

    $card.Controls.Clear()
    $pnlRight.Controls.Clear()
    $script:scaleElements = @()
    if (Test-Path $script:tempWork) { Remove-Item $script:tempWork -Recurse -Force -ErrorAction SilentlyContinue }

    $folderPath = $script:queue[$script:qIndex]
    $anchorFile = Get-ChildItem -Path $folderPath -Filter "*Full.3mf" | Select-Object -First 1
    if (-not $anchorFile) { $script:qIndex++; Load-QueueItem; return }

    $form.Text = "Batch Pre-Flight Editor - Queue $($script:qIndex + 1) of $($script:queue.Count)"

    # Extract
    $script:tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -ItemType Directory -Path $script:tempWork | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($anchorFile.FullName, $script:tempWork)

    # Read Colors
    $activeSlots = @()
    $projPath   = Join-Path $script:tempWork "Metadata\project_settings.config"
    $modSetPath = Join-Path $script:tempWork "Metadata\model_settings.config"
    $modelFile  = (Get-ChildItem -Path $script:tempWork -Filter '3dmodel.model' -Recurse | Select-Object -First 1)

    $SlotMap = [ordered]@{}
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

    $UsedSlots = New-Object System.Collections.Generic.HashSet[string]
    $UsedSlots.Add("1") | Out-Null
    if (Test-Path $modSetPath) {
        try {
            [xml]$modXml = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
            foreach ($node in $modXml.SelectNodes('//metadata[contains(@key, "extruder")]')) {
                $val = $node.GetAttribute('value')
                if (-not [string]::IsNullOrWhiteSpace($val)) { $UsedSlots.Add($val) | Out-Null }
            }
        } catch {
            $modContent = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
            $extMatches = [regex]::Matches($modContent, '(?i)extruder[^>]*?"(\d+)"')
            foreach ($m in $extMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
        }
    }
    if ($modelFile -and (Test-Path $modelFile.FullName)) {
        try {
            $modelContent = [System.IO.File]::ReadAllText($modelFile.FullName)
            $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
            foreach ($m in $matMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
        } catch {}
    }

    foreach ($hex in $SlotMap.Keys) {
        $slotId = $SlotMap[$hex]
        if ($UsedSlots.Contains($slotId)) {
            $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
            $matchedName = if ($HexToName.Contains($checkHex)) { $HexToName[$checkHex] } else { "" }
            $activeSlots += [PSCustomObject]@{ OldHex = $checkHex; Name = $matchedName; Grams = 0 }
        }
    }
    if ($activeSlots.Count -gt 4) { $activeSlots = $activeSlots[0..3] }

    # --- Build Left Canvas ---
    $plateImgPath = Join-Path $script:tempWork "Metadata\plate_1.png"
    $pbModel = New-Object System.Windows.Forms.PictureBox
    $pbModel.SizeMode = 'Zoom'; $pbModel.BackColor = $cBG
    $pbModel.AllowDrop = $true
    if (Test-Path $plateImgPath) {
        $fs = New-Object System.IO.FileStream($plateImgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pbModel.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
    }
    Add-ScaledElement $pbModel 10 80 250 360 0

    $pbModel.Add_DragEnter({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($files[0] -match '(?i)\.(png|jpg|jpeg)$') { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
        }
    })
    $pbModel.Add_DragDrop({
        param($s, $e)
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        $dropped = $files[0]
        if ($dropped -match '(?i)\.(png|jpg|jpeg)$') {
            $dest = Join-Path $folderPath (Split-Path $dropped -Leaf)
            if ($dropped -ne $dest) { Copy-Item -Path $dropped -Destination $dest -Force }
            $fs = New-Object System.IO.FileStream($dest, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            if ($s.Image) { $s.Image.Dispose() }
            $s.Image = [System.Drawing.Image]::FromStream($fs)
            $fs.Close()
        }
    })

    # Read-Only Labels mapped to the canvas (Replaced TextBoxes)
    $lblAdjTitle = New-Object System.Windows.Forms.Label
    $lblAdjTitle.BackColor = $cBG; $lblAdjTitle.ForeColor = $cAccent
    $lblAdjTitle.TextAlign = 'TopRight'
    Add-ScaledElement $lblAdjTitle 320 15 180 40 18

    $lblCharTitle = New-Object System.Windows.Forms.Label
    $lblCharTitle.BackColor = $cBG; $lblCharTitle.ForeColor = $cText
    $lblCharTitle.TextAlign = 'TopRight'
    Add-ScaledElement $lblCharTitle 60 15 250 40 18

    $lblSkipTime = New-Object System.Windows.Forms.Label
    $lblSkipTime.Text = "Skip Time: 18 min"
    $lblSkipTime.BackColor = $cBG; $lblSkipTime.ForeColor = $cText
    $lblSkipTime.TextAlign = 'BottomLeft'
    Add-ScaledElement $lblSkipTime 10 472 250 30 14

    $startY = 80; $boxSize = 76; $slotSpacing = $boxSize + 10; $boxX = 512 - 10 - $boxSize
    $script:uiSlots = @()

    for ($i = 0; $i -lt $activeSlots.Count; $i++) {
        $slotData = $activeSlots[$i]
        $swatch = New-Object System.Windows.Forms.Panel
        $swatch.BorderStyle = 'FixedSingle'
        try { $swatch.BackColor = clr $slotData.OldHex } catch { $swatch.BackColor = clr "#333333" }
        $swatch.Tag = ($i + 1).ToString()
        $swatch.Add_Paint({
            param($s, $e)
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
            $e.Graphics.DrawString($s.Tag, $s.Font, $brush, [int]($s.Width*0.35), [int]($s.Height*0.25))
            $brush.Dispose()
        })
        Add-ScaledElement $swatch $boxX $startY $boxSize $boxSize 16

        $lblStatus = New-Object System.Windows.Forms.Label
        $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $lblStatus.TextAlign = 'BottomRight'; $lblStatus.BackColor = $cBG
        if ($slotData.Name) { $lblStatus.Text = "[MATCHED]"; $lblStatus.ForeColor = $cGreen }
        else { $lblStatus.Text = "[UNMATCHED]"; $lblStatus.ForeColor = $cRed }
        Add-ScaledElement $lblStatus ($boxX - 180) ($startY) 170 15 8

        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.BackColor = $cInput; $combo.ForeColor = $cText
        $combo.DropDownStyle = 'DropDown'
        $combo.AutoCompleteMode = 'SuggestAppend'
        $combo.AutoCompleteSource = 'ListItems'
        foreach ($k in $LibraryColors.Keys) { [void]$combo.Items.Add($k) }
        if ($slotData.Name) { $combo.Text = $slotData.Name } else { $combo.Text = "Select Color..." }
        $combo.Tag = @{ Swatch = $swatch; StatusLbl = $lblStatus; OrigName = $slotData.Name }
        $combo.add_TextChanged({
            param($s, $e)
            $data = $s.Tag
            if ($LibraryColors.Contains($s.Text)) { $data.Swatch.BackColor = clr $LibraryColors[$s.Text] }
            if ($s.Text -eq $data.OrigName) {
                if ($data.OrigName) { $data.StatusLbl.Text = "[MATCHED]"; $data.StatusLbl.ForeColor = clr "#4CAF72" }
                else { $data.StatusLbl.Text = "[UNMATCHED]"; $data.StatusLbl.ForeColor = clr "#D95F5F" }
            } else {
                $data.StatusLbl.Text = "[CHANGED]"; $data.StatusLbl.ForeColor = clr "#E8A135"
            }
            $s.SelectionStart = 0; $s.SelectionLength = 0
        }.GetNewClosure())
        Add-ScaledElement $combo ($boxX - 180) ($startY + 15) 170 25 10

        $lblGrams = New-Object System.Windows.Forms.Label
        $lblGrams.Text = "$($slotData.Grams) g"
        $lblGrams.ForeColor = $cGrayTxt; $lblGrams.TextAlign = 'TopRight'
        Add-ScaledElement $lblGrams ($boxX - 90) ($startY + 45) 80 25 10

        $script:uiSlots += [PSCustomObject]@{ OldHex = $slotData.OldHex; Combo = $combo }
        $startY += $slotSpacing
    }
    $pbModel.SendToBack()

    # --- Build Right Rename Panel ---
    $diParent = [System.IO.DirectoryInfo]::new($folderPath)
    $diGrand = $diParent.Parent
    $gpName = if ($diGrand) { $diGrand.Name } else { "(root)" }
    $fills = SmartFill $anchorFile.Name $gpName

    $y = 10
    $lblFolder = New-Object System.Windows.Forms.Label
    $lblFolder.Text = "Folder: $($diParent.Name)"
    $lblFolder.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblFolder.ForeColor = $cText; $lblFolder.AutoSize = $true
    $lblFolder.Location = New-Object System.Drawing.Point(10, $y)
    $pnlRight.Controls.Add($lblFolder)
    $y += 40

    # Grandparent Theme Block
    $gpBox = New-Object System.Windows.Forms.Panel
    $gpBox.BackColor = clr "#1C1D23"
    $gpBox.Size = New-Object System.Drawing.Size(650, 70)
    $gpBox.Location = New-Object System.Drawing.Point(10, $y)

    $lblGP = New-Object System.Windows.Forms.Label
    $lblGP.Text = "Grandparent Theme:  $gpName"
    $lblGP.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblGP.ForeColor = $cAmber; $lblGP.AutoSize = $true
    $lblGP.Location = New-Object System.Drawing.Point(10, 10)
    $gpBox.Controls.Add($lblGP)

    $tbTheme = New-Object System.Windows.Forms.TextBox
    $tbTheme.Text = $gpName
    $tbTheme.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tbTheme.BackColor = $cInput; $tbTheme.ForeColor = $cText
    $tbTheme.BorderStyle = 'FixedSingle'
    $tbTheme.Location = New-Object System.Drawing.Point(10, 35)
    $tbTheme.Size = New-Object System.Drawing.Size(250, 24)
    $gpBox.Controls.Add($tbTheme)

    $chkSkip = New-Object System.Windows.Forms.CheckBox
    $chkSkip.Text = "Don't rename grandparent folder"
    $chkSkip.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkSkip.ForeColor = $cMuted; $chkSkip.AutoSize = $true
    $chkSkip.Location = New-Object System.Drawing.Point(280, 35)
    $gpBox.Controls.Add($chkSkip)
    $pnlRight.Controls.Add($gpBox)
    $y += 80

    # Parent Char/Adj Block
    $pBox = New-Object System.Windows.Forms.Panel
    $pBox.BackColor = clr "#16171B"
    $pBox.BorderStyle = 'FixedSingle'
    $pBox.Size = New-Object System.Drawing.Size(650, 80)
    $pBox.Location = New-Object System.Drawing.Point(10, $y)

    $lblChar = New-Object System.Windows.Forms.Label; $lblChar.Text = "Character *"
    $lblChar.ForeColor = $cMuted; $lblChar.AutoSize = $true; $lblChar.Location = New-Object System.Drawing.Point(10, 10)
    $pBox.Controls.Add($lblChar)
    $tbChar = New-Object System.Windows.Forms.TextBox; $tbChar.Text = $fills.Char
    $tbChar.BackColor = $cInput; $tbChar.ForeColor = $cText; $tbChar.BorderStyle = 'FixedSingle'
    $tbChar.Location = New-Object System.Drawing.Point(10, 30); $tbChar.Size = New-Object System.Drawing.Size(200, 24)
    $pBox.Controls.Add($tbChar)

    $lblAdj = New-Object System.Windows.Forms.Label; $lblAdj.Text = "Adjective (Optional)"
    $lblAdj.ForeColor = $cMuted; $lblAdj.AutoSize = $true; $lblAdj.Location = New-Object System.Drawing.Point(230, 10)
    $pBox.Controls.Add($lblAdj)
    $tbAdj = New-Object System.Windows.Forms.TextBox; $tbAdj.Text = $fills.Adj
    $tbAdj.BackColor = $cInput; $tbAdj.ForeColor = $cText; $tbAdj.BorderStyle = 'FixedSingle'
    $tbAdj.Location = New-Object System.Drawing.Point(230, 30); $tbAdj.Size = New-Object System.Drawing.Size(200, 24)
    $pBox.Controls.Add($tbAdj)

    $lblTh = New-Object System.Windows.Forms.Label; $lblTh.Text = "Theme"
    $lblTh.ForeColor = $cMuted; $lblTh.AutoSize = $true; $lblTh.Location = New-Object System.Drawing.Point(450, 10)
    $pBox.Controls.Add($lblTh)
    $lblThVal = New-Object System.Windows.Forms.Label; $lblThVal.Text = $tbTheme.Text
    $lblThVal.BackColor = clr "#181920"; $lblThVal.ForeColor = clr "#4A4D5C"
    $lblThVal.BorderStyle = 'FixedSingle'; $lblThVal.Location = New-Object System.Drawing.Point(450, 30)
    $lblThVal.Size = New-Object System.Drawing.Size(180, 22); $lblThVal.TextAlign = 'MiddleLeft'
    $pBox.Controls.Add($lblThVal)
    $pnlRight.Controls.Add($pBox)
    $y += 90

    # Files List
    $pnlFiles = New-Object System.Windows.Forms.Panel
    $pnlFiles.AutoScroll = $true; $pnlFiles.BackColor = clr "#111214"
    $pnlFiles.Location = New-Object System.Drawing.Point(10, $y)
    $pnlFiles.Size = New-Object System.Drawing.Size(650, 350)
    $pnlRight.Controls.Add($pnlFiles)

    $script:fileRows = @()
    $files = Get-ChildItem -Path $folderPath -File | Sort-Object Name
    $fy = 0
    foreach ($fi in $files) {
        $parsed = ParseFile $fi.Name
        $fRow = New-Object System.Windows.Forms.Panel
        $fRow.BackColor = if (($script:fileRows.Count % 2) -eq 0) { clr "#16171B" } else { clr "#1A1C22" }
        $fRow.Size = New-Object System.Drawing.Size(620, 50); $fRow.Location = New-Object System.Drawing.Point(0, $fy)

        $sBadge = New-Object System.Windows.Forms.TextBox
        $sBadge.Text = $parsed.Suffix; $sBadge.BackColor = $cInput; $sBadge.ForeColor = $cAmber
        $sBadge.BorderStyle = 'FixedSingle'; $sBadge.TextAlign = 'Center'; $sBadge.Font = New-Object System.Drawing.Font("Consolas", 8)
        $sBadge.Location = New-Object System.Drawing.Point(10, 15); $sBadge.Size = New-Object System.Drawing.Size(60, 20)
        $fRow.Controls.Add($sBadge)

        $lOld = New-Object System.Windows.Forms.Label
        $lOld.Text = $fi.Name; $lOld.ForeColor = clr "#6B6E7A"; $lOld.Font = New-Object System.Drawing.Font("Consolas", 8)
        $lOld.Location = New-Object System.Drawing.Point(80, 15); $lOld.Size = New-Object System.Drawing.Size(220, 20)
        $fRow.Controls.Add($lOld)

        $lArr = New-Object System.Windows.Forms.Label
        $lArr.Text = "->"; $lArr.ForeColor = $cMuted; $lArr.Location = New-Object System.Drawing.Point(300, 15); $lArr.Size = New-Object System.Drawing.Size(25, 20)
        $fRow.Controls.Add($lArr)

        $lNew = New-Object System.Windows.Forms.Label
        $lNew.ForeColor = clr "#4CAF72"; $lNew.Font = New-Object System.Drawing.Font("Consolas", 8)
        $lNew.Location = New-Object System.Drawing.Point(330, 15); $lNew.Size = New-Object System.Drawing.Size(280, 20)
        $fRow.Controls.Add($lNew)

        $script:fileRows += [PSCustomObject]@{ FileInfo = $fi; SuffixBox = $sBadge; NewLbl = $lNew; Ext = $parsed.Extension }
        $sBadge.Add_TextChanged({ script:Update-RenamePreview })
        $pnlFiles.Controls.Add($fRow)
        $fy += 50
    }

    # Action Buttons
    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = if ($script:qIndex -eq ($script:queue.Count - 1)) { "Apply & Finish" } else { "Apply & Next" }
    $btnNext.BackColor = clr "#4CAF72"; $btnNext.ForeColor = clr "#FFFFFF"
    $btnNext.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnNext.FlatStyle = 'Flat'; $btnNext.FlatAppearance.BorderSize = 0
    $btnNext.Size = New-Object System.Drawing.Size(200, 45); $btnNext.Location = New-Object System.Drawing.Point(460, ($y + 360))
    $pnlRight.Controls.Add($btnNext)

    # Save to global scope for Apply & Updates
    $script:rpTBChar  = $tbChar
    $script:rpTBAdj   = $tbAdj
    $script:rpTBTheme = $tbTheme
    $script:rpLblChar = $lblCharTitle
    $script:rpLblAdj  = $lblAdjTitle
    $script:rpChkSkip = $chkSkip
    $script:rpAnchor  = $anchorFile
    $script:rpFolder  = $folderPath
    $script:rpDiGrand = $diGrand
    $script:btnNext   = $btnNext

    # Wire up events globally so they cleanly trigger the Preview refresh
    $tbTheme.Add_TextChanged({ $lblThVal.Text = $tbTheme.Text; script:Update-RenamePreview })
    $tbChar.Add_TextChanged({ script:Update-RenamePreview })
    $tbAdj.Add_TextChanged({ script:Update-RenamePreview })

    # Initial data push to trigger both Canvas & File List populations
    script:Update-RenamePreview

    $btnNext.Add_Click({
        $script:btnNext.Text = "Processing..."; $script:btnNext.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()

        # 1. Color Patching
        $allTextFiles = Get-ChildItem -Path $script:tempWork -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
        foreach ($file in $allTextFiles) {
            $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
            $modified = $false
            foreach ($slot in $script:uiSlots) {
                $selName = $slot.Combo.Text
                if ($LibraryColors.Contains($selName)) {
                    $newHex = $LibraryColors[$selName].ToUpper()
                    $oldHex = $slot.OldHex.ToUpper()
                    $oldHex6 = $oldHex.Substring(0,7); $newHex6 = $newHex.Substring(0,7)
                    if ($content -match "(?i)$oldHex") { $content = $content -ireplace [regex]::Escape($oldHex), $newHex; $modified = $true }
                    if ($content -match "(?i)$oldHex6") { $content = $content -ireplace [regex]::Escape($oldHex6), $newHex6; $modified = $true }
                }
            }
            if ($modified) { [System.IO.File]::WriteAllText($file.FullName, $content, (New-Object System.Text.UTF8Encoding($false))) }
        }

        # 2. Repack 3MF to original name (so it gets renamed with everything else)
        Rename-Item $script:rpAnchor.FullName "$($script:rpAnchor.BaseName)_old.3mf" -Force
        $zip = [System.IO.Compression.ZipFile]::Open($script:rpAnchor.FullName, 'Create')
        Get-ChildItem $script:tempWork -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($script:tempWork.Length).TrimStart('\','/').Replace('\','/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
        }
        $zip.Dispose()

        # 3. Rename Logic
        $oldFolder = $script:rpFolder
        $newFolder = $oldFolder
        $oldGrand  = if ($script:rpDiGrand) { $script:rpDiGrand.FullName } else { "" }
        $newGrand  = $oldGrand

        # Files First
        foreach ($r in $script:fileRows) {
            $oldPath = $r.FileInfo.FullName
            $newPath = Join-Path $oldFolder $r.NewLbl.Text
            if ($oldPath -ne $newPath -and (Test-Path $oldPath)) { Rename-Item $oldPath (Split-Path $newPath -Leaf) -Force }
        }

        # Parent Folder
        $ch = $script:rpTBChar.Text -replace '[^a-zA-Z0-9]', ''; $ad = $script:rpTBAdj.Text -replace '[^a-zA-Z0-9]', ''; $th = $script:rpTBTheme.Text -replace '[^a-zA-Z0-9]', ''
        $pParts = @(); if ($ch) { $pParts += $ch }; if ($ad) { $pParts += $ad }; if ($th) { $pParts += $th }
        $newParentName = $pParts -join '_'
        if ($newParentName -ne '' -and $newParentName -ne (Split-Path $oldFolder -Leaf)) {
            $newFolder = Join-Path (Split-Path $oldFolder -Parent) $newParentName
            Rename-Item $oldFolder $newParentName -Force
        }

        # Grandparent Folder
        if (-not $script:rpChkSkip.Checked -and $th -ne '' -and $oldGrand -ne '' -and $th -ne (Split-Path $oldGrand -Leaf)) {
            $newGrand = Join-Path (Split-Path $oldGrand -Parent) $th
            Rename-Item $oldGrand $th -Force
            # Adjust newFolder path if GP changed
            $newFolder = $newFolder.Replace($oldGrand, $newGrand)
        }

        # 4. Queue Path Maintenance (Fix broken paths for future items)
        for ($i = $script:qIndex + 1; $i -lt $script:queue.Count; $i++) {
            $qPath = $script:queue[$i]
            if ($oldGrand -ne $newGrand -and $qPath.StartsWith($oldGrand)) { $qPath = $qPath.Replace($oldGrand, $newGrand) }
            if ($oldFolder -ne $newFolder -and $qPath.StartsWith($oldFolder)) { $qPath = $qPath.Replace($oldFolder, $newFolder) }
            $script:queue[$i] = $qPath
        }

        $script:qIndex++
        Load-QueueItem
    })

    # Trigger resize to snap layout
    $form.Width = $form.Width + 1
}

# --- 6. Form Scaling Engine ---
$script:resizeTimer = New-Object System.Windows.Forms.Timer
$script:resizeTimer.Interval = 50
$script:resizeTimer.Add_Tick({
    $script:resizeTimer.Stop()
    $form.SuspendLayout()

    # Right panel width is 680
    $availW = $form.ClientSize.Width - 680
    $availH = $form.ClientSize.Height - 40
    $sq = [math]::Min($availW, $availH)
    if ($sq -lt 300) { $sq = 300 }

    $card.Location = New-Object System.Drawing.Point(20, 20)
    $card.Size = New-Object System.Drawing.Size($sq, $sq)

    $scale = $sq / 512.0
    foreach ($item in $script:scaleElements) {
        $item.Ctrl.Location = New-Object System.Drawing.Point([int]($item.X * $scale), [int]($item.Y * $scale))
        if ($item.Ctrl -is [System.Windows.Forms.ComboBox]) { $item.Ctrl.Width = [int]($item.W * $scale) }
        else { $item.Ctrl.Size = New-Object System.Drawing.Size([int]($item.W * $scale), [int]($item.H * $scale)) }

        if ($item.Font -gt 0) {
            $newFont = [float]($item.Font * $scale)
            if ($newFont -lt 4) { $newFont = 4 }

            $oldFont = $item.Ctrl.Font
            $item.Ctrl.Font = New-Object System.Drawing.Font($oldFont.FontFamily, $newFont, $oldFont.Style)
        }
    }

    $pnlRight.Location = New-Object System.Drawing.Point(($card.Right + 20), 20)
    $pnlRight.Size = New-Object System.Drawing.Size(660, ($form.ClientSize.Height - 40))
    if ($script:btnNext) {
        $script:btnNext.Top = $pnlRight.Height - 60
    }

    $form.ResumeLayout()
    $card.Invalidate()
})

$form.Add_Resize({ $script:resizeTimer.Stop(); $script:resizeTimer.Start() })

$form.Add_Shown({ Load-QueueItem })
$form.Add_FormClosed({ if (Test-Path $script:tempWork) { Remove-Item $script:tempWork -Recurse -Force -ErrorAction SilentlyContinue } })

[System.Windows.Forms.Application]::Run($form)