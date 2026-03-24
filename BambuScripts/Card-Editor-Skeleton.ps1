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

# --- 2. Advanced Helpers ---
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
    $stem = $anchorName -replace '(?i)\.gcode\.3mf$|\.3mf$', ''
    $stem = $stem -replace '(?i)_Full$', ''
    $escaped = [regex]::Escape($gpName)

    if ($gpName -ne '' -and $stem -imatch "^(.+)_${escaped}$") { $prefix = $Matches[1] } else { $prefix = $stem }
    $parts = $prefix -split '_'
    if ($parts.Count -ge 2) { return @{ Char = $parts[0]; Adj = ($parts[1..($parts.Count-1)] -join '') } }
    return @{ Char = $prefix; Adj = '' }
}

function clr($hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { return [System.Drawing.Color]::Gray }
    if ($hex.Length -ge 9) { $hex = "#" + $hex.Substring(1,6) }
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
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
                if (-not $colorMap.ContainsKey($key)) {
                    $colorMap[$key] = [System.Drawing.Color]::FromArgb($px.A, $rng.Next(0,256), $rng.Next(0,256), $rng.Next(0,256))
                }
                $bmp.SetPixel($x, $y, $colorMap[$key])
            }
        }
        $bmp.Save($destPath) | Out-Null
        $bmp.Dispose()
        return $true
    } catch { return $false }
}

function Show-ImageViewer($imagePath, $title) {
    $vForm = New-Object System.Windows.Forms.Form
    $vForm.Text = $title
    $vForm.Size = New-Object System.Drawing.Size(800, 820)
    $vForm.StartPosition = "CenterScreen"
    $vForm.BackColor = [System.Drawing.Color]::Black
    $vForm.MinimumSize = New-Object System.Drawing.Size(300, 300)
    $pb = New-Object System.Windows.Forms.PictureBox
    $pb.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pb.BackColor = [System.Drawing.Color]::Black
    try {
        $bytes = [System.IO.File]::ReadAllBytes($imagePath)
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        $pb.Image = [System.Drawing.Image]::FromStream($ms)
    } catch {}
    $vForm.Controls.Add($pb)
    $vForm.ShowDialog() | Out-Null
    if ($pb.Image) { $pb.Image.Dispose() }
}

$cBG      = clr "#000000"; $cUI      = clr "#16171B"
$cText    = clr "#FFFFFF"; $cMuted   = clr "#A0A0A0"
$cAccent  = clr "#FFD700"; $cInput   = clr "#1E2028"
$cGrayTxt = clr "#B4B4B4"; $cGreen   = clr "#4CAF72"
$cRed     = clr "#D95F5F"; $cBorder  = clr "#2A2C35"
$cAmber   = clr "#E8A135"

# --- 3. Build Grandparent-Grouped Queue ---
$gpQueue = [ordered]@{}

$foundFiles = @()
if ($args.Count -eq 0) {
    $pickedFolder = [ModernFolderPicker]::Pick([IntPtr]::Zero, "Select a Master Folder to scan for Full.3mf files")
    if (-not $pickedFolder) { exit }
    $foundFiles = Get-ChildItem -Path $pickedFolder -Filter "*Full.3mf" -Recurse -File
} else {
    foreach ($p in $args) {
        if (Test-Path $p -PathType Container) {
            $foundFiles += Get-ChildItem -Path $p -Filter "*Full.3mf" -Recurse -File
        } elseif ($p -match '(?i)Full\.3mf$') {
            $foundFiles += Get-Item $p
        }
    }
}

if ($foundFiles.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No Full.3mf files found.", "Empty Queue") | Out-Null
    exit
}

foreach ($f in $foundFiles) {
    $parentPath = $f.DirectoryName
    $gp = $f.Directory.Parent
    $gpPath = if ($gp) { $gp.FullName } else { "ROOT_" + $parentPath }

    if (-not $gpQueue.Contains($gpPath)) { $gpQueue[$gpPath] = [ordered]@{} }
    if (-not $gpQueue[$gpPath].Contains($parentPath)) { $gpQueue[$gpPath][$parentPath] = $f }
}

# --- 4. Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Batch Pre-Flight Editor - Queue Dashboard"
$form.ClientSize = New-Object System.Drawing.Size(1550, 850)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)
$form.BackColor = $cUI; $form.StartPosition = "CenterScreen"

$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Height = 60; $pnlTop.Dock = 'Top'; $pnlTop.BackColor = clr "#1C1D23"
$form.Controls.Add($pnlTop)

$lblGlobalTitle = New-Object System.Windows.Forms.Label
$lblGlobalTitle.Text = "Loading files into queue..."
$lblGlobalTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblGlobalTitle.ForeColor = $cText; $lblGlobalTitle.AutoSize = $true
$lblGlobalTitle.Location = New-Object System.Drawing.Point(15, 15)
$pnlTop.Controls.Add($lblGlobalTitle)

$btnProcessAll = New-Object System.Windows.Forms.Button
$btnProcessAll.Text = "Process All Remaining"
$btnProcessAll.BackColor = clr "#4CAF72"; $btnProcessAll.ForeColor = clr "#FFFFFF"
$btnProcessAll.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnProcessAll.FlatStyle = 'Flat'; $btnProcessAll.FlatAppearance.BorderSize = 0
$btnProcessAll.Size = New-Object System.Drawing.Size(220, 40)
$btnProcessAll.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Right"
$btnProcessAll.Location = New-Object System.Drawing.Point(1300, 10)
$btnProcessAll.Enabled = $false
$pnlTop.Controls.Add($btnProcessAll)

$mainScroll = New-Object System.Windows.Forms.Panel
$mainScroll.Dock = 'Fill'; $mainScroll.AutoScroll = $true; $mainScroll.BackColor = clr "#0D0E10"
$form.Controls.Add($mainScroll)
$mainScroll.BringToFront()

$script:jobs = [System.Collections.Generic.List[object]]::new()

function Add-ScaledElement($pJob, $targetPanel, $ctrl, $bx, $by, $bw, $bh, $bf) {
    $pJob.ScaleElements.Add([PSCustomObject]@{ Target = $targetPanel; Ctrl = $ctrl; X = $bx; Y = $by; W = $bw; H = $bh; Font = $bf })
    $ctrl.AutoSize = $false
    $targetPanel.Controls.Add($ctrl)
}

# --- 5. Sync Functions ---
function Update-ParentPreview($pJob, $gpJob) {
    $ch = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
    $ad = $pJob.TBAdj.Text -replace '[^a-zA-Z0-9]', ''
    $th = $gpJob.TBTheme.Text -replace '[^a-zA-Z0-9]', ''

    $pJob.LblChar.Text = $ch
    $pJob.LblAdj.Text  = if ($ad) { "($ad)" } else { "" }

    foreach ($r in $pJob.FileRows) {
        $sf = $r.SuffixBox.Text -replace '[^a-zA-Z0-9]', ''
        $parts = [System.Collections.Generic.List[string]]::new()
        if ($ch) { $parts.Add($ch) }
        if ($ad) { $parts.Add($ad) }
        if ($th) { $parts.Add($th) }
        if ($sf) { $parts.Add($sf) }
        $r.NewLbl.Text = ($parts -join '_') + $r.Ext
    }
}

# --- 6. Global Apply Function ---
function Apply-GpJob($gpJob) {
    if ($gpJob.IsDone -or $gpJob.Container.IsDisposed) { return }
    $gpJob.BtnApply.Text = "Processing..."; $gpJob.BtnApply.Enabled = $false
    $gpJob.Container.BackColor = clr "#2A2C35"
    [System.Windows.Forms.Application]::DoEvents()

    $th = $gpJob.TBTheme.Text -replace '[^a-zA-Z0-9]', ''
    $oldGrand  = if ($gpJob.DiGrand) { $gpJob.DiGrand.FullName } else { "" }
    $newGrand  = $oldGrand

    # Process all parents sequentially
    foreach ($pJob in $gpJob.Parents) {
        $allTextFiles = Get-ChildItem -Path $pJob.TempWork -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
        foreach ($file in $allTextFiles) {
            $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
            $modified = $false
            foreach ($slot in $pJob.UISlots) {
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

        Rename-Item $pJob.AnchorFile.FullName "$($pJob.AnchorFile.BaseName)_old.3mf" -Force
        $cleanChar = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
        $cleanAdj  = $pJob.TBAdj.Text -replace '[^a-zA-Z0-9]', ''
        $newName = if ($cleanAdj) { "${cleanChar}_${cleanAdj}_Full.3mf" } else { "${cleanChar}_Full.3mf" }
        $newFilePath = Join-Path $pJob.FolderPath $newName

        $zip = [System.IO.Compression.ZipFile]::Open($newFilePath, 'Create')
        Get-ChildItem $pJob.TempWork -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($pJob.TempWork.Length).TrimStart('\','/').Replace('\','/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
        }
        $zip.Dispose()
        Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue

        foreach ($r in $pJob.FileRows) {
            $oldPath = $r.OldPath
            $targetName = $r.NewLbl.Text
            $newPath = Join-Path $pJob.FolderPath $targetName
            if ($oldPath -ne $newPath -and (Test-Path $oldPath)) { Rename-Item $oldPath $targetName -Force }
        }

        $pParts = @(); if ($cleanChar) { $pParts += $cleanChar }; if ($cleanAdj) { $pParts += $cleanAdj }; if ($th) { $pParts += $th }
        $newParentName = $pParts -join '_'
        $oldFolder = $pJob.FolderPath
        if ($newParentName -ne '' -and $newParentName -ne (Split-Path $oldFolder -Leaf)) {
            $newFolder = Join-Path (Split-Path $oldFolder -Parent) $newParentName
            Rename-Item $oldFolder $newParentName -Force
            $pJob.FolderPath = $newFolder
        }
    }

    # Rename Grandparent Folder
    if (-not $gpJob.ChkSkip.Checked -and $th -ne '' -and $oldGrand -ne '' -and $th -ne (Split-Path $oldGrand -Leaf)) {
        $newGrand = Join-Path (Split-Path $oldGrand -Parent) $th
        Rename-Item $oldGrand $th -Force
    }

    # Update Path Hierarchy for Downstream Queues
    if ($oldGrand -ne '' -and $oldGrand -ne $newGrand) {
        foreach ($otherGp in $script:jobs) {
            if (-not $otherGp.IsDone -and -not $otherGp.Container.IsDisposed) {
                if ($otherGp.GpPath.StartsWith($oldGrand)) {
                    $otherGp.GpPath = $otherGp.GpPath.Replace($oldGrand, $newGrand)
                    $otherGp.DiGrand = [System.IO.DirectoryInfo]::new($otherGp.GpPath)
                    foreach ($otherP in $otherGp.Parents) {
                        $otherP.FolderPath = $otherP.FolderPath.Replace($oldGrand, $newGrand)
                        foreach ($fr in $otherP.FileRows) { $fr.OldPath = $fr.OldPath.Replace($oldGrand, $newGrand) }
                    }
                }
            }
        }
    }

    $gpJob.BtnApply.Text = "Applied"
    $gpJob.Container.Enabled = $false
    $gpJob.IsDone = $true
}

# --- 7. Build Job Rows ---
function Build-PJob($parentPath, $anchorFile, $gpJob) {
    $tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -ItemType Directory -Path $tempWork | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($anchorFile.FullName, $tempWork)

    $activeSlots = @()
    $projPath   = Join-Path $tempWork "Metadata\project_settings.config"
    $modSetPath = Join-Path $tempWork "Metadata\model_settings.config"
    $modelFile  = (Get-ChildItem -Path $tempWork -Filter '3dmodel.model' -Recurse | Select-Object -First 1)

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

    $pJob = @{
        FolderPath = $parentPath; AnchorFile = $anchorFile; TempWork = $tempWork
        UISlots = [System.Collections.Generic.List[object]]::new()
        FileRows = [System.Collections.Generic.List[object]]::new()
        ScaleElements = [System.Collections.Generic.List[object]]::new()
    }

    $rowPanel = New-Object System.Windows.Forms.Panel
    $rowPanel.BackColor = clr "#16171B"; $rowPanel.BorderStyle = 'FixedSingle'
    $gpJob.ParentListPanel.Controls.Add($rowPanel)
    $pJob.RowPanel = $rowPanel

    $card = New-Object System.Windows.Forms.Panel
    $card.BackColor = $cBG; $card.BorderStyle = 'FixedSingle'
    $rowPanel.Controls.Add($card)
    $pJob.Card = $card

    $pickCard = New-Object System.Windows.Forms.Panel
    $pickCard.BackColor = $cBG; $pickCard.BorderStyle = 'FixedSingle'
    $rowPanel.Controls.Add($pickCard)
    $pJob.PickCard = $pickCard

    $pnlRight = New-Object System.Windows.Forms.Panel
    $pnlRight.BackColor = clr "#16171B"
    $rowPanel.Controls.Add($pnlRight)
    $pJob.PnlRight = $pnlRight

    # --- Left Canvas: Plate Image Handling ---
    $localPng = Get-ChildItem -Path $parentPath -Filter "*.png" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $isReplaced = $false
    if ($localPng) {
        $plateImgPath = $localPng.FullName
        $isReplaced = $true
    } else {
        $plateImgPath = Join-Path $tempWork "Metadata\plate_1.png"
    }

    $pbModel = New-Object System.Windows.Forms.PictureBox
    $pbModel.SizeMode = 'Zoom'; $pbModel.BackColor = $cBG; $pbModel.AllowDrop = $true
    if (Test-Path $plateImgPath) {
        $fs = New-Object System.IO.FileStream($plateImgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pbModel.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pbModel.Tag = @{ P = $pJob; ImgPath = $plateImgPath }
        $pbModel.Add_DoubleClick({ $t = $this.Tag; Show-ImageViewer $t.ImgPath "Plate Preview" })
    }
    Add-ScaledElement $pJob $card $pbModel 10 80 250 360 0

    $lblImgStatus = New-Object System.Windows.Forms.Label
    $lblImgStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lblImgStatus.BackColor = $cBG; $lblImgStatus.TextAlign = 'TopLeft'
    if ($isReplaced) {
        $lblImgStatus.Text = "[REPLACED]"; $lblImgStatus.ForeColor = $cGreen
    } else {
        $lblImgStatus.Text = "[DEFAULT]"; $lblImgStatus.ForeColor = $cAmber
    }
    Add-ScaledElement $pJob $card $lblImgStatus 15 85 100 15 8
    $pJob.ImgStatusLbl = $lblImgStatus

    $pbModel.Add_DragEnter({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($files[0] -match '(?i)\.(png|jpg|jpeg)$') { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
        }
    })
    $pbModel.Add_DragDrop({
        param($s, $e)
        $p = $s.Tag.P
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        $dropped = $files[0]
        if ($dropped -match '(?i)\.(png|jpg|jpeg)$') {
            $dest = Join-Path $p.FolderPath (Split-Path $dropped -Leaf)
            if ($dropped -ne $dest) { Copy-Item -Path $dropped -Destination $dest -Force }
            $fs = New-Object System.IO.FileStream($dest, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            if ($s.Image) { $s.Image.Dispose() }
            $s.Image = [System.Drawing.Image]::FromStream($fs)
            $s.Tag.ImgPath = $dest
            $fs.Close()
            $p.ImgStatusLbl.Text = "[REPLACED]"; $p.ImgStatusLbl.ForeColor = clr "#4CAF72"
        }
    })

    $lblAdjTitle = New-Object System.Windows.Forms.Label
    $lblAdjTitle.BackColor = $cBG; $lblAdjTitle.ForeColor = $cAccent; $lblAdjTitle.TextAlign = 'TopRight'
    Add-ScaledElement $pJob $card $lblAdjTitle 320 15 180 40 18
    $pJob.LblAdj = $lblAdjTitle

    $lblCharTitle = New-Object System.Windows.Forms.Label
    $lblCharTitle.BackColor = $cBG; $lblCharTitle.ForeColor = $cText; $lblCharTitle.TextAlign = 'TopRight'
    Add-ScaledElement $pJob $card $lblCharTitle 60 15 250 40 18
    $pJob.LblChar = $lblCharTitle

    $lblSkipTime = New-Object System.Windows.Forms.Label
    $lblSkipTime.Text = "Skip Time: 18 min"
    $lblSkipTime.BackColor = $cBG; $lblSkipTime.ForeColor = $cText; $lblSkipTime.TextAlign = 'BottomLeft'
    Add-ScaledElement $pJob $card $lblSkipTime 10 472 250 30 14

    $startY = 80; $boxSize = 76; $slotSpacing = $boxSize + 10; $boxX = 512 - 10 - $boxSize
    for ($i = 0; $i -lt $activeSlots.Count; $i++) {
        $slotData = $activeSlots[$i]
        $swatch = New-Object System.Windows.Forms.Panel
        $swatch.BorderStyle = 'FixedSingle'
        try { $swatch.BackColor = clr $slotData.OldHex } catch { $swatch.BackColor = clr "#333333" }
        $swatch.Tag = ($i + 1).ToString()
        $swatch.Add_Paint({
            param($s, $e)
            $bg = $s.BackColor
            $lum = (0.299 * $bg.R + 0.587 * $bg.G + 0.114 * $bg.B)
            $textColor = if ($lum -gt 128) { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::White }
            $brush = New-Object System.Drawing.SolidBrush($textColor)
            $e.Graphics.DrawString($s.Tag, $s.Font, $brush, [int]($s.Width*0.35), [int]($s.Height*0.25))
            $brush.Dispose()
        })
        Add-ScaledElement $pJob $card $swatch $boxX $startY $boxSize $boxSize 16

        $lblStatus = New-Object System.Windows.Forms.Label
        $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $lblStatus.TextAlign = 'BottomRight'; $lblStatus.BackColor = $cBG
        if ($slotData.Name) { $lblStatus.Text = "[MATCHED]"; $lblStatus.ForeColor = $cGreen }
        else { $lblStatus.Text = "[UNMATCHED]"; $lblStatus.ForeColor = $cRed }
        Add-ScaledElement $pJob $card $lblStatus ($boxX - 180) ($startY) 170 15 8

        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.BackColor = $cInput; $combo.ForeColor = $cText; $combo.DropDownStyle = 'DropDown'
        $combo.AutoCompleteMode = 'SuggestAppend'; $combo.AutoCompleteSource = 'ListItems'
        foreach ($k in $LibraryColors.Keys) { [void]$combo.Items.Add($k) }
        if ($slotData.Name) { $combo.Text = $slotData.Name } else { $combo.Text = "Select Color..." }

        $combo.add_MouseWheel({
            param($s, $e)
            $hw = $e -as [System.Windows.Forms.HandledMouseEventArgs]
            if ($hw) { $hw.Handled = $true }
        })

        $combo.Tag = @{ Swatch = $swatch; StatusLbl = $lblStatus; OrigName = $slotData.Name }
        $combo.add_TextChanged({
            param($s, $e)
            $data = $s.Tag
            if ($LibraryColors.Contains($s.Text)) {
                $data.Swatch.BackColor = clr $LibraryColors[$s.Text]
                $data.Swatch.Invalidate()
            }
            if ($s.Text -eq $data.OrigName) {
                if ($data.OrigName) { $data.StatusLbl.Text = "[MATCHED]"; $data.StatusLbl.ForeColor = clr "#4CAF72" }
                else { $data.StatusLbl.Text = "[UNMATCHED]"; $data.StatusLbl.ForeColor = clr "#D95F5F" }
            } else {
                $data.StatusLbl.Text = "[CHANGED]"; $data.StatusLbl.ForeColor = clr "#E8A135"
            }
            $s.SelectionStart = 0; $s.SelectionLength = 0
        }.GetNewClosure())
        Add-ScaledElement $pJob $card $combo ($boxX - 180) ($startY + 15) 170 25 10

        $lblGrams = New-Object System.Windows.Forms.Label
        $lblGrams.Text = "$($slotData.Grams) g"; $lblGrams.ForeColor = $cGrayTxt; $lblGrams.TextAlign = 'TopRight'
        Add-ScaledElement $pJob $card $lblGrams ($boxX - 90) ($startY + 45) 80 25 10

        $pJob.UISlots.Add([PSCustomObject]@{ OldHex = $slotData.OldHex; Combo = $combo })
        $startY += $slotSpacing
    }
    $pbModel.SendToBack()

    # --- Middle Canvas: Pick / Merge Check Image ---
    $gcodeFile = Get-ChildItem -Path $parentPath -Filter "*Full.gcode.3mf" | Select-Object -First 1
    $pickPath = $null
    if ($gcodeFile) {
        $zip = $null
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)(^|/)pick_1\.png$" } | Select-Object -First 1
            if ($pickEntry) {
                $rawPickPath = Join-Path $tempWork "pick_1_raw.png"
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $rawPickPath, $true)
                $pickPath = Join-Path $tempWork "pick_1.png"
                Invoke-RandomizePickColors $rawPickPath $pickPath | Out-Null
                if (-not (Test-Path $pickPath)) { $pickPath = $rawPickPath }
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    }

    $pbPick = New-Object System.Windows.Forms.PictureBox
    $pbPick.SizeMode = 'Zoom'; $pbPick.BackColor = clr "#0D0E10"
    Add-ScaledElement $pJob $pickCard $pbPick 0 0 512 512 0

    if ($pickPath -and (Test-Path $pickPath)) {
        $fs = New-Object System.IO.FileStream($pickPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pbPick.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pbPick.Tag = @{ Path = $pickPath; Title = "Merge Pick Preview" }
        $pbPick.Add_DoubleClick({ $t = $this.Tag; Show-ImageViewer $t.Path $t.Title })
    } else {
        $lblNoPick = New-Object System.Windows.Forms.Label
        $lblNoPick.Text = "[NO GCODE]"
        $lblNoPick.ForeColor = $cRed
        $lblNoPick.TextAlign = 'MiddleCenter'
        Add-ScaledElement $pJob $pickCard $lblNoPick 0 0 512 512 16
        $lblNoPick.BringToFront()
    }

    # --- Right Panel (Tasks, Rename Settings, File List) ---
    $diParent = [System.IO.DirectoryInfo]::new($parentPath)
    $gpName = if ($gpJob.DiGrand) { $gpJob.DiGrand.Name } else { "" }
    $fills = SmartFill $anchorFile.Name $gpName

    $y = 10
    $lblFolder = New-Object System.Windows.Forms.Label
    $lblFolder.Text = "Folder: $($diParent.Name)"
    $lblFolder.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblFolder.ForeColor = $cText; $lblFolder.AutoSize = $true
    $lblFolder.Location = New-Object System.Drawing.Point(10, $y)
    $pnlRight.Controls.Add($lblFolder)

    $btnRemoveP = New-Object System.Windows.Forms.Button
    $btnRemoveP.Text = "Remove Folder"
    $btnRemoveP.BackColor = $cRed; $btnRemoveP.ForeColor = clr "#FFFFFF"
    $btnRemoveP.FlatStyle = 'Flat'; $btnRemoveP.FlatAppearance.BorderSize = 0
    $btnRemoveP.Size = New-Object System.Drawing.Size(120, 25)
    $btnRemoveP.Location = New-Object System.Drawing.Point(500, $y)
    $pnlRight.Controls.Add($btnRemoveP)
    $btnRemoveP.Tag = @{ P = $pJob; G = $gpJob }
    $btnRemoveP.Add_Click({
        $t = $this.Tag
        $t.P.RowPanel.Dispose()
        $t.G.Parents.Remove($t.P) | Out-Null
        if ($t.G.Parents.Count -eq 0) {
            $t.G.Container.Dispose()
            $script:jobs.Remove($t.G) | Out-Null
        }
        $script:resizeTimer.Start()
    })
    $pJob.BtnRemove = $btnRemoveP

    $y += 30

    $pTasks = New-Object System.Windows.Forms.Panel
    $pTasks.BackColor = clr "#1C1D23"; $pTasks.BorderStyle = 'FixedSingle'
    $pTasks.Size = New-Object System.Drawing.Size(620, 75)
    $pTasks.Location = New-Object System.Drawing.Point(10, $y)

    $lblTasks = New-Object System.Windows.Forms.Label
    $lblTasks.Text = "Tasks:"
    $lblTasks.ForeColor = $cMuted; $lblTasks.AutoSize = $true
    $lblTasks.Location = New-Object System.Drawing.Point(10, 10)
    $pTasks.Controls.Add($lblTasks)

    $chkMerge = New-Object System.Windows.Forms.CheckBox
    $chkMerge.Text = "Merge"; $chkMerge.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkMerge.ForeColor = $cText; $chkMerge.AutoSize = $true; $chkMerge.Checked = $true
    $chkMerge.Location = New-Object System.Drawing.Point(60, 9)
    $pTasks.Controls.Add($chkMerge)

    $chkSlice = New-Object System.Windows.Forms.CheckBox
    $chkSlice.Text = "Slice / Export Gcode"; $chkSlice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkSlice.ForeColor = $cText; $chkSlice.AutoSize = $true; $chkSlice.Checked = $true
    $chkSlice.Location = New-Object System.Drawing.Point(140, 9)
    $pTasks.Controls.Add($chkSlice)

    $chkExtract = New-Object System.Windows.Forms.CheckBox
    $chkExtract.Text = "Extract Data"; $chkExtract.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkExtract.ForeColor = $cText; $chkExtract.AutoSize = $true; $chkExtract.Checked = $true
    $chkExtract.Location = New-Object System.Drawing.Point(300, 9)
    $pTasks.Controls.Add($chkExtract)

    $chkImage = New-Object System.Windows.Forms.CheckBox
    $chkImage.Text = "Generate Image Card"; $chkImage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkImage.ForeColor = $cText; $chkImage.AutoSize = $true; $chkImage.Checked = $true
    $chkImage.Location = New-Object System.Drawing.Point(410, 9)
    $pTasks.Controls.Add($chkImage)

    $btnSelAll = New-Object System.Windows.Forms.Button
    $btnSelAll.Text = "Select All"; $btnSelAll.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnSelAll.BackColor = clr "#2A2C35"; $btnSelAll.ForeColor = $cText
    $btnSelAll.FlatStyle = 'Flat'; $btnSelAll.FlatAppearance.BorderSize = 0
    $btnSelAll.Size = New-Object System.Drawing.Size(100, 25)
    $btnSelAll.Location = New-Object System.Drawing.Point(10, 40)
    $pTasks.Controls.Add($btnSelAll)

    $btnDeselAll = New-Object System.Windows.Forms.Button
    $btnDeselAll.Text = "Deselect All"; $btnDeselAll.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnDeselAll.BackColor = clr "#2A2C35"; $btnDeselAll.ForeColor = $cText
    $btnDeselAll.FlatStyle = 'Flat'; $btnDeselAll.FlatAppearance.BorderSize = 0
    $btnDeselAll.Size = New-Object System.Drawing.Size(100, 25)
    $btnDeselAll.Location = New-Object System.Drawing.Point(120, 40)
    $pTasks.Controls.Add($btnDeselAll)

    $btnRevert = New-Object System.Windows.Forms.Button
    $btnRevert.Text = "Revert Merge"; $btnRevert.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnRevert.BackColor = clr "#D95F5F"; $btnRevert.ForeColor = clr "#FFFFFF"
    $btnRevert.FlatStyle = 'Flat'; $btnRevert.FlatAppearance.BorderSize = 0
    $btnRevert.Size = New-Object System.Drawing.Size(100, 25)
    $btnRevert.Location = New-Object System.Drawing.Point(230, 40)
    $pTasks.Controls.Add($btnRevert)

    $pnlRight.Controls.Add($pTasks)

    $pJob.ChkMerge = $chkMerge; $pJob.ChkSlice = $chkSlice
    $pJob.ChkExtract = $chkExtract; $pJob.ChkImage = $chkImage

    $tasksData = @{ Merge = $chkMerge; Slice = $chkSlice; Extract = $chkExtract; Image = $chkImage }
    $chkSlice.Tag = $tasksData
    $chkSlice.Add_CheckedChanged({ if ($this.Checked) { $this.Tag.Extract.Checked = $true } })
    $chkImage.Tag = $tasksData
    $chkImage.Add_CheckedChanged({ if ($this.Checked) { $this.Tag.Extract.Checked = $true } })
    $chkExtract.Tag = $tasksData
    $chkExtract.Add_CheckedChanged({
        if (-not $this.Checked) {
            if ($this.Tag.Slice.Checked) { $this.Tag.Slice.Checked = $false }
            if ($this.Tag.Image.Checked) { $this.Tag.Image.Checked = $false }
        }
    })

    $btnSelAll.Tag = $tasksData
    $btnSelAll.Add_Click({
        $t = $this.Tag
        $t.Merge.Enabled = $true; $t.Slice.Enabled = $true; $t.Extract.Enabled = $true; $t.Image.Enabled = $true
        $t.Merge.Checked = $true; $t.Slice.Checked = $true; $t.Extract.Checked = $true; $t.Image.Checked = $true
    })
    $btnDeselAll.Tag = $tasksData
    $btnDeselAll.Add_Click({
        $t = $this.Tag
        $t.Merge.Enabled = $true; $t.Slice.Enabled = $true; $t.Extract.Enabled = $true; $t.Image.Enabled = $true
        $t.Merge.Checked = $false; $t.Slice.Checked = $false; $t.Extract.Checked = $false; $t.Image.Checked = $false
    })
    $btnRevert.Tag = $tasksData
    $btnRevert.Add_Click({
        $t = $this.Tag
        $t.Merge.Checked = $false; $t.Slice.Checked = $false; $t.Extract.Checked = $false; $t.Image.Checked = $false
        $t.Merge.Enabled = $false; $t.Slice.Enabled = $false; $t.Extract.Enabled = $false; $t.Image.Enabled = $false
    })

    $y += 85

    $pBox = New-Object System.Windows.Forms.Panel
    $pBox.BackColor = clr "#1C1D23"; $pBox.BorderStyle = 'FixedSingle'
    $pBox.Size = New-Object System.Drawing.Size(620, 80)
    $pBox.Location = New-Object System.Drawing.Point(10, $y)

    $lblChar = New-Object System.Windows.Forms.Label; $lblChar.Text = "Character *"
    $lblChar.ForeColor = $cMuted; $lblChar.AutoSize = $true; $lblChar.Location = New-Object System.Drawing.Point(10, 10)
    $pBox.Controls.Add($lblChar)
    $tbChar = New-Object System.Windows.Forms.TextBox; $tbChar.Text = $fills.Char
    $tbChar.BackColor = $cInput; $tbChar.ForeColor = $cText; $tbChar.BorderStyle = 'FixedSingle'
    $tbChar.Location = New-Object System.Drawing.Point(10, 30); $tbChar.Size = New-Object System.Drawing.Size(200, 24)
    $pBox.Controls.Add($tbChar)
    $pJob.TBChar = $tbChar

    $lblAdj = New-Object System.Windows.Forms.Label; $lblAdj.Text = "Adjective (Optional)"
    $lblAdj.ForeColor = $cMuted; $lblAdj.AutoSize = $true; $lblAdj.Location = New-Object System.Drawing.Point(230, 10)
    $pBox.Controls.Add($lblAdj)
    $tbAdj = New-Object System.Windows.Forms.TextBox; $tbAdj.Text = $fills.Adj
    $tbAdj.BackColor = $cInput; $tbAdj.ForeColor = $cText; $tbAdj.BorderStyle = 'FixedSingle'
    $tbAdj.Location = New-Object System.Drawing.Point(230, 30); $tbAdj.Size = New-Object System.Drawing.Size(200, 24)
    $pBox.Controls.Add($tbAdj)
    $pJob.TBAdj = $tbAdj
    $pnlRight.Controls.Add($pBox)
    $y += 90

    $pnlFiles = New-Object System.Windows.Forms.Panel
    $pnlFiles.BackColor = clr "#111214"
    $pnlFiles.Location = New-Object System.Drawing.Point(10, $y)
    $pnlRight.Controls.Add($pnlFiles)
    $pJob.PnlFiles = $pnlFiles

    $files = Get-ChildItem -Path $parentPath -File | Sort-Object Name
    $fy = 0
    foreach ($fi in $files) {
        $parsed = ParseFile $fi.Name
        $fRow = New-Object System.Windows.Forms.Panel
        $fRow.BackColor = if (($pJob.FileRows.Count % 2) -eq 0) { clr "#16171B" } else { clr "#1A1C22" }
        $fRow.Size = New-Object System.Drawing.Size(600, 50); $fRow.Location = New-Object System.Drawing.Point(0, $fy)

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
        $lNew.Location = New-Object System.Drawing.Point(330, 15); $lNew.Size = New-Object System.Drawing.Size(260, 20)
        $fRow.Controls.Add($lNew)

        $pJob.FileRows.Add([PSCustomObject]@{ OldPath = $fi.FullName; SuffixBox = $sBadge; NewLbl = $lNew; Ext = $parsed.Extension })

        $sBadge.Tag = @{ P = $pJob; G = $gpJob }
        $sBadge.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })
        $pnlFiles.Controls.Add($fRow)
        $fy += 50
    }

    # Auto-size the height of the files list based entirely on contents
    $pnlFiles.Size = New-Object System.Drawing.Size(620, $fy)

    $tbChar.Tag = @{ P = $pJob; G = $gpJob }; $tbChar.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })
    $tbAdj.Tag = @{ P = $pJob; G = $gpJob };  $tbAdj.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })

    Update-ParentPreview $pJob $gpJob
    return $pJob
}

function Build-GpJob($gpPath, $parentDict) {
    $diGrand = if ($gpPath -notlike "ROOT_*") { [System.IO.DirectoryInfo]::new($gpPath) } else { $null }
    $gpName = if ($diGrand) { $diGrand.Name } else { "(No Parent Folder)" }

    $gpJob = @{
        GpPath = $gpPath; DiGrand = $diGrand
        Parents = [System.Collections.Generic.List[object]]::new()
        IsDone = $false
    }
    $script:jobs.Add($gpJob)

    $container = New-Object System.Windows.Forms.Panel
    $container.BackColor = clr "#1C1D23"; $container.BorderStyle = 'FixedSingle'
    $mainScroll.Controls.Add($container)
    $gpJob.Container = $container

    # Header
    $header = New-Object System.Windows.Forms.Panel
    $header.BackColor = clr "#2A2C35"; $header.Height = 60
    $container.Controls.Add($header)
    $gpJob.Header = $header

    $lblGP = New-Object System.Windows.Forms.Label
    $lblGP.Text = "Grandparent Theme:  $gpName"
    $lblGP.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblGP.ForeColor = $cAmber; $lblGP.AutoSize = $true
    $lblGP.Location = New-Object System.Drawing.Point(10, 20)
    $header.Controls.Add($lblGP)

    $tbTheme = New-Object System.Windows.Forms.TextBox
    $tbTheme.Text = $gpName; $tbTheme.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $tbTheme.BackColor = $cInput; $tbTheme.ForeColor = $cText; $tbTheme.BorderStyle = 'FixedSingle'
    $tbTheme.Location = New-Object System.Drawing.Point(220, 18); $tbTheme.Size = New-Object System.Drawing.Size(250, 24)
    $header.Controls.Add($tbTheme)
    $gpJob.TBTheme = $tbTheme

    $chkSkip = New-Object System.Windows.Forms.CheckBox
    $chkSkip.Text = "Don't rename folder"
    $chkSkip.Font = New-Object System.Drawing.Font("Segoe UI", 9); $chkSkip.ForeColor = $cText; $chkSkip.AutoSize = $true
    $chkSkip.Location = New-Object System.Drawing.Point(490, 20)
    $header.Controls.Add($chkSkip)
    $gpJob.ChkSkip = $chkSkip

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Remove from Queue"
    $btnRemove.BackColor = $cRed; $btnRemove.ForeColor = clr "#FFFFFF"
    $btnRemove.FlatStyle = 'Flat'; $btnRemove.FlatAppearance.BorderSize = 0
    $btnRemove.Size = New-Object System.Drawing.Size(160, 30)
    $btnRemove.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Right"
    $btnRemove.Location = New-Object System.Drawing.Point(1300, 15)
    $header.Controls.Add($btnRemove)
    $btnRemove.Tag = $gpJob
    $btnRemove.Add_Click({
        $j = $this.Tag
        $j.Container.Dispose()
        $script:jobs.Remove($j) | Out-Null
        $script:resizeTimer.Start()
    })

    # Parents Container
    $parentList = New-Object System.Windows.Forms.Panel
    $parentList.BackColor = clr "#1C1D23"
    $container.Controls.Add($parentList)
    $gpJob.ParentListPanel = $parentList

    foreach ($pKey in $parentDict.Keys) {
        $pJob = Build-PJob $pKey $parentDict[$pKey] $gpJob
        $gpJob.Parents.Add($pJob)
    }

    # Theme Sync Event
    $tbTheme.Tag = $gpJob
    $tbTheme.Add_TextChanged({
        foreach ($p in $this.Tag.Parents) {
            if ($p.LblThVal) { $p.LblThVal.Text = $this.Text }
            Update-ParentPreview $p $this.Tag
        }
    })

    # Footer
    $footer = New-Object System.Windows.Forms.Panel
    $footer.BackColor = clr "#2A2C35"; $footer.Height = 60
    $container.Controls.Add($footer)
    $gpJob.Footer = $footer

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Apply & Save Group"
    $btnApply.BackColor = clr "#4CAF72"; $btnApply.ForeColor = clr "#FFFFFF"
    $btnApply.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnApply.FlatStyle = 'Flat'; $btnApply.FlatAppearance.BorderSize = 0
    $btnApply.Size = New-Object System.Drawing.Size(200, 40)
    $btnApply.Location = New-Object System.Drawing.Point(10, 10)
    $footer.Controls.Add($btnApply)
    $gpJob.BtnApply = $btnApply

    $btnApply.Tag = $gpJob
    $btnApply.Add_Click({ Apply-GpJob $this.Tag })
}

# --- 8. Dynamic Loading & Form Events ---
$btnProcessAll.Add_Click({
    $btnProcessAll.Text = "Processing Queue..."; $btnProcessAll.Enabled = $false
    foreach ($gpJob in $script:jobs) {
        if (-not $gpJob.IsDone -and -not $gpJob.Container.IsDisposed) {
            Apply-GpJob $gpJob
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    $btnProcessAll.Text = "All Items Processed!"
})

$form.Add_Shown({
    $form.SuspendLayout()
    $idx = 1
    foreach ($gpPath in $gpQueue.Keys) {
        $lblGlobalTitle.Text = "Extracting & Analyzing Group $idx of $($gpQueue.Count)..."
        [System.Windows.Forms.Application]::DoEvents()
        Build-GpJob $gpPath $gpQueue[$gpPath]
        $idx++
    }
    $lblGlobalTitle.Text = "Queue Dashboard ($($gpQueue.Count) Group(s) found)"
    $btnProcessAll.Enabled = $true
    $form.ResumeLayout()
    $form.Width += 1
})

$script:resizeTimer = New-Object System.Windows.Forms.Timer
$script:resizeTimer.Interval = 100
$script:resizeTimer.Add_Tick({
    if ([System.Windows.Forms.Control]::MouseButtons -ne [System.Windows.Forms.MouseButtons]::None) {
        $script:resizeTimer.Start()
        return
    }

    $script:resizeTimer.Stop()
    $form.SuspendLayout()

    $savedScroll = $mainScroll.AutoScrollPosition
    $mainScroll.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)

    $availW = $mainScroll.ClientSize.Width - 680
    $sq = [math]::Min(($availW / 2), 512)
    if ($sq -lt 250) { $sq = 250 }
    $scale = $sq / 512.0

    $yOffset = 10
    foreach ($gpJob in $script:jobs) {
        if ($gpJob.Container.IsDisposed) { continue }

        $gpJob.Header.Width = $mainScroll.ClientSize.Width - 25
        $gpJob.Footer.Width = $mainScroll.ClientSize.Width - 25
        $gpJob.Footer.Controls[0].Left = $gpJob.Footer.Width - 220
        $gpJob.Header.Controls[3].Left = $gpJob.Header.Width - 180

        $pyOffset = 0
        foreach ($pJob in $gpJob.Parents) {
            # Calculate necessary height for the right panel based purely on file count
            $pnlFilesHeight = $pJob.FileRows.Count * 50
            $rightH = 215 + $pnlFilesHeight + 10 # 215 offsets all the textboxes above the file list
            $rowH = [math]::Max(($sq + 20), $rightH)

            $pJob.RowPanel.Location = New-Object System.Drawing.Point(0, $pyOffset)
            $pJob.RowPanel.Size = New-Object System.Drawing.Size(($mainScroll.ClientSize.Width - 25), $rowH)

            $pJob.Card.Location = New-Object System.Drawing.Point(10, 10)
            $pJob.Card.Size = New-Object System.Drawing.Size($sq, $sq)

            $pJob.PickCard.Location = New-Object System.Drawing.Point(($sq + 20), 10)
            $pJob.PickCard.Size = New-Object System.Drawing.Size($sq, $sq)

            $pJob.PnlRight.Location = New-Object System.Drawing.Point(($sq * 2 + 30), 10)
            $pJob.PnlRight.Size = New-Object System.Drawing.Size(640, $rightH)
            $pJob.PnlFiles.Size = New-Object System.Drawing.Size(620, $pnlFilesHeight)

            if ($pJob.BtnRemove) { $pJob.BtnRemove.Left = 640 - 130 }

            foreach ($item in $pJob.ScaleElements) {
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
            $pyOffset += $rowH + 10
        }

        $gpJob.ParentListPanel.Location = New-Object System.Drawing.Point(0, $gpJob.Header.Height)
        $gpJob.ParentListPanel.Size = New-Object System.Drawing.Size(($mainScroll.ClientSize.Width - 25), $pyOffset)

        $gpJob.Footer.Location = New-Object System.Drawing.Point(0, ($gpJob.Header.Height + $gpJob.ParentListPanel.Height))

        $gpContainerH = $gpJob.Header.Height + $gpJob.ParentListPanel.Height + $gpJob.Footer.Height
        $gpJob.Container.Location = New-Object System.Drawing.Point(10, $yOffset)
        $gpJob.Container.Size = New-Object System.Drawing.Size(($mainScroll.ClientSize.Width - 25), $gpContainerH)

        $yOffset += $gpContainerH + 20
    }

    $form.ResumeLayout()
    $mainScroll.AutoScrollPosition = New-Object System.Drawing.Point([math]::Abs($savedScroll.X), [math]::Abs($savedScroll.Y))
    $mainScroll.Invalidate()
})

$form.Add_Resize({ $script:resizeTimer.Stop(); $script:resizeTimer.Start() })

$form.Add_FormClosed({
    foreach ($gpJob in $script:jobs) {
        foreach ($pJob in $gpJob.Parents) {
            if (Test-Path $pJob.TempWork) { Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
})

[System.Windows.Forms.Application]::Run($form)