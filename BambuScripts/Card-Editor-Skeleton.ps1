Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- ROBUST FLAT ROUTING ENGINE ---
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = [System.Environment]::CurrentDirectory }
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"
# ----------------------------------

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
                    $colorMap[$key] = [System.Drawing.Color]::FromArgb($px.A, $rng.Next(20,256), $rng.Next(20,256), $rng.Next(20,256))
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
$btnProcessAll.Text = "Add All To Queue"
$btnProcessAll.BackColor = clr "#4CAF72"; $btnProcessAll.ForeColor = clr "#FFFFFF"
$btnProcessAll.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnProcessAll.FlatStyle = 'Flat'; $btnProcessAll.FlatAppearance.BorderSize = 0
$btnProcessAll.Size = New-Object System.Drawing.Size(220, 40)
$btnProcessAll.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Right"
$btnProcessAll.Location = New-Object System.Drawing.Point(1300, 10)
$btnProcessAll.Enabled = $false
$pnlTop.Controls.Add($btnProcessAll)

$btnCombineData = New-Object System.Windows.Forms.Button
$btnCombineData.Text = "Combine TSV Data"
$btnCombineData.BackColor = clr "#E8A135" # Amber color to distinguish it
$btnCombineData.ForeColor = clr "#FFFFFF"
$btnCombineData.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnCombineData.FlatStyle = 'Flat'; $btnCombineData.FlatAppearance.BorderSize = 0
$btnCombineData.Size = New-Object System.Drawing.Size(180, 40)
$btnCombineData.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Right"
$btnCombineData.Location = New-Object System.Drawing.Point(1100, 10)
$pnlTop.Controls.Add($btnCombineData)

$btnCombineData.Add_Click({
    $targetDirs = @()
    # Figure out what folders to roll up (Grandparents, or Parents if no Grandparent exists)
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
            $key = ($line -split "`t")
            $combined[$key] = $line
        }

        if ($combined.Count -gt 0) {
            $combined.Values | Set-Content -Path $outTsvPath -Encoding UTF8
            $combinedCount++
            # Store values for the clipboard
            foreach ($val in $combined.Values) { [void]$clipboardArray.Add($val) }
        }
    }

    if ($clipboardArray.Count -gt 0) {
        try {
            # Join all rows with a standard Windows newline and push to clipboard
            $clipboardText = $clipboardArray.ToArray() -join "`r`n"
            [System.Windows.Forms.Clipboard]::SetText($clipboardText)

            [System.Windows.Forms.MessageBox]::Show("Successfully combined TSV data for $combinedCount group(s).`n`nThe combined data ($($clipboardArray.Count) rows) has been copied to your clipboard!", "Combine Data Complete", 0, 64) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Combined TSV data for $combinedCount group(s), but failed to copy to clipboard. Another program may be locking it.", "Combine Data Partial Success", 0, 48) | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No new data found to combine.", "Combine Data Empty", 0, 64) | Out-Null
    }
})

$mainScroll = New-Object System.Windows.Forms.Panel
$mainScroll.Dock = 'Fill'; $mainScroll.AutoScroll = $true; $mainScroll.BackColor = clr "#0D0E10"
$form.Controls.Add($mainScroll)
$mainScroll.BringToFront()

# --- 5. Async Worker Engine Data ---
$script:jobs = New-Object System.Collections.ArrayList
$script:processQueue = New-Object System.Collections.Queue
$script:activeProcess = $null
$script:activeProcessJob = $null

function Add-ScaledElement($pJob, $targetPanel, $ctrl, $bx, $by, $bw, $bh, $bf) {
    $pJob.ScaleElements.Add([PSCustomObject]@{ Target = $targetPanel; Ctrl = $ctrl; X = $bx; Y = $by; W = $bw; H = $bh; Font = $bf }) | Out-Null
    $ctrl.AutoSize = $false
    $targetPanel.Controls.Add($ctrl)
}

function Validate-PJob($pJob) {
    if ($pJob.IsQueued -or $pJob.IsDone) { return }

    $colorsSafe = $true
    foreach ($slot in $pJob.UISlots) {
        if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $colorsSafe = $false }
    }

    if ($pJob.HasCollision) {
        $pJob.BtnApply.Text = "Name Collision!"
        $pJob.BtnApply.BackColor = clr "#D95F5F"
        $pJob.BtnApply.Enabled = $false
    } elseif (-not $colorsSafe) {
        $pJob.BtnApply.Text = "Unmatched Colors"
        $pJob.BtnApply.BackColor = clr "#E8A135"
        $pJob.BtnApply.Enabled = $false
    } else {
        $pJob.BtnApply.Text = "Add to Queue"
        $pJob.BtnApply.BackColor = clr "#4CAF72"
        $pJob.BtnApply.Enabled = $true
    }
}

function Update-ParentPreview($pJob, $gpJob) {
    # 1. Read values from edit boxes (Preserve original casing)
    $ch = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
    $ad = $pJob.TBAdj.Text -replace '[^a-zA-Z0-9]', ''
    $th = $gpJob.TBTheme.Text -replace '[^a-zA-Z0-9]', ''

    # 2. Update component labels on the Card (Top Right, ALL CAPS)
    $pJob.LblCharCard.Text = $ch.ToUpper()
    $pJob.LblAdjCard.Text  = if ($ad) { "($($ad.ToUpper()))" } else { "" }
    $pJob.LblThemeCard.Text = $th.ToUpper()

    # 3. Update the combined name label in the edit bar (Top Right)
    $pJob.LblFolderName.Text = "$ch $ad".Trim()

    # 4. Collision detection & Target Name building
    $nameCounts = @{}
    foreach ($r in $pJob.FileRows) {
        $sf = $r.SuffixBox.Text -replace '[^a-zA-Z0-9]', ''
        $parts = New-Object System.Collections.ArrayList
        if ($ch) { [void]$parts.Add($ch) }
        if ($ad) { [void]$parts.Add($ad) }
        if ($th) { [void]$parts.Add($th) }
        if ($sf) { [void]$parts.Add($sf) }
        $r.TargetName = ($parts.ToArray() -join '_') + $r.Ext
        if (-not $nameCounts.ContainsKey($r.TargetName)) { $nameCounts[$r.TargetName] = 0 }
        $nameCounts[$r.TargetName]++
    }

    $hasCollision = $false
    foreach ($r in $pJob.FileRows) {
        $r.NewLbl.Text = $r.TargetName
        if ($nameCounts[$r.TargetName] -gt 1) {
            $r.NewLbl.ForeColor = clr "#D95F5F"
            $hasCollision = $true
        } else { $r.NewLbl.ForeColor = clr "#4CAF72" }
    }
    $pJob.HasCollision = $hasCollision
    Validate-PJob $pJob
}

function Add-FileRow($pJob, $gpJob, $fi, $fy) {
    $parsed = ParseFile $fi.Name
    $fRow = New-Object System.Windows.Forms.Panel
    $fRow.BackColor = if (($pJob.FileRows.Count % 2) -eq 0) { clr "#16171B" } else { clr "#1A1C22" }
    $fRow.Size = New-Object System.Drawing.Size(600, 50)
    $fRow.Location = New-Object System.Drawing.Point(0, $fy)

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
    $lArr.Text = "->"; $lArr.ForeColor = $cMuted
    $lArr.Location = New-Object System.Drawing.Point(300, 15); $lArr.Size = New-Object System.Drawing.Size(25, 20)
    $fRow.Controls.Add($lArr)

    $lNew = New-Object System.Windows.Forms.Label
    $lNew.ForeColor = clr "#4CAF72"; $lNew.Font = New-Object System.Drawing.Font("Consolas", 8)
    $lNew.Location = New-Object System.Drawing.Point(330, 15); $lNew.Size = New-Object System.Drawing.Size(240, 20)
    $fRow.Controls.Add($lNew)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = "X"; $btnDel.BackColor = clr "#D95F5F"; $btnDel.ForeColor = clr "#FFFFFF"
    $btnDel.FlatStyle = 'Flat'; $btnDel.FlatAppearance.BorderSize = 0
    $btnDel.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
    $btnDel.Location = New-Object System.Drawing.Point(575, 15); $btnDel.Size = New-Object System.Drawing.Size(20, 20)
    $fRow.Controls.Add($btnDel)

    $frObj = [PSCustomObject]@{ OldPath = $fi.FullName; SuffixBox = $sBadge; NewLbl = $lNew; Ext = $parsed.Extension; TargetName = "" }
    $pJob.FileRows.Add($frObj) | Out-Null

    $sBadge.Tag = @{ P = $pJob; G = $gpJob }
    $sBadge.Add_TextChanged({ $t = $this.Tag; Update-ParentPreview $t.P $t.G })

    $btnDel.Tag = @{ P = $pJob; G = $gpJob; Row = $fRow; FileRow = $frObj; Path = $fi.FullName }
    $btnDel.Add_Click({
        $t = $this.Tag
        $res = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to permanently delete:`n$($t.Path | Split-Path -Leaf)?", "Confirm File Deletion", 4, 32)
        if ($res -eq 'Yes') {
            Remove-Item $t.Path -Force -ErrorAction SilentlyContinue
            $t.P.FileRows.Remove($t.FileRow) | Out-Null
            $t.Row.Dispose()

            $newY = 0
            foreach ($ctrl in $t.P.PnlFiles.Controls) {
                if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Visible) {
                    $ctrl.Location = New-Object System.Drawing.Point(0, $newY)
                    $newY += 50
                }
            }
            $t.P.PnlFiles.Size = New-Object System.Drawing.Size(620, $newY)
            Update-ParentPreview $t.P $t.G
            $script:resizeTimer.Start()
        }
    })

    $pJob.PnlFiles.Controls.Add($fRow)
}

function Refresh-PJob($pJob, $gpJob) {
    $pJob.ProcessingOverlay.Visible = $false
    $pJob.PickProcessingOverlay.Visible = $false

    # --- 1. RESET ANCHOR FILE TRACKING FIRST ---
    $pJob.ProcessedAnchorPath = ""
    $newAnchor = Get-ChildItem -Path $pJob.FolderPath -File | Where-Object { $_.Name -match '(?i)Full\.3mf$' -and $_.Name -notmatch '(?i)\.gcode\.3mf$' } | Select-Object -First 1
    if ($newAnchor) { $pJob.AnchorFile = $newAnchor }

    if ($pJob.PbPlateFinished) {
        if ($pJob.PbPlateFinished.Image) { $pJob.PbPlateFinished.Image.Dispose(); $pJob.PbPlateFinished.Image = $null }
        $pJob.PbPlateFinished.Visible = $false
    }

    # --- 2. RE-EXTRACT METADATA TO FIX INSTANT FAIL ---
    if (Test-Path $pJob.TempWork) {
        Remove-Item $pJob.TempWork -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Do NOT use New-Item here! Let the ZipFile class create the folder itself to prevent .NET crashes.
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pJob.AnchorFile.FullName, $pJob.TempWork)
    } catch {}

    # Keep the UI Folder Label perfectly in sync!
    if ($pJob.LblFolder) {
        $pJob.LblFolder.Text = "Folder: " + (Split-Path $pJob.FolderPath -Leaf)
    }
    # --------------------------------------------------

    $diParent = [System.IO.DirectoryInfo]::new($pJob.FolderPath)
    $customPng = Get-ChildItem -Path $pJob.FolderPath -Filter "*.png" | Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" } | Select-Object -First 1

    $gcodeFile = Get-ChildItem -Path $pJob.FolderPath -Filter "*Full.gcode.3mf" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    # 1. Update [CURRENT] Thumbnail (Gcode Image)
    $gcodeImgPath = $null
    if ($gcodeFile) {
        $extractedGcodePlate = Join-Path $pJob.TempWork "plate_1_gcode.png"
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $entry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
            if ($entry) {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $extractedGcodePlate, $true)
                $gcodeImgPath = $extractedGcodePlate
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    }

    if ($gcodeImgPath -and (Test-Path $gcodeImgPath)) {
        $fs = New-Object System.IO.FileStream($gcodeImgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        if ($pJob.PbCurrent.Image) { $pJob.PbCurrent.Image.Dispose() }
        $pJob.PbCurrent.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pJob.PbCurrent.Tag = @{ Path = $gcodeImgPath }
        $pJob.PbCurrent.Visible = $true
        $pJob.LblCurrent.Visible = $true
    } else {
        if ($pJob.PbCurrent.Image) { $pJob.PbCurrent.Image.Dispose(); $pJob.PbCurrent.Image = $null }
        $pJob.PbCurrent.Visible = $false
        $pJob.LblCurrent.Visible = $false
    }

    # 2. Update Editable Sketched Area (Custom Base PNG)

    if ($pJob.CustomImagePath -and (Test-Path $pJob.CustomImagePath)) {
        if ($pJob.PbPlate.Image) { $pJob.PbPlate.Image.Dispose() }
        $fs = New-Object System.IO.FileStream($pJob.CustomImagePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pJob.PbPlate.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pJob.PbPlate.Tag.ImgPath = $pJob.CustomImagePath
    } elseif ($customPng) {
        if ($pJob.PbPlate.Image) { $pJob.PbPlate.Image.Dispose() }
        $pJob.CustomImagePath = $customPng.FullName
        $fs = New-Object System.IO.FileStream($pJob.CustomImagePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pJob.PbPlate.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pJob.PbPlate.Tag.ImgPath = $pJob.CustomImagePath
    } else {
        if ($pJob.PbPlate.Image) { $pJob.PbPlate.Image.Dispose(); $pJob.PbPlate.Image = $null }
        $pJob.CustomImagePath = $null
    }

    $pJob.ImgStatusLbl.Text = $statusText
    $pJob.ImgStatusLbl.ForeColor = $statusColor
    $pJob.ImgStatusLbl.BringToFront()

    $pickPath = $null
    if ($gcodeFile) {
        $zip = $null
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

    if ($pickPath -and (Test-Path $pickPath)) {
        $fs = New-Object System.IO.FileStream($pickPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        if ($pJob.PbPick.Image) { $pJob.PbPick.Image.Dispose() }
        $pJob.PbPick.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pJob.PbPick.Tag.Path = $pickPath
        foreach ($ctrl in $pJob.PickCard.Controls) {
            if ($ctrl -is [System.Windows.Forms.Label] -and $ctrl.Text -eq "[NO GCODE]") { $ctrl.Visible = $false }
        }
    } else {
        if ($pJob.PbPick.Image) { $pJob.PbPick.Image.Dispose(); $pJob.PbPick.Image = $null }
        foreach ($ctrl in $pJob.PickCard.Controls) {
            if ($ctrl -is [System.Windows.Forms.Label] -and $ctrl.Text -eq "[NO GCODE]") { $ctrl.Visible = $true; $ctrl.BringToFront() }
        }
    }

    foreach ($ctrl in $pJob.PnlFiles.Controls) { $ctrl.Dispose() }
    $pJob.PnlFiles.Controls.Clear()
    $pJob.FileRows.Clear()

    $files = Get-ChildItem -Path $pJob.FolderPath -File | Sort-Object Name
    $fy = 0
    foreach ($fi in $files) {
        Add-FileRow $pJob $gpJob $fi $fy
        $fy += 50
    }

    $pJob.PnlFiles.Size = New-Object System.Drawing.Size(620, $fy)
    Update-ParentPreview $pJob $gpJob

    $pJob.BtnApply.Text = "Add to Queue"
    $pJob.BtnApply.BackColor = clr "#4CAF72"
    $pJob.BtnApply.Enabled = $true
    $pJob.BtnApply.Width = 150
    if ($pJob.BtnRevertDone) { $pJob.BtnRevertDone.Visible = $false }

    $pJob.RowPanel.Enabled = $true
    $pJob.RowPanel.BackColor = clr "#16171B"
    $pJob.IsDone = $false
    $pJob.IsQueued = $false

    $pJob.ChkMerge.Enabled = $true
    $pJob.ChkSlice.Enabled = $true
    $pJob.ChkExtract.Enabled = $true
    $pJob.ChkImage.Enabled = $true

    # --- ADD THESE 4 LINES TO RESTORE TASK STATE ---
    $pJob.ChkMerge.Checked = $true
    $pJob.ChkSlice.Checked = $true
    $pJob.ChkExtract.Checked = $true
    $pJob.ChkImage.Checked = $true
    # -----------------------------------------------

    $script:resizeTimer.Start()
}

function Enqueue-PJob($pJob, $gpJob) {
    if ($pJob.IsQueued) { return }
    if ($pJob.IsDone) { return }
    if ($pJob.HasCollision) { return }
    foreach ($slot in $pJob.UISlots) {
        if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { return }
    }

    $pJob.IsQueued = $true
    $pJob.BtnApply.Text = "Queued..."
    $pJob.BtnApply.BackColor = $cAmber
    $pJob.RowPanel.Enabled = $false

    $pJob.ProcessingOverlay.Text = "[ PREPARING ]"
    $pJob.ProcessingOverlay.BringToFront()
    $pJob.ProcessingOverlay.Visible = $true

    $pJob.PickProcessingOverlay.Text = "[ PREPARING ]"
    $pJob.PickProcessingOverlay.BringToFront()
    $pJob.PickProcessingOverlay.Visible = $true

    $script:processQueue.Enqueue(@{ PJob = $pJob; GpJob = $gpJob })
}

function Start-NextProcess {
    if ($script:activeProcess -ne $null -or $script:processQueue.Count -eq 0) { return }

    $jobWrapper = $script:processQueue.Dequeue()
    $pJob = $jobWrapper.PJob
    $gpJob = $jobWrapper.GpJob
    $script:activeProcessJob = $jobWrapper

    $pJob.BtnApply.Text = "Processing..."
    $pJob.IsQueued = $false

    $th       = $gpJob.TBTheme.Text -replace '[^a-zA-Z0-9]', ''
    $oldGrand = if ($gpJob.DiGrand) { $gpJob.DiGrand.FullName } else { "" }

    $colorsSafe = $true
    foreach ($slot in $pJob.UISlots) {
        if ($slot.StatusLbl.Text -eq "[UNMATCHED]") { $colorsSafe = $false }
    }
    if (-not $colorsSafe) {
        $pJob.ChkMerge.Checked = $false
        $pJob.ChkImage.Checked = $false
    }

    $allTextFiles  = Get-ChildItem -Path $pJob.TempWork -Recurse -File |
                     Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
    $modifiedFiles = New-Object System.Collections.ArrayList

    foreach ($file in $allTextFiles) {
        $content  = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
        $modified = $false
        foreach ($slot in $pJob.UISlots) {
            $selName = $slot.Combo.Text
            if ($LibraryColors.Contains($selName)) {
                $newHex = $LibraryColors[$selName].ToUpper()
                $oldHex = $slot.OldHex.ToUpper()

                # Strict 7-char and 9-char mapping to prevent regex bleeding
                $oldHex9 = if ($oldHex.Length -eq 7) { $oldHex + "FF" } else { $oldHex }
                $oldHex7 = $oldHex.Substring(0,7)
                $newHex9 = if ($newHex.Length -eq 7) { $newHex + "FF" } else { $newHex }
                $newHex7 = $newHex.Substring(0,7)

                # Prioritize replacing the full 9-char hex first, then the 7-char
                if ($content -match "(?i)$oldHex9") {
                    $content  = $content -ireplace [regex]::Escape($oldHex9), $newHex9
                    $modified = $true
                }
                if ($content -match "(?i)$oldHex7") {
                    $content  = $content -ireplace [regex]::Escape($oldHex7), $newHex7
                    $modified = $true
                }
            }
        }
        if ($modified) {
            [System.IO.File]::WriteAllText($file.FullName, $content, (New-Object System.Text.UTF8Encoding($false)))
            $modifiedFiles.Add($file) | Out-Null
        }
    }

    $anchorTargetName = ""
    $anchorFileRow    = $null
    $currentAnchorLocation = $pJob.AnchorFile.FullName # Fallback

    foreach ($r in $pJob.FileRows) {
        # Match by name instead of full path, since parent folders might have been renamed by previous jobs
        if ((Split-Path $r.OldPath -Leaf) -eq $pJob.AnchorFile.Name) {
            $anchorTargetName = $r.NewLbl.Text
            $anchorFileRow    = $r
            $currentAnchorLocation = $r.OldPath
            break
        }
    }
    if ($anchorTargetName -eq "") { $anchorTargetName = $pJob.AnchorFile.Name }

    $newFilePath = Join-Path $pJob.FolderPath $anchorTargetName

    if ($currentAnchorLocation -ne $newFilePath) {
        if (Test-Path $newFilePath) { Remove-Item $newFilePath -Force -ErrorAction SilentlyContinue }
        try { Rename-Item $currentAnchorLocation $anchorTargetName -Force } catch {}
    }

    if ($null -ne $anchorFileRow) { $anchorFileRow.OldPath = $newFilePath }

    if ($modifiedFiles.Count -gt 0) {
        $zip = [System.IO.Compression.ZipFile]::Open($newFilePath, 'Update')
        foreach ($file in $modifiedFiles) {
            $rel   = $file.FullName.Substring($pJob.TempWork.Length).TrimStart('\','/').Replace('\','/')
            $entry = $zip.GetEntry($rel)
            if ($entry) { $entry.Delete() }
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $rel) | Out-Null
        }
        $zip.Dispose()

        # CRITICAL FIX: Force garbage collection to release the ZipArchive lock
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    $pJob.ProcessedAnchorPath = $newFilePath
    [System.GC]::Collect()
    Start-Sleep -Milliseconds 100

    foreach ($r in $pJob.FileRows) {
        if ($r.OldPath -eq $pJob.ProcessedAnchorPath) { continue }
        $targetName = $r.NewLbl.Text
        $newPath    = Join-Path $pJob.FolderPath $targetName
        if ($r.OldPath -ne $newPath -and (Test-Path $r.OldPath)) {
            Rename-Item $r.OldPath $targetName -Force
            $r.OldPath = $newPath
        }
    }

    $cleanChar     = $pJob.TBChar.Text -replace '[^a-zA-Z0-9]', ''
    $cleanAdj      = $pJob.TBAdj.Text  -replace '[^a-zA-Z0-9]', ''
    $pParts        = New-Object System.Collections.ArrayList
    if ($cleanChar) { $pParts.Add($cleanChar) | Out-Null }
    if ($cleanAdj)  { $pParts.Add($cleanAdj)  | Out-Null }
    if ($th)        { $pParts.Add($th)         | Out-Null }
    $newParentName = $pParts.ToArray() -join '_'

    $oldFolder = $pJob.FolderPath
    if ($newParentName -ne '' -and $newParentName -ne (Split-Path $oldFolder -Leaf)) {
        $newFolder = Join-Path (Split-Path $oldFolder -Parent) $newParentName
        try {
            Rename-Item $oldFolder $newParentName -Force -ErrorAction Stop
            $pJob.FolderPath          = $newFolder
            $pJob.ProcessedAnchorPath = $pJob.ProcessedAnchorPath.Replace($oldFolder, $newFolder)
            if ($pJob.CustomImagePath) { $pJob.CustomImagePath = $pJob.CustomImagePath.Replace($oldFolder, $newFolder) }
            foreach ($r in $pJob.FileRows) { $r.OldPath = $r.OldPath.Replace($oldFolder, $newFolder) }
        } catch {}
    }

    if (-not $gpJob.ChkSkip.Checked -and $th -ne '' -and $oldGrand -ne '' -and
        $th -ne (Split-Path $oldGrand -Leaf)) {
        $newGrand = Join-Path (Split-Path $oldGrand -Parent) $th
        try {
            Rename-Item $oldGrand $th -Force -ErrorAction Stop
            $gpJob.GpPath  = $newGrand
            $gpJob.DiGrand = [System.IO.DirectoryInfo]::new($newGrand)

            foreach ($p in $gpJob.Parents) {
                $p.FolderPath = $p.FolderPath.Replace($oldGrand, $newGrand)
                if ($p.ProcessedAnchorPath) {
                    $p.ProcessedAnchorPath = $p.ProcessedAnchorPath.Replace($oldGrand, $newGrand)
                }
                if ($p.CustomImagePath) {
                    $p.CustomImagePath = $p.CustomImagePath.Replace($oldGrand, $newGrand)
                }
                foreach ($fr in $p.FileRows) { $fr.OldPath = $fr.OldPath.Replace($oldGrand, $newGrand) }
            }

            foreach ($otherGp in $script:jobs) {
                if ($otherGp -ne $gpJob -and $otherGp.GpPath.StartsWith($oldGrand)) {
                    $otherGp.GpPath  = $otherGp.GpPath.Replace($oldGrand, $newGrand)
                    $otherGp.DiGrand = [System.IO.DirectoryInfo]::new($otherGp.GpPath)
                    foreach ($otherP in $otherGp.Parents) {
                        $otherP.FolderPath = $otherP.FolderPath.Replace($oldGrand, $newGrand)
                        if ($otherP.ProcessedAnchorPath) {
                            $otherP.ProcessedAnchorPath = $otherP.ProcessedAnchorPath.Replace($oldGrand, $newGrand)
                        }
                        foreach ($fr in $otherP.FileRows) { $fr.OldPath = $fr.OldPath.Replace($oldGrand, $newGrand) }
                    }
                }
            }
        } catch {}
    }

    $workerScript = Join-Path $env:TEMP "AsyncWorker_$([guid]::NewGuid().ToString().Substring(0,8)).ps1"
    $sb           = New-Object System.Text.StringBuilder

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pJob.ProcessedAnchorPath)
    if ($baseName.ToLower().EndsWith("full")) {
        $basePrefix = $baseName.Substring(0, $baseName.Length - 4)
    } else {
        $basePrefix = $baseName + "_"
    }

    $dir        = $pJob.FolderPath
    $statusFile = Join-Path $dir "AsyncWorker_Status.txt"

    # NOTE: ALL paths written to the builder must evaluate $dir NOW (use "$dir")
    # instead of passing the literal string `$dir` which evaluates to $null in the background process.

    [void]$sb.AppendLine("`$ErrorActionPreference = 'Continue'")
    [void]$sb.AppendLine("Add-Type -AssemblyName System.IO.Compression.FileSystem")

    $doMerge   = $pJob.ChkMerge.Checked
    $doSlice   = $pJob.ChkSlice.Checked
    $doExtract = $pJob.ChkExtract.Checked
    $doImage   = $pJob.ChkImage.Checked

    $anchorPath = $pJob.ProcessedAnchorPath
    $nestPath   = Join-Path $dir "$($basePrefix)Nest.3mf"
    $finalPath  = Join-Path $dir "$($basePrefix)Final.3mf"
    $tempOut    = Join-Path $dir "$($baseName)_merged_temp.3mf"
    $tempIso    = Join-Path $env:TEMP "iso_$([guid]::NewGuid().ToString().Substring(0,8))"
    $slicedFile = Join-Path $dir "$($baseName).gcode.3mf"
    $singleFile = Join-Path $dir "$($basePrefix)Final.gcode.3mf"

    # TSV name strips _Full so it saves as Character_Adj_Theme_Data.tsv
    $tsvBaseName = $baseName -replace '(?i)_Full$', ''
    $tsvFile     = Join-Path $dir "${tsvBaseName}_Data.tsv"

    if ($doMerge) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value `"MERGING...`" -Force")
        [void]$sb.AppendLine("& `"$scriptDir\merge_3mf_worker.ps1`" -WorkDir `"$($pJob.TempWork)`" -InputPath `"$anchorPath`" -OutputPath `"$tempOut`" -DoColors `"0`"")
        [void]$sb.AppendLine("if (Test-Path `"$tempOut`") {")
        [void]$sb.AppendLine("    if (Test-Path `"$nestPath`") { Remove-Item `"$nestPath`" -Force }")
        [void]$sb.AppendLine("    Rename-Item -Path `"$anchorPath`" -NewName `"$($basePrefix)Nest.3mf`" -Force")
        [void]$sb.AppendLine("    Rename-Item -Path `"$tempOut`"    -NewName `"$($baseName).3mf`"      -Force")
        [void]$sb.AppendLine("    Get-ChildItem -Path `"$dir`" -Filter `"*MergeReport*.txt`" -ErrorAction SilentlyContinue | ForEach-Object { Remove-Item `$_.FullName -Force -ErrorAction SilentlyContinue }")
        [void]$sb.AppendLine("} else { Write-Host '[ERROR] Merge produced no output - aborting.' -ForegroundColor Red; exit 1 }")

        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value `"ISOLATING FINAL...`" -Force")
        [void]$sb.AppendLine("New-Item -ItemType Directory -Path `"$tempIso`" -Force | Out-Null")
        [void]$sb.AppendLine("[System.IO.Compression.ZipFile]::ExtractToDirectory(`"$nestPath`", `"$tempIso`")")
        [void]$sb.AppendLine("& `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `"$tempIso`" -OutputPath `"$finalPath`"")
        [void]$sb.AppendLine("Remove-Item `"$tempIso`" -Recurse -Force -ErrorAction SilentlyContinue")
    }

    if ($doSlice) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value `"SLICING $($baseName)...`" -Force")
        [void]$sb.AppendLine("Start-Sleep -Seconds 3") # Let the Slicer CLI breathe between files
        [void]$sb.AppendLine("& `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$anchorPath`" -IsolatedPath `"$finalPath`"")
    } elseif ($doExtract -or $doImage) {
        # Re-slice Final for skip-time data. If Final.3mf is missing, isolate it from Nest first.
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value `"RE-SLICING FINAL FOR DATA...`" -Force")
        [void]$sb.AppendLine("if (-not (Test-Path `"$finalPath`") -and (Test-Path `"$nestPath`")) {")
        [void]$sb.AppendLine("    `$tempIsoR = Join-Path `$env:TEMP `"iso_reslice_$([guid]::NewGuid().ToString().Substring(0,8))`"")
        [void]$sb.AppendLine("    New-Item -ItemType Directory -Path `$tempIsoR -Force | Out-Null")
        [void]$sb.AppendLine("    [System.IO.Compression.ZipFile]::ExtractToDirectory(`"$nestPath`", `$tempIsoR)")
        [void]$sb.AppendLine("    & `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `$tempIsoR -OutputPath `"$finalPath`"")
        [void]$sb.AppendLine("    Remove-Item `$tempIsoR -Recurse -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("}")
        [void]$sb.AppendLine("if (Test-Path `"$finalPath`") {")
        [void]$sb.AppendLine("    Start-Sleep -Seconds 3")
        [void]$sb.AppendLine("    & `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$finalPath`"")
        [void]$sb.AppendLine("}")
    }

    if ($doExtract -or $doImage) {
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value `"EXTRACTING DATA...`" -Force")
        $imageFlag = if ($doImage) { "-GenerateImage" } else { "" }
        [void]$sb.AppendLine("if (Test-Path `"$slicedFile`") {")
        [void]$sb.AppendLine("    Remove-Item `"$tsvFile`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("    & `"$scriptDir\Extract-3MFData.ps1`" -InputFile `"$slicedFile`" -SingleFile `"$singleFile`" -IndividualTsvPath `"$tsvFile`" $imageFlag")
        [void]$sb.AppendLine("    Remove-Item `"$singleFile`" -Force -ErrorAction SilentlyContinue")
        [void]$sb.AppendLine("} else {")
        [void]$sb.AppendLine("    Set-Content -Path `"$statusFile`" -Value `"[ERROR] SLICE FAILED - MISSING GCODE`" -Force")
        [void]$sb.AppendLine("    Start-Sleep -Seconds 4")
        [void]$sb.AppendLine("}")
    }

    if ($doImage) {
        # ReplaceImageNew.bat handles Synology cloud stubs with retry logic.
        # It finds *Full.gcode.3mf in the folder and injects the matching _slicePreview.png.
        [void]$sb.AppendLine("Set-Content -Path `"$statusFile`" -Value `"IMAGE INJECTION...`" -Force")
        [void]$sb.AppendLine("`$batPath = Join-Path `"$scriptDir`" `"ReplaceImageNew.bat`"")
        [void]$sb.AppendLine("if (Test-Path `$batPath) {")
        [void]$sb.AppendLine("    `$argList = '/c `"`"' + `$batPath + '`" `"' + `"$dir`" + '`"`"'")
        [void]$sb.AppendLine("    Start-Process -FilePath `"cmd.exe`" -ArgumentList `$argList -Wait -WindowStyle Hidden")
        [void]$sb.AppendLine("}")
    }

    [void]$sb.AppendLine("Get-ChildItem -Path `"$dir`" -Filter `"*ProcessLog*.txt`" -ErrorAction SilentlyContinue | Remove-Item -Force")
    [void]$sb.AppendLine("Remove-Item `"$statusFile`" -Force -ErrorAction SilentlyContinue")
    [void]$sb.AppendLine("Remove-Item `"$($pJob.TempWork)`" -Recurse -Force -ErrorAction SilentlyContinue")

    Set-Content -Path $workerScript -Value $sb.ToString()
    $script:activeProcess = Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$workerScript`"" -PassThru -WindowStyle Hidden
}

$script:queueTimer = New-Object System.Windows.Forms.Timer
$script:queueTimer.Interval = 500
$script:queueTimer.Add_Tick({
    if ($script:activeProcess -ne $null) {
        if (-not $script:activeProcess.HasExited) {
            $wrapper = $script:activeProcessJob
            $pJob = $wrapper.PJob
            $statusFile = Join-Path $pJob.FolderPath "AsyncWorker_Status.txt"
            if (Test-Path $statusFile) {
                try {
                    $fs = [System.IO.File]::Open($statusFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $sr = New-Object System.IO.StreamReader($fs)
                    $statusText = $sr.ReadToEnd()
                    $sr.Dispose()
                    $fs.Dispose()

                    if ($statusText) {
                        $pJob.ProcessingOverlay.Text = "[ $($statusText.Trim()) ]"
                        $pJob.PickProcessingOverlay.Text = "[ $($statusText.Trim()) ]"
                    }
                } catch {}
            }
        } else {
            $wrapper = $script:activeProcessJob
            $pJob = $wrapper.PJob
            $gpJob = $wrapper.GpJob

            $dir = $pJob.FolderPath
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pJob.ProcessedAnchorPath)

            $gcodeFile = Join-Path $dir "$($baseName).gcode.3mf"
            if (-not (Test-Path $gcodeFile)) {
                $gcodeFile = $pJob.ProcessedAnchorPath
            }

            $tempExtract = Join-Path $env:TEMP "finalRead_$([guid]::NewGuid().ToString().Substring(0,8))"
            New-Item -ItemType Directory -Path $tempExtract | Out-Null

            if (Test-Path $gcodeFile) {
                $retry = 0
                $zip = $null
                while ($retry -lt 10 -and $null -eq $zip) {
                    try {
                        $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile)
                    } catch {
                        $retry++
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 200
                    }
                }

                if ($null -ne $zip) {
                    try {
                        $plateEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
                        if ($plateEntry) {
                            $newPlate = Join-Path $tempExtract "plate_1.png"
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($plateEntry, $newPlate, $true)
                            $fs = New-Object System.IO.FileStream($newPlate, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)

                            if ($pJob.PbPlateFinished.Image) { $pJob.PbPlateFinished.Image.Dispose() }
                            $pJob.PbPlateFinished.Image = [System.Drawing.Image]::FromStream($fs)

                            if ($pJob.PbPlate.Image) { $pJob.PbPlate.Image.Dispose() }
                            $pJob.PbPlate.Image = [System.Drawing.Image]::FromStream($fs)

                            $fs.Close()

                            $pJob.PbPlateFinished.Tag = @{ ImgPath = $newPlate }
                            $pJob.PbPlateFinished.Visible = $true
                            $pJob.PbPlateFinished.BringToFront()

                            $pJob.ImgStatusLbl.Text = "[COMPILED]"
                            $pJob.ImgStatusLbl.ForeColor = clr "#4CAF72"
                            $pJob.ImgStatusLbl.BringToFront()
                        }

                        $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)(^|/)pick_1\.png$" } | Select-Object -First 1
                        if ($pickEntry) {
                            $rawPickPath = Join-Path $tempExtract "pick_1_raw.png"
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $rawPickPath, $true)
                            $newPick = Join-Path $tempExtract "pick_1.png"
                            Invoke-RandomizePickColors $rawPickPath $newPick | Out-Null
                            if (-not (Test-Path $newPick)) { $newPick = $rawPickPath }

                            $fs = New-Object System.IO.FileStream($newPick, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                            if ($pJob.PbPick.Image) { $pJob.PbPick.Image.Dispose() }
                            $pJob.PbPick.Image = [System.Drawing.Image]::FromStream($fs)
                            $fs.Close()
                            $pJob.PbPick.Tag.Path = $newPick

                            foreach ($ctrl in $pJob.PickCard.Controls) {
                                if ($ctrl -is [System.Windows.Forms.Label] -and $ctrl.Text -eq "[NO GCODE]") {
                                    $ctrl.Visible = $false
                                }
                            }
                        }

                        $modSetPath = Join-Path $pJob.TempWork "Metadata\model_settings.config"
                        if (Test-Path $modSetPath) {
                            $VerificationSlots = New-Object System.Collections.Generic.Dictionary[string, string]
                            try {
                                [xml]$modXml = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
                                foreach ($node in $modXml.SelectNodes('//metadata[contains(@key, "filament_colou?r")]')) {
                                    $val = $node.GetAttribute('value').ToUpper()
                                    if (-not [string]::IsNullOrWhiteSpace($val)) {
                                        $VerificationSlots[$val] = $true
                                    }
                                }
                            } catch {}
                        }

                        foreach ($slot in $pJob.UISlots) {
                            $selectedName = $slot.Combo.Text
                            if ($LibraryColors.Contains($selectedName)) {
                                $verifiedHex = $LibraryColors[$selectedName]
                                $slot.Swatch.BackColor = clr $verifiedHex
                                $slot.Swatch.Invalidate()
                                $slot.StatusLbl.Text = "[VERIFIED]"
                                $slot.StatusLbl.ForeColor = clr "#4CAF72"
                            }
                        }
                    } finally {
                        $zip.Dispose()
                    }
                }
            }

            foreach ($ctrl in $pJob.PnlFiles.Controls) { $ctrl.Dispose() }
            $pJob.PnlFiles.Controls.Clear()
            $pJob.FileRows.Clear()

            $files = Get-ChildItem -Path $pJob.FolderPath -File | Sort-Object Name
            $fy = 0
            foreach ($fi in $files) {
                Add-FileRow $pJob $gpJob $fi $fy
                $fy += 50
            }
            $pJob.PnlFiles.Size = New-Object System.Drawing.Size(620, $fy)
            Update-ParentPreview $pJob $gpJob

            $pJob.ProcessingOverlay.Visible = $false
            $pJob.PickProcessingOverlay.Visible = $false
            $pJob.RowPanel.Enabled = $true
            $pJob.RowPanel.BackColor = clr "#16171B"
            $pJob.IsDone = $true
            $pJob.ChkMerge.Enabled = $true
            $pJob.ChkSlice.Enabled = $true
            $pJob.ChkExtract.Enabled = $true
            $pJob.ChkImage.Enabled = $true

            $pJob.BtnApply.Text = "KEEP"
            $pJob.BtnApply.BackColor = clr "#4CAF72"
            $pJob.BtnApply.Enabled = $true
            $pJob.BtnApply.Width = 70
            if ($pJob.BtnRevertDone) { $pJob.BtnRevertDone.Visible = $true }

            $script:resizeTimer.Start()
            $script:activeProcess = $null
            $script:activeProcessJob = $null
        }
    } else {
        if ($script:processQueue.Count -gt 0) { Start-NextProcess }
    }
})

function Build-PJob($parentPath, $anchorFile, $gpJob) {
    $tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -ItemType Directory -Path $tempWork | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($anchorFile.FullName, $tempWork)

    $activeSlots = New-Object System.Collections.ArrayList
    $projPath   = Join-Path $tempWork "Metadata\project_settings.config"
    $modSetPath = Join-Path $tempWork "Metadata\model_settings.config"

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

    # 1. Check model_settings.config for extruder overrides
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
            foreach ($m in $extMatches) { $UsedSlots.Add($m.Groups.Value) | Out-Null }
        }
    }

    # 2. Check 3dmodel.model for native object colors (CRITICAL FIX)
    $modelFile = (Get-ChildItem -Path $tempWork -Filter '3dmodel.model' -Recurse | Select-Object -First 1)
    if ($modelFile -and (Test-Path $modelFile.FullName)) {
        try {
            $modelContent = [System.IO.File]::ReadAllText($modelFile.FullName, [System.Text.Encoding]::UTF8)
            $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
            foreach ($m in $matMatches) { $UsedSlots.Add($m.Groups.Value) | Out-Null }
        } catch {}
    }

    foreach ($hex in $SlotMap.Keys) {
        $slotId = $SlotMap[$hex]
        if ($UsedSlots.Contains($slotId)) {
            $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
            $matchedName = if ($HexToName.Contains($checkHex)) { $HexToName[$checkHex] } else { "" }
            $activeSlots.Add([PSCustomObject]@{ OldHex = $checkHex; Name = $matchedName; Grams = 0 }) | Out-Null
        }
    }
    if ($activeSlots.Count -gt 4) { $activeSlots = $activeSlots[0..3] }

    $pJob = @{
        FolderPath = $parentPath; AnchorFile = $anchorFile; TempWork = $tempWork
        ProcessedAnchorPath = ""
        CustomImagePath = $null
        UISlots = New-Object System.Collections.ArrayList
        FileRows = New-Object System.Collections.ArrayList
        ScaleElements = New-Object System.Collections.ArrayList
        IsDone = $false
        IsQueued = $false
        HasCollision = $false
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

    $overlay = New-Object System.Windows.Forms.Label
    $overlay.Text = "PROCESSING IN BACKGROUND..."
    $overlay.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $overlay.BackColor = [System.Drawing.Color]::FromArgb(220, 232, 161, 53)
    $overlay.ForeColor = [System.Drawing.Color]::Black
    $overlay.TextAlign = 'MiddleCenter'
    $overlay.Dock = 'Fill'
    $overlay.Visible = $false
    $card.Controls.Add($overlay)
    $pJob.ProcessingOverlay = $overlay

    $pickOverlay = New-Object System.Windows.Forms.Label
    $pickOverlay.Text = "PROCESSING IN BACKGROUND..."
    $pickOverlay.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $pickOverlay.BackColor = [System.Drawing.Color]::FromArgb(220, 232, 161, 53)
    $pickOverlay.ForeColor = [System.Drawing.Color]::Black
    $pickOverlay.TextAlign = 'MiddleCenter'
    $pickOverlay.Dock = 'Fill'
    $pickOverlay.Visible = $false
    $pickCard.Controls.Add($pickOverlay)
    $pJob.PickProcessingOverlay = $pickOverlay

    # --- ADD THE NEW [CURRENT] THUMBNAIL UI (TOP LEFT) ---
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "[CURRENT]"
    $lblCurrent.ForeColor = $cAmber; $lblCurrent.BackColor = $cBG
    $lblCurrent.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lblCurrent.TextAlign = 'TopLeft'
    Add-ScaledElement $pJob $card $lblCurrent 10 10 110 15 8
    $pJob.LblCurrent = $lblCurrent

    $pbCurrent = New-Object System.Windows.Forms.PictureBox
    $pbCurrent.SizeMode = 'Zoom'; $pbCurrent.BackColor = clr "#0D0E10"
    $pbCurrent.BorderStyle = 'FixedSingle'; $pbCurrent.Cursor = 'Hand'
    Add-ScaledElement $pJob $card $pbCurrent 10 25 110 110 0
    $pJob.PbCurrent = $pbCurrent
    $pbCurrent.Add_DoubleClick({ $t = $this.Tag; if ($t.Path -and (Test-Path $t.Path)) { Show-ImageViewer $t.Path "Current Compiled Image" } })
    # -----------------------------------------------------

    $pbPlateFinished = New-Object System.Windows.Forms.PictureBox
    $pbPlateFinished.SizeMode = 'Zoom'; $pbPlateFinished.BackColor = $cBG; $pbPlateFinished.Visible = $false
    Add-ScaledElement $pJob $card $pbPlateFinished 0 0 512 512 0
    $pJob.PbPlateFinished = $pbPlateFinished
    $pbPlateFinished.Add_DoubleClick({ $t = $this.Tag; if ($t.ImgPath -and (Test-Path $t.ImgPath)) { Show-ImageViewer $t.ImgPath "Final Card Preview" } })

    # --- EXTRACT AND LOAD IMAGES ---
    $diParent = [System.IO.DirectoryInfo]::new($parentPath)
    $customPng = Get-ChildItem -Path $parentPath -Filter "*.png" | Where-Object { $_.Name -notmatch "(?i)_slicePreview\.png" } | Select-Object -First 1
    $gcodeFile = Get-ChildItem -Path $parentPath -Filter "*Full.gcode.3mf" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    # 1. Update [CURRENT] Thumbnail (Gcode Image)
    $gcodeImgPath = $null
    if ($gcodeFile) {
        $extractedGcodePlate = Join-Path $tempWork "plate_1_gcode.png"
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodeFile.FullName)
            $entry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
            if ($entry) {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $extractedGcodePlate, $true)
                $gcodeImgPath = $extractedGcodePlate
            }
        } catch {} finally { if ($null -ne $zip) { $zip.Dispose() } }
    }

    if ($gcodeImgPath -and (Test-Path $gcodeImgPath)) {
        $fs = New-Object System.IO.FileStream($gcodeImgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pbCurrent.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pbCurrent.Tag = @{ Path = $gcodeImgPath }
        $pbCurrent.Visible = $true
        $lblCurrent.Visible = $true
    } else {
        $pbCurrent.Visible = $false
        $lblCurrent.Visible = $false
    }

    # 2. Update Editable Sketched Area (Custom Base PNG or Default)

    $pbModel = New-Object System.Windows.Forms.PictureBox
    $pbModel.SizeMode = 'Zoom'; $pbModel.BackColor = $cBG; $pbModel.AllowDrop = $true
    Add-ScaledElement $pJob $card $pbModel 10 80 250 360 0
    $pJob.PbPlate = $pbModel

    if ($customPng) {
        $pJob.CustomImagePath = $customPng.FullName
        $fs = New-Object System.IO.FileStream($pJob.CustomImagePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $pbModel.Image = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()
        $pbModel.Tag = @{ P = $pJob; ImgPath = $pJob.CustomImagePath }
        $statusText = "[CUSTOM SKETCH]"
        $statusColor = clr "#4CAF72"
    } else {
        $baseImgPath = Join-Path $tempWork "Metadata\plate_1.png"
        if (Test-Path $baseImgPath) {
            $fs = New-Object System.IO.FileStream($baseImgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $pbModel.Image = [System.Drawing.Image]::FromStream($fs)
            $fs.Close()
            $pbModel.Tag = @{ P = $pJob; ImgPath = $baseImgPath }
        }
    }
    $pbModel.Add_DoubleClick({ $t = $this.Tag; if ($t.ImgPath -and (Test-Path $t.ImgPath)) { Show-ImageViewer $t.ImgPath "Plate Preview" } })

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
            $p.ImgStatusLbl.Text = "[CUSTOM BASE]"
            $p.ImgStatusLbl.ForeColor = clr "#4CAF72"
            $p.CustomImagePath = $dest
        }
    })

    $lblThemeCard = New-Object System.Windows.Forms.Label
    $lblThemeCard.BackColor = $cBG; $lblThemeCard.ForeColor = clr "#808080"; $lblThemeCard.TextAlign = 'TopRight'
    Add-ScaledElement $pJob $card $lblThemeCard 320 50 180 18 10
    $pJob.LblThemeCard = $lblThemeCard

    $lblCharCard = New-Object System.Windows.Forms.Label
    $lblCharCard.BackColor = $cBG; $lblCharCard.ForeColor = $cAccent; $lblCharCard.TextAlign = 'TopRight'
    Add-ScaledElement $pJob $card $lblCharCard 320 15 180 40 18
    $pJob.LblCharCard = $lblCharCard

    $lblAdjCard = New-Object System.Windows.Forms.Label
    $lblAdjCard.BackColor = $cBG; $lblAdjCard.ForeColor = $cText; $lblAdjCard.TextAlign = 'TopRight'
    Add-ScaledElement $pJob $card $lblAdjCard 60 15 250 40 18
    $pJob.LblAdjCard = $lblAdjCard

    $lblThemeCard = New-Object System.Windows.Forms.Label
    $lblThemeCard.BackColor = $cBG; $lblThemeCard.ForeColor = clr "#808080"; $lblThemeCard.TextAlign = 'TopRight'
    Add-ScaledElement $pJob $card $lblThemeCard 320 50 180 18 10
    $pJob.LblThemeCard = $lblThemeCard

    $lblSkipTime = New-Object System.Windows.Forms.Label
    $lblSkipTime.Text = "Skip Time: -- min"
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

        $combo.Tag = @{ Swatch = $swatch; StatusLbl = $lblStatus; OrigName = $slotData.Name; P = $pJob }
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
            Validate-PJob $data.P
        }.GetNewClosure())
        Add-ScaledElement $pJob $card $combo ($boxX - 180) ($startY + 15) 170 25 10

        $lblGrams = New-Object System.Windows.Forms.Label
        $lblGrams.Text = "$($slotData.Grams) g"; $lblGrams.ForeColor = $cGrayTxt; $lblGrams.TextAlign = 'TopRight'
        Add-ScaledElement $pJob $card $lblGrams ($boxX - 90) ($startY + 45) 80 25 10

        $pJob.UISlots.Add([PSCustomObject]@{ OldHex = $slotData.OldHex; Combo = $combo; StatusLbl = $lblStatus; Swatch = $swatch }) | Out-Null
        $startY += $slotSpacing
    }
    $pbModel.SendToBack()

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
    $pJob.PbPick = $pbPick

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
    $pJob.LblFolder = $lblFolder

    $btnRefreshP = New-Object System.Windows.Forms.Button
    $btnRefreshP.Text = "Refresh"
    $btnRefreshP.BackColor = clr "#2A2C35"; $btnRefreshP.ForeColor = clr "#FFFFFF"
    $btnRefreshP.FlatStyle = 'Flat'; $btnRefreshP.FlatAppearance.BorderSize = 0
    $btnRefreshP.Size = New-Object System.Drawing.Size(100, 25)
    $btnRefreshP.Location = New-Object System.Drawing.Point(390, $y)
    $pnlRight.Controls.Add($btnRefreshP)
    $btnRefreshP.Tag = @{ P = $pJob; G = $gpJob }
    $btnRefreshP.Add_Click({
        $t = $this.Tag
        Refresh-PJob $t.P $t.G
    })
    $pJob.BtnRefresh = $btnRefreshP

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

    $tasksData = @{ Merge = $chkMerge; Slice = $chkSlice; Extract = $chkExtract; Image = $chkImage; PJob = $pJob; GpJob = $gpJob }
    $chkSlice.Tag = $tasksData
    $chkSlice.Add_CheckedChanged({ if ($this.Checked) { $this.Tag.Extract.Checked = $true } })

    $chkImage.Tag = $tasksData
    $chkImage.Add_CheckedChanged({
        if ($this.Checked) {
            $t = $this.Tag
            $pj = $t.PJob

            # Check if any TSV data file already exists in this folder
            $tsvExists = (Get-ChildItem -Path $pj.FolderPath -Filter "*_Data.tsv" -ErrorAction SilentlyContinue).Count -gt 0

            $t.Extract.Checked = $true

            # Only force a slice if we don't have a TSV to read from!
            if (-not $tsvExists) {
                $t.Slice.Checked = $true
            }
        }
    })

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
        $pj = $t.PJob
        $gp = $t.GpJob

        $batPath = Join-Path $scriptDir "RevertMerge.bat"
        if (Test-Path $batPath) {
            $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }

            $pj.BtnApply.Text = "Reverting..."
            $pj.BtnApply.BackColor = clr "#E8A135"
            $pj.RowPanel.Enabled = $false

            $t.Merge.Checked = $false; $t.Slice.Checked = $false; $t.Extract.Checked = $false; $t.Image.Checked = $false
            $t.Merge.Enabled = $false; $t.Slice.Enabled = $false; $t.Extract.Enabled = $false; $t.Image.Enabled = $false

            [System.Windows.Forms.Application]::DoEvents()

            try {
                $argList = '/c ""' + $batPath + '" "' + $targetPath + '""'
                $proc = Start-Process -FilePath "cmd.exe" -ArgumentList $argList -WindowStyle Hidden -PassThru
                $timeout = 100
                while (-not $proc.HasExited -and $timeout -gt 0) {
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 100
                    $timeout--
                }
                if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
            } catch {}

            Refresh-PJob $pj $gp

            $pj.ImgStatusLbl.Text = "[REVERTED]"
            $pj.ImgStatusLbl.ForeColor = clr "#D95F5F"
            $pj.RowPanel.BackColor = clr "#2D1C1C"
            $pj.BtnApply.Text = "Reverted"
            $pj.BtnApply.Enabled = $false
            $pj.RowPanel.Enabled = $true
        }
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
        Add-FileRow $pJob $gpJob $fi $fy
        $fy += 50
    }

    $pnlFiles.Size = New-Object System.Drawing.Size(620, $fy)

    $btnApplyP = New-Object System.Windows.Forms.Button
    $btnApplyP.Text = "Add to Queue"
    $btnApplyP.BackColor = clr "#4CAF72"; $btnApplyP.ForeColor = clr "#FFFFFF"
    $btnApplyP.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnApplyP.FlatStyle = 'Flat'; $btnApplyP.FlatAppearance.BorderSize = 0
    $btnApplyP.Size = New-Object System.Drawing.Size(150, 35)
    $pnlRight.Controls.Add($btnApplyP)
    $pJob.BtnApply = $btnApplyP

    $btnRevertDone = New-Object System.Windows.Forms.Button
    $btnRevertDone.Text = "REVERT"
    $btnRevertDone.BackColor = clr "#D95F5F"
    $btnRevertDone.ForeColor = clr "#FFFFFF"
    $btnRevertDone.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRevertDone.FlatStyle = 'Flat'
    $btnRevertDone.FlatAppearance.BorderSize = 0
    $btnRevertDone.Size = New-Object System.Drawing.Size(75, 35)
    $btnRevertDone.Visible = $false
    $pnlRight.Controls.Add($btnRevertDone)
    $pJob.BtnRevertDone = $btnRevertDone

    $btnApplyP.Tag = @{ P = $pJob; G = $gpJob }
    $btnApplyP.Add_Click({
        $t = $this.Tag
        if ($this.Text -eq "KEEP") {
            $this.Text = "Finished"
            $this.BackColor = clr "#333333"
            $this.Enabled = $false
            $this.Width = 150
            if ($t.P.BtnRevertDone) { $t.P.BtnRevertDone.Visible = $false }
            $script:resizeTimer.Start()
        } else {
            Enqueue-PJob $t.P $t.G
        }
    })

    $btnRevertDone.Tag = @{ P = $pJob; G = $gpJob }
    $btnRevertDone.Add_Click({
        $t = $this.Tag
        $pj = $t.P; $gp = $t.G
        $batPath = Join-Path $scriptDir "RevertMerge.bat"
        if (Test-Path $batPath) {
            $targetPath = if ($pj.ProcessedAnchorPath -ne "") { $pj.ProcessedAnchorPath } else { $pj.AnchorFile.FullName }
            $pj.BtnApply.Text = "Reverting..."
            $pj.BtnApply.Width = 150
            $pj.BtnRevertDone.Visible = $false
            $pj.RowPanel.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $argList = '/c ""' + $batPath + '" "' + $targetPath + '""'
                $proc = Start-Process -FilePath "cmd.exe" -ArgumentList $argList -WindowStyle Hidden -PassThru
                $timeout = 100
                while (-not $proc.HasExited -and $timeout -gt 0) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100; $timeout-- }
                if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
            } catch {}
            Refresh-PJob $pj $gp
        }
    })

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
        Parents = New-Object System.Collections.ArrayList
    }
    $script:jobs.Add($gpJob) | Out-Null

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
    $gpJob.LblGP = $lblGP

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
    $gpJob.BtnRemove = $btnRemove
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
        $gpJob.Parents.Add($pJob) | Out-Null
    }

    # Theme Sync Event
    $tbTheme.Tag = $gpJob
    $tbTheme.Add_TextChanged({
        foreach ($p in $this.Tag.Parents) {
            if ($p.LblThemeCard) { $p.LblThemeCard.Text = $this.Text }
            Update-ParentPreview $p $this.Tag
        }
    })
}

# --- 8. Dynamic Loading & Form Events ---
$btnProcessAll.Add_Click({
    foreach ($gpJob in $script:jobs) {
        if (-not $gpJob.Container.IsDisposed) {
            foreach ($pJob in $gpJob.Parents) {
                Enqueue-PJob $pJob $gpJob
            }
        }
    }
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
    $lblGlobalTitle.Text = "Queue Dashboard ($($gpQueue.Count) Theme(s) found)"
    $btnProcessAll.Enabled = $true
    $form.ResumeLayout()
    $form.Width += 1
    $script:queueTimer.Start()
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
        if ($gpJob.BtnRemove) { $gpJob.BtnRemove.Left = $gpJob.Header.Width - 180 }

        # --- NEW: PERFECT DYNAMIC SPACING ---
        if ($gpJob.LblGP -and $gpJob.TBTheme) {
            $tbX = $gpJob.LblGP.Right + 15
            if ($tbX -lt 220) { $tbX = 220 }
            $gpJob.TBTheme.Left = $tbX
            if ($gpJob.ChkSkip) { $gpJob.ChkSkip.Left = $gpJob.TBTheme.Right + 20 }
        }
        # ------------------------------------

        $pyOffset = 0
        foreach ($pJob in $gpJob.Parents) {
            $pnlFilesHeight = $pJob.FileRows.Count * 50
            $rightH = 215 + $pnlFilesHeight + 60
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

            if ($pJob.BtnRefresh) { $pJob.BtnRefresh.Left = 640 - 240 }
            if ($pJob.BtnRemove) { $pJob.BtnRemove.Left = 640 - 130 }

            if ($pJob.BtnApply) {
                if ($pJob.BtnRevertDone -and $pJob.BtnRevertDone.Visible) {
                    $pJob.BtnApply.Location = New-Object System.Drawing.Point((640 - 160), ($rightH - 45))
                    $pJob.BtnApply.Width = 70
                    $pJob.BtnRevertDone.Location = New-Object System.Drawing.Point((640 - 85), ($rightH - 45))
                } else {
                    $pJob.BtnApply.Location = New-Object System.Drawing.Point((640 - 160), ($rightH - 45))
                    $pJob.BtnApply.Width = 150
                }
            }

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

        $gpContainerH = $gpJob.Header.Height + $gpJob.ParentListPanel.Height
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

[System.Windows.Forms.Application]::Run($form)