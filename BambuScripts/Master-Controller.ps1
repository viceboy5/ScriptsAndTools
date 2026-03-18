# Master-Controller.ps1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:isRunning = $false
$script:cancelRun = $false
$script:isRevertMode = $false

# --- 1. BUILD THE MAIN WINDOW ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Wiggliteerz Master Controller"
$form.Size = New-Object System.Drawing.Size(700, 650)
$form.MinimumSize = New-Object System.Drawing.Size(700, 500)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox = $true

$defaultFont = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Font = $defaultFont

# --- 2. UI CONTROLS ---
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Wiggliteerz Build Automation"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

# --- Folder Queue (supports Browse + drag-and-drop) ---
$lstFolders = New-Object System.Windows.Forms.ListBox
$lstFolders.Location = New-Object System.Drawing.Point(20, 50)
$lstFolders.Size = New-Object System.Drawing.Size(410, 80)
$lstFolders.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$lstFolders.HorizontalScrollbar = $true
$lstFolders.Font = New-Object System.Drawing.Font("Consolas", 8)
$lstFolders.AllowDrop = $true

# Drag-and-drop: accept one or many folders dropped onto the list
$lstFolders.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

$lstFolders.Add_DragDrop({
    param($s, $e)
    $dropped = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($item in $dropped) {
        if ((Test-Path $item -PathType Container) -and -not $lstFolders.Items.Contains($item)) {
            $lstFolders.Items.Add($item) | Out-Null
        }
    }
})

$form.Controls.Add($lstFolders)

function Open-FolderDialog($startPath) {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select Target Folder"
    if ($startPath -and (Test-Path $startPath)) { $dialog.InitialDirectory = $startPath }
    $dialog.Filter = "Folders|\n"
    $dialog.AddExtension = $false
    $dialog.CheckFileExists = $false
    $dialog.DereferenceLinks = $true
    $dialog.ValidateNames = $false
    $dialog.FileName = "Select Folder"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return Split-Path $dialog.FileName
    }
    return $null
}

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Add Folder"
$btnBrowse.Location = New-Object System.Drawing.Point(440, 50)
$btnBrowse.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnBrowse.Size = New-Object System.Drawing.Size(95, 27)
$btnBrowse.Add_Click({
    $start = if ($lstFolders.Items.Count -gt 0) { $lstFolders.Items[$lstFolders.Items.Count-1] } else { "" }
    $picked = Open-FolderDialog $start
    if ($picked -and $picked -ne "" -and -not $lstFolders.Items.Contains($picked)) {
        $lstFolders.Items.Add($picked) | Out-Null
    }
})
$form.Controls.Add($btnBrowse)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove"
$btnRemove.Location = New-Object System.Drawing.Point(440, 85)
$btnRemove.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnRemove.Size = New-Object System.Drawing.Size(95, 27)
$btnRemove.Add_Click({
    if ($lstFolders.SelectedIndex -ge 0) {
        $lstFolders.Items.RemoveAt($lstFolders.SelectedIndex)
    }
})
$form.Controls.Add($btnRemove)

# Checkboxes
$chkColors = New-Object System.Windows.Forms.CheckBox
$chkColors.Text = "1. Scan & Pick Colors"
$chkColors.Location = New-Object System.Drawing.Point(30, 155)
$chkColors.AutoSize = $true
$form.Controls.Add($chkColors)

$chkMerge = New-Object System.Windows.Forms.CheckBox
$chkMerge.Text = "2. Merge Geometries"
$chkMerge.Location = New-Object System.Drawing.Point(30, 185)
$chkMerge.AutoSize = $true
$form.Controls.Add($chkMerge)

$chkSlice = New-Object System.Windows.Forms.CheckBox
$chkSlice.Text = "3. Slice & Export Gcode"
$chkSlice.Location = New-Object System.Drawing.Point(220, 155)
$chkSlice.AutoSize = $true
$form.Controls.Add($chkSlice)

$chkExtract = New-Object System.Windows.Forms.CheckBox
$chkExtract.Text = "4. Extract Data / TSV"
$chkExtract.Location = New-Object System.Drawing.Point(220, 185)
$chkExtract.AutoSize = $true
$form.Controls.Add($chkExtract)

$chkImage = New-Object System.Windows.Forms.CheckBox
$chkImage.Text = "5. Generate Image Card"
$chkImage.Location = New-Object System.Drawing.Point(390, 155)
$chkImage.AutoSize = $true
$form.Controls.Add($chkImage)

# Smart Dependency
# Smart Dependency
$updateDependencies = {
    # Image Generation uses the Extractor script, but shouldn't force a TSV overwrite!
    if ($chkSlice.Checked) { $chkExtract.Checked = $true }
}
$chkSlice.Add_CheckedChanged($updateDependencies)
# (We completely removed the $chkImage dependency here)
$chkSlice.Add_CheckedChanged($updateDependencies)
$chkImage.Add_CheckedChanged($updateDependencies)

# --- 3. LIVE LOGGING CONSOLE ---
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 220)
$txtLog.Size = New-Object System.Drawing.Size(505, 320)
$txtLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LightGray
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtLog.ScrollBars = 'Vertical'
$form.Controls.Add($txtLog)

function Write-Log ($Message, $Color = "LightGray") {
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor = [System.Drawing.Color]::$Color
    $txtLog.AppendText("$Message`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Wait-Responsive($milliseconds) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $milliseconds) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
        if ($script:cancelRun) { break }
    }
    $sw.Stop()
}

# --- 4. ACTION BUTTONS ---
$btnFullProcess = New-Object System.Windows.Forms.Button
$btnFullProcess.Text = "Full Process"
$btnFullProcess.Location = New-Object System.Drawing.Point(20, 555)
$btnFullProcess.Size = New-Object System.Drawing.Size(100, 35)
$btnFullProcess.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnFullProcess.BackColor = [System.Drawing.Color]::LightSkyBlue
$form.Controls.Add($btnFullProcess)

$btnRevert = New-Object System.Windows.Forms.Button
$btnRevert.Text = "View Merge Results"
$btnRevert.Location = New-Object System.Drawing.Point(130, 555)
$btnRevert.Size = New-Object System.Drawing.Size(110, 35)
$btnRevert.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnRevert.BackColor = [System.Drawing.Color]::SteelBlue
$btnRevert.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnRevert)

$btnCombineTsv = New-Object System.Windows.Forms.Button
$btnCombineTsv.Text = "Combine Data"
$btnCombineTsv.Size = New-Object System.Drawing.Size(120, 35)
$btnCombineTsv.Location = New-Object System.Drawing.Point(270, 555)
$btnCombineTsv.BackColor = [System.Drawing.Color]::MediumPurple
$btnCombineTsv.ForeColor = [System.Drawing.Color]::White
# Left+Right+Bottom anchor keeps it centered as the window resizes
$btnCombineTsv.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$btnCombineTsv.Add_Click({
    if ($lstFolders.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Add at least one folder to the list first.", "No Folders", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    foreach ($targetDir in $lstFolders.Items) {
        $tsvFiles = Get-ChildItem -Path $targetDir -Filter "*_Data.tsv" -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notmatch "(?i)^.*_Design_Data\.tsv$" }
        if ($tsvFiles.Count -eq 0) {
            Write-Log "  [*] No _Data.tsv files found in: $targetDir" "DarkGray"
            continue
        }
        $folderName = Split-Path $targetDir -Leaf
        $outTsvPath = Join-Path $targetDir "${folderName}_Data.tsv"
        Write-Log "`n-> Combining $($tsvFiles.Count) TSV(s) into: ${folderName}_Data.tsv" "Cyan"
        $combined = [ordered]@{}
        foreach ($tsv in $tsvFiles) {
            if ($tsv.FullName -eq $outTsvPath) { continue }
            $line = Get-Content $tsv.FullName -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $key = ($line -split "`t")[0]
            $combined[$key] = $line
        }
        $combined.Values | Set-Content -Path $outTsvPath -Encoding UTF8
        Write-Log "  [+] Written: $outTsvPath ($($combined.Count) rows)" "LightGreen"
        [System.Windows.Forms.Application]::DoEvents()
    }
})
$form.Controls.Add($btnCombineTsv)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(455, 555)
$btnCancel.Size = New-Object System.Drawing.Size(90, 35)
$btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($btnCancel)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Selected"
$btnStart.Location = New-Object System.Drawing.Point(560, 555)
$btnStart.Size = New-Object System.Drawing.Size(110, 35)
$btnStart.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$btnStart.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnStart)


# --- MACRO BUTTON LOGIC ---
$btnFullProcess.Add_Click({
    $chkColors.Checked = $true
    $chkMerge.Checked  = $true
    $chkSlice.Checked  = $true
    $chkExtract.Checked = $true
    $chkImage.Checked  = $true
    $btnStart.PerformClick()
})

$btnRevert.Add_Click({
    if ($lstFolders.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Add at least one folder to the list first.", "No Folders", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $manualQueue = @()
    foreach ($dir in $lstFolders.Items) {
        $files = Get-ChildItem -Path $dir -Filter "*Full.3mf" -Recurse -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -notmatch "(?i)\.gcode\.3mf$" -and $_.FullName -notmatch "(?i)\\old"
        }
        foreach ($f in $files) {
            $manualQueue += [PSCustomObject]@{
                File      = $f
                BaseName  = $f.BaseName
                InputName = $f.Name
                InputPath = $f.FullName
                FileDir   = $f.DirectoryName
                TempWork  = $null
                TargetDir = $dir
            }
        }
    }
    if ($manualQueue.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No *Full.3mf files found in the queued folders.", "Nothing to Review", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    Show-ResultsWindow $manualQueue @() $scriptDir
})

# --- 5. CORE EXECUTION LOGIC ---
$btnStart.Add_Click({
    if ($script:isRunning) { return }

    $script:isRunning = $true
    $script:cancelRun = $false

    # Lock the UI
    $lstFolders.Enabled = $false
    $btnRemove.Enabled = $false
    $btnBrowse.Enabled = $false
    $chkColors.Enabled = $false
    $chkMerge.Enabled = $false
    $chkSlice.Enabled = $false
    $chkExtract.Enabled = $false
    $chkImage.Enabled = $false
    $btnStart.Enabled = $false
    $btnFullProcess.Enabled = $false

    $btnCancel.Text = "Stop Engine"
    $btnCancel.BackColor = [System.Drawing.Color]::LightCoral

    $targetDirs = @($lstFolders.Items | ForEach-Object { $_ })
    $targetDir = if ($targetDirs.Count -gt 0) { $targetDirs[0] } else { "" }
    $txtLog.Clear()
    Write-Log "=== AUTOMATION ENGINE ENGAGED ===" "Cyan"
    Write-Log "Queued $($targetDirs.Count) folder(s)" "DarkGray"

    # ---------------------------------------------------------
    # ROUTE A: REVERT MODE
    # ---------------------------------------------------------
    if ($script:isRevertMode) {
        Write-Log "MODE: REVERT MERGE" "Orange"
        Write-Log "---------------------------------" "Cyan"
        Write-Log "Scanning directory and subfolders for *Nest.3mf files..." "Yellow"
        Wait-Responsive 500

        $nestFiles = Get-ChildItem -Path $targetDir -Filter "*Nest.3mf" -Recurse | Where-Object {
            $_.FullName -notmatch "(?i)\\old"
        }

        if ($nestFiles.Count -eq 0) { Write-Log "[-] No Nest files found to revert." "Red" }

        $env:WORKER_MODE = "1"

        foreach ($file in $nestFiles) {
            if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }
            Write-Log "`n=== Reverting: $($file.Name) ===" "Orange"

            try {
                $batPath = Join-Path $scriptDir "RevertMerge.bat"
                $command = "& `"$batPath`" `"$($file.FullName)`" *>&1"

                Invoke-Expression $command | ForEach-Object {
                    Write-Log "     $_" "LightGray"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            } catch { Write-Log "  [!] Error running revert script: $_" "Red" }
        }
    }
    # ---------------------------------------------------------
    # ROUTE B: STANDARD PROCESSING (Direct Worker Calls)
    # ---------------------------------------------------------
    else {
        $generatedPreviews = New-Object System.Collections.Generic.List[string]
        $globalQueue = @()

        $doColors  = $chkColors.Checked
        $doMerge   = $chkMerge.Checked
        $doSlice   = $chkSlice.Checked
        $doExtract = $chkExtract.Checked
        $doImage   = $chkImage.Checked

        # =========================================================
        # PHASE 1: INTERACTIVE PREPARATION (all folders)
        # =========================================================
        Write-Log "`n=========================================" "Magenta"
        Write-Log " PHASE 1: INTERACTIVE PREPARATION" "White"
        Write-Log "=========================================" "Magenta"

        foreach ($targetDir in $targetDirs) {
            if ($script:cancelRun) { break }
            Write-Log "`n--- FOLDER: $targetDir ---" "Cyan"

            $oldDirs = Get-ChildItem -Path $targetDir -Recurse -Directory | Where-Object { $_.Name -match "(?i)old" }
            if ($oldDirs) {
                $msgResult = [System.Windows.Forms.MessageBox]::Show(
                    "Found folders containing 'old' in:`n$targetDir`n`nDelete them before processing?`n`n'No' keeps them but they will be ignored.",
                    "Cleanup Old Folders",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($msgResult -eq 'Yes') {
                    Write-Log "Cleaning up 'old' folders..." "Yellow"
                    foreach ($dir in $oldDirs) {
                        if (Test-Path $dir.FullName) { Remove-Item $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                    Write-Log "[+] Cleanup complete." "LightGreen"
                } else {
                    Write-Log "[*] Skipped cleanup. 'Old' folders will be ignored." "DarkGray"
                }
            }

            Write-Log "Scanning for source .3mf files..." "Yellow"
            Wait-Responsive 500

            $allFiles = Get-ChildItem -Path $targetDir -Filter "*Full.3mf" -Recurse | Where-Object {
                $_.Name -notmatch "(?i)\.gcode\.3mf$" -and $_.FullName -notmatch "(?i)\\old"
            }
            if ($allFiles.Count -eq 0) { Write-Log "[-] No valid *Full.3mf files found." "Red"; continue }

            foreach ($file in $allFiles) {
                if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }
                $baseName  = $file.BaseName
                $inputName = $file.Name
                $inputPath = $file.FullName
                $fileDir   = $file.DirectoryName
                Write-Log "`n=== Pre-Flight: $inputName ===" "White"

                if ($doImage) {
                    $existingPng = Get-ChildItem -Path $fileDir -Filter "*.png" | Where-Object { $_.Name -ne "$baseName.png" -and $_.Name -notlike "*_slicePreview.png" } | Select-Object -First 1
                    if (-not $existingPng) {
                        Write-Log "  [!] No custom image found for $inputName." "Yellow"
                        Write-Log "  -> Waiting for user to select an image..." "Cyan"
                        $imgDialog = New-Object System.Windows.Forms.OpenFileDialog
                        $imgDialog.Title = "Select Custom Image for $inputName (Cancel to skip)"
                        $imgDialog.Filter = "PNG Images (*.png)|*.png|All Files (*.*)|*.*"
                        $imgDialog.InitialDirectory = $fileDir
                        if ($imgDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $selectedImg = $imgDialog.FileName
                            $destImg = Join-Path $fileDir (Split-Path $selectedImg -Leaf)
                            if ($selectedImg -ne $destImg) { Copy-Item -Path $selectedImg -Destination $destImg -Force }
                            Write-Log "  [+] Image selected: $(Split-Path $selectedImg -Leaf)" "LightGreen"
                        } else {
                            Write-Log "  [-] No image selected. Will use slicer fallback." "DarkGray"
                        }
                    } else {
                        Write-Log "  [+] Found custom image: $($existingPng.Name)" "LightGreen"
                    }
                }

                $tempWork = $null
                $extractFailed = $false
                if ($doColors -or $doMerge) {
                    $tempWork = Join-Path $env:TEMP ("merge_work_" + [System.IO.Path]::GetRandomFileName())
                    New-Item -ItemType Directory -Path $tempWork | Out-Null
                    Write-Log "  -> Extracting .3mf archive..." "DarkGray"
                    try {
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($inputPath, $tempWork)
                    } catch {
                        Write-Log "  [!] Error extracting 3MF: $_" "Red"
                        Remove-Item -Path $tempWork -Recurse -Force -ErrorAction SilentlyContinue
                        $extractFailed = $true
                    }
                    if (-not $extractFailed -and $doColors) {
                        Write-Log "  -> Scanning Colors..." "Cyan"
                        try {
                            $command = "& `"$scriptDir\update_colors_worker.ps1`" -WorkDir `"$tempWork`" -FileName `"$inputName`" -OriginalZip `"$inputPath`" *>&1"
                            Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                        } catch { Write-Log "  [!] Error: $_" "Red" }
                    }
                }
                if ($script:cancelRun) { break }

                if (-not $extractFailed) {
                    $globalQueue += [PSCustomObject]@{
                        File      = $file
                        BaseName  = $baseName
                        InputName = $inputName
                        InputPath = $inputPath
                        FileDir   = $fileDir
                        TempWork  = $tempWork
                        TargetDir = $targetDir
                    }
                }
            } # end foreach file
        } # end foreach targetDir (Phase 1)

        # =========================================================
        # PHASE 2: UNATTENDED PROCESSING (all queued items)
        # =========================================================
        if (-not $script:cancelRun -and $globalQueue.Count -gt 0) {
            Write-Log "`n=========================================" "Magenta"
            Write-Log " PHASE 2: UNATTENDED PROCESSING" "White"
            Write-Log "=========================================" "Magenta"

            foreach ($item in $globalQueue) {
                if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }
                $file      = $item.File
                $baseName  = $item.BaseName
                $inputName = $item.InputName
                $inputPath = $item.InputPath
                $fileDir   = $item.FileDir
                $tempWork  = $item.TempWork
                $targetDir = $item.TargetDir
                Write-Log "`n=== Processing: $inputName ===" "White"

                $basePrefix = $baseName.Substring(0, $baseName.Length - 4)
                $nestBase   = $basePrefix + "Nest"
                $finalBase  = $basePrefix + "Final"

                if ($doMerge) {
                    Write-Log "  -> Merging Geometries..." "Cyan"
                    try {
                        $tempOutPath = Join-Path $fileDir "$baseName`_merged_temp.3mf"
                        $repPath     = Join-Path $fileDir "$baseName`_MergeReport.txt"
                        $doColFlag   = if ($doColors) { "1" } else { "0" }
                        $command = "& `"$scriptDir\merge_3mf_worker.ps1`" -WorkDir `"$tempWork`" -InputPath `"$inputPath`" -OutputPath `"$tempOutPath`" -ReportPath `"$repPath`" -DoColors `"$doColFlag`" *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                        $nestName = "$nestBase.3mf"
                        $nestPath = Join-Path $fileDir $nestName
                        Rename-Item -Path $inputPath -NewName $nestName -Force
                        Rename-Item -Path $tempOutPath -NewName $inputName -Force
                        Write-Log "  -> Isolating Final Object..." "Cyan"
                        $finalPath = Join-Path $fileDir "$finalBase.3mf"
                        if (Test-Path $finalPath) { Remove-Item $finalPath -Force }
                        $tempSingle = Join-Path $env:TEMP ("single_work_" + [System.IO.Path]::GetRandomFileName())
                        New-Item -ItemType Directory -Path $tempSingle | Out-Null
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($nestPath, $tempSingle)
                        $isoCmd = "& `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `"$tempSingle`" -OutputPath `"$finalPath`" *>&1"
                        Invoke-Expression $isoCmd | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                        Remove-Item -Path $tempSingle -Recurse -Force -ErrorAction SilentlyContinue
                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
                if ($tempWork) { Remove-Item -Path $tempWork -Recurse -Force -ErrorAction SilentlyContinue }
                if ($script:cancelRun) { break }

                if ($doSlice) {
                    Write-Log "  -> Slicing & Exporting Gcode..." "Cyan"
                    $isolatedPath = Join-Path $fileDir "$finalBase.3mf"
                    try {
                        $command = "& `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$inputPath`" -IsolatedPath `"$isolatedPath`" *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
                if ($script:cancelRun) { break }

                if ($doExtract -or $doImage) {
                    $slicedFile = Join-Path $fileDir "$baseName.gcode.3mf"
                    $singleFile = Join-Path $fileDir "$finalBase.gcode.3mf"
                    $safeToExtract = $doExtract
                    if ($safeToExtract -and (-not (Test-Path $slicedFile) -or -not (Test-Path $singleFile))) {
                        Write-Log "  [!] Missing .gcode.3mf files. Falling back to TSV-only mode." "Yellow"
                        $safeToExtract = $false
                    }
                    $extractArgs = @(
                        "-InputFile",         "`"$slicedFile`"",
                        "-IndividualTsvPath", "`"$(Join-Path $fileDir "$baseName`_Data.tsv")`""
                    )
                    if (Test-Path $singleFile) { $extractArgs += "-SingleFile",   "`"$singleFile`"" }
                    if ($doImage)              { $extractArgs += "-GenerateImage" }
                    if (-not $safeToExtract)   { $extractArgs += "-SkipExtraction" }

                    $previewBaseName    = $baseName -replace '(?i)[ ._-]Full$', ''
                    $expectedPreviewPng = Join-Path $fileDir "${previewBaseName}_slicePreview.png"

                    Write-Log "  -> Extracting Data / Generating Image..." "Cyan"
                    try {
                        $command = "& `"$scriptDir\Extract-3MFData.ps1`" $extractArgs *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                        $waitMs = 0
                        while (-not (Test-Path $expectedPreviewPng) -and $waitMs -lt 5000) {
                            Start-Sleep -Milliseconds 200; $waitMs += 200
                        }
                        if (Test-Path $expectedPreviewPng) { $generatedPreviews.Add($expectedPreviewPng) | Out-Null }
                        if ($safeToExtract -and (Test-Path $singleFile)) {
                            Remove-Item $singleFile -Force -ErrorAction SilentlyContinue
                            Write-Log "  [+] Cleaned up temporary $finalBase.gcode.3mf" "DarkGray"
                        }
                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
            } # end foreach item (Phase 2)
        }
    } # end else Route B

    # =========================================================
    # PHASE 3: FINAL IMAGE INJECTION (Base-Solo Function)
    # =========================================================
    if (-not $script:cancelRun -and $doImage) {
        Write-Log "`n=========================================" "Magenta"
        Write-Log " PHASE 3: FINAL IMAGE INJECTION" "White"
        Write-Log "=========================================" "Magenta"

        # Use the explicitly tracked list from Phase 2 rather than scanning
        # (scanning can miss files that haven't synced to disk yet on Synology Drive)
        if ($null -eq $generatedPreviews) { $generatedPreviews = @() }

        if ($generatedPreviews.Count -gt 0) {

            # -------------------------------------------------------
            # PRE-FLIGHT: Force Synology Drive to fully download each
            # .gcode.3mf before the bat tries to rename/extract it.
            # We find every gcode file that corresponds to a preview,
            # read it fully to trigger the download, then wait for its
            # file size to stop changing (i.e. download is complete).
            # -------------------------------------------------------
            Write-Log "`n-> Pre-flight: ensuring all files are downloaded from cloud..." "Yellow"
            foreach ($png in $generatedPreviews) {
                $folder     = Split-Path $png -Parent
                $gcodeFiles = Get-ChildItem -Path $folder -Filter "*Full.gcode.3mf" -ErrorAction SilentlyContinue
                foreach ($gf in $gcodeFiles) {
                    [System.Windows.Forms.Application]::DoEvents()
                    Write-Log "  -> Checking: $($gf.Name)" "DarkGray"

                    # Read the file to trigger Synology Drive download
                    try {
                        $stream = [System.IO.File]::OpenRead($gf.FullName)
                        $buf = New-Object byte[] 65536
                        while ($stream.Read($buf, 0, $buf.Length) -gt 0) {
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                        $stream.Close()
                    } catch {
                        Write-Log "     [WARN] Could not read $($gf.Name): $_" "Orange"
                        continue
                    }

                    # Wait for file size to stabilise (download complete)
                    $prevSize = -1
                    $stableCount = 0
                    $waitSecs = 0
                    while ($stableCount -lt 3 -and $waitSecs -lt 120) {
                        [System.Windows.Forms.Application]::DoEvents()
                        $curSize = (Get-Item $gf.FullName -ErrorAction SilentlyContinue).Length
                        if ($curSize -eq $prevSize) {
                            $stableCount++
                        } else {
                            $stableCount = 0
                            Write-Log "     [WAIT] Downloading... ($([math]::Round($curSize/1MB,1)) MB)" "Yellow"
                        }
                        $prevSize = $curSize
                        $waitSecs += 2
                        Start-Sleep -Milliseconds 2000
                    }

                    if ($waitSecs -ge 120) {
                        Write-Log "     [WARN] $($gf.Name) may not be fully downloaded after 2 min." "Orange"
                    } else {
                        Write-Log "     [OK] $($gf.Name) is local ($([math]::Round((Get-Item $gf.FullName).Length/1MB,1)) MB)" "LightGreen"
                    }
                }
            }

            $batchPath = Join-Path $scriptDir "ReplaceImageNew.bat"
            if (Test-Path $batchPath) {
                Write-Log "-> Found $($generatedPreviews.Count) image(s) to inject across $($targetDirs.Count) folder(s)..." "Cyan"

                # Call the bat once per target folder so it scans the right place each time
                foreach ($injectDir in $targetDirs) {
                    if ($script:cancelRun) { break }

                    # Skip folders that produced no preview (nothing to inject)
                    $hasPng = $generatedPreviews | Where-Object { $_.StartsWith($injectDir) }
                    if (-not $hasPng) { continue }

                    Write-Log "  -> Injecting: $injectDir" "DarkGray"
                    $batLog = Join-Path $env:TEMP "BambuInject_$([System.IO.Path]::GetRandomFileName()).log"
                    $proc = Start-Process -FilePath $batchPath `
                        -ArgumentList "`"$injectDir`"" `
                        -PassThru -NoNewWindow `
                        -RedirectStandardOutput $batLog

                    $lastLine = 0
                    while (-not $proc.HasExited) {
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 150
                        if (Test-Path $batLog) {
                            $lines = @(Get-Content $batLog -ErrorAction SilentlyContinue)
                            for ($li = $lastLine; $li -lt $lines.Count; $li++) {
                                $line = $lines[$li].Trim()
                                if     ($line -match "^\[INJECTED\]") { Write-Log "     $line" "LightGreen" }
                                elseif ($line -match "^\[ERROR\]")    { Write-Log "     $line" "Red" }
                                elseif ($line -match "^\[WAIT\]")     { Write-Log "     $line" "Yellow" }
                                elseif ($line -match "^\[SKIP\]")     { Write-Log "     $line" "DarkGray" }
                                elseif ($line -match "^\[WARN\]")     { Write-Log "     $line" "Orange" }
                                elseif ($line -match "^\[OK\]")       { Write-Log "     $line" "LightGreen" }
                                elseif ($line -ne "")                  { Write-Log "     $line" "DarkGray" }
                            }
                            $lastLine = $lines.Count
                        }
                    }
                    # Flush remaining lines
                    if (Test-Path $batLog) {
                        $lines = @(Get-Content $batLog -ErrorAction SilentlyContinue)
                        for ($li = $lastLine; $li -lt $lines.Count; $li++) {
                            $line = $lines[$li].Trim()
                            if     ($line -match "^\[INJECTED\]") { Write-Log "     $line" "LightGreen" }
                            elseif ($line -match "^\[ERROR\]")    { Write-Log "     $line" "Red" }
                            elseif ($line -match "^\[WAIT\]")     { Write-Log "     $line" "Yellow" }
                            elseif ($line -match "^\[SKIP\]")     { Write-Log "     $line" "DarkGray" }
                            elseif ($line -match "^\[WARN\]")     { Write-Log "     $line" "Orange" }
                            elseif ($line -match "^\[OK\]")       { Write-Log "     $line" "LightGreen" }
                            elseif ($line -ne "")                  { Write-Log "     $line" "DarkGray" }
                        }
                        Remove-Item $batLog -Force -ErrorAction SilentlyContinue
                    }
                } # end foreach injectDir

                # Final verification: PNG consumed = success, still on disk = failure
                $failCount = 0; $successCount = 0
                foreach ($png in $generatedPreviews) {
                    if (Test-Path $png) {
                        $failCount++
                        Write-Log "  [!] PNG still on disk - injection failed: $(Split-Path $png -Leaf)" "Red"
                    } else {
                        $successCount++
                    }
                }
                if ($failCount -eq 0) {
                    Write-Log "[+] All $successCount injection(s) confirmed successful." "LightGreen"
                } else {
                    Write-Log "[!] $successCount succeeded, $failCount failed. See errors above." "Orange"
                }
            }
        } else {
            Write-Log "[*] No new generated images found to inject." "DarkGray"
        }
    }

    # --- Show results review window ---
    if (-not $script:isRevertMode -and $null -ne $globalQueue -and $globalQueue.Count -gt 0) {
        Show-ResultsWindow $globalQueue $generatedPreviews $scriptDir
    }

    # --- [ EXISTING CODE CONTINUES BELOW ] ---
    # Unlock the UI
    $script:isRunning = $false
    # ...

    # Unlock the UI
    $script:isRunning = $false
    $script:isRevertMode = $false

    $btnCancel.Enabled = $true
    $btnCancel.Text = "Close"
    $btnCancel.BackColor = [System.Drawing.Color]::LightGray

    $lstFolders.Enabled = $true
    $btnRemove.Enabled = $true
    $btnBrowse.Enabled = $true
    $chkColors.Enabled = $true
    $chkMerge.Enabled = $true
    $chkSlice.Enabled = $true
    $chkExtract.Enabled = $true
    $chkImage.Enabled = $true
    $btnStart.Enabled = $true
    $btnFullProcess.Enabled = $true
})

$btnCancel.Add_Click({
    if ($script:isRunning) {
        Write-Log "Stopping engine... finishing current step..." "Yellow"
        $script:cancelRun = $true
        $btnCancel.Enabled = $false
    } else {
        $form.Close()
    }
})

# =============================================================================
# RESULTS REVIEW WINDOW
# =============================================================================

function Invoke-RandomizePickColors($sourcePath, $destPath) {
    Add-Type -AssemblyName System.Drawing
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
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
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

function Show-ResultsWindow($queue, $previews, $scriptDir) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $rForm = New-Object System.Windows.Forms.Form
    $rForm.Text = "Results Review"
    $rForm.Size = New-Object System.Drawing.Size(900, 650)
    $rForm.StartPosition = "CenterScreen"
    $rForm.MinimumSize = New-Object System.Drawing.Size(700, 400)
    $rForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = "Review Results  -  Double-click any image to view full size"
    $lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblHeader.ForeColor = [System.Drawing.Color]::White
    $lblHeader.Location = New-Object System.Drawing.Point(10, 10)
    $lblHeader.Size = New-Object System.Drawing.Size(870, 24)
    $lblHeader.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $rForm.Controls.Add($lblHeader)

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Location = New-Object System.Drawing.Point(10, 40)
    $scroll.Size = New-Object System.Drawing.Size(865, 520)
    $scroll.AutoScroll = $true
    $scroll.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $scroll.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $rForm.Controls.Add($scroll)

    $scroll.Add_SizeChanged({
        $availW   = $scroll.ClientSize.Width - 20
        $newThumb = [int][math]::Max(120, [math]::Min(400, ($availW - 220) / 2))
        $newRowH  = $newThumb + 60
        $yPos = 5
        foreach ($row in $scroll.Controls) {
            if (-not ($row -is [System.Windows.Forms.Panel])) { continue }
            $row.Height = $newRowH
            $row.Location = New-Object System.Drawing.Point(0, $yPos)
            $yPos += $newRowH + 8
            foreach ($ctrl in $row.Controls) {
                if ($ctrl -is [System.Windows.Forms.PictureBox] -and $ctrl.Tag -is [hashtable]) {
                    if ($ctrl.Tag.Role -eq "plate") {
                        $ctrl.Location = New-Object System.Drawing.Point(5, 30)
                        $ctrl.Size = New-Object System.Drawing.Size($newThumb, $newThumb)
                    } elseif ($ctrl.Tag.Role -eq "pick") {
                        $ctrl.Location = New-Object System.Drawing.Point(($newThumb + 15), 30)
                        $ctrl.Size = New-Object System.Drawing.Size($newThumb, $newThumb)
                    }
                } elseif ($ctrl -is [System.Windows.Forms.Label] -and $ctrl.Tag -is [string]) {
                    if ($ctrl.Tag -eq "plate_lbl") {
                        $ctrl.Location = New-Object System.Drawing.Point(5, ($newThumb + 33))
                        $ctrl.Size = New-Object System.Drawing.Size($newThumb, 18)
                    } elseif ($ctrl.Tag -eq "pick_lbl") {
                        $ctrl.Location = New-Object System.Drawing.Point(($newThumb + 15), ($newThumb + 33))
                        $ctrl.Size = New-Object System.Drawing.Size($newThumb, 18)
                    }
                } elseif ($ctrl -is [System.Windows.Forms.Button]) {
                    $ctrl.Location = New-Object System.Drawing.Point(($newThumb * 2 + 25), $ctrl.Location.Y)
                } elseif ($ctrl -is [System.Windows.Forms.Label] -and $ctrl.Tag -eq $null) {
                    $ctrl.Location = New-Object System.Drawing.Point(($newThumb * 2 + 25), $ctrl.Location.Y)
                }
            }
        }
        $scroll.AutoScrollMinSize = New-Object System.Drawing.Size(0, $yPos)
    })

    $btnKeepAll = New-Object System.Windows.Forms.Button
    $btnKeepAll.Text = "Keep All & Close"
    $btnKeepAll.Size = New-Object System.Drawing.Size(150, 35)
    $btnKeepAll.Location = New-Object System.Drawing.Point(10, 570)
    $btnKeepAll.BackColor = [System.Drawing.Color]::LightGreen
    $btnKeepAll.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $rForm.Controls.Add($btnKeepAll)

    $btnUndoAll = New-Object System.Windows.Forms.Button
    $btnUndoAll.Text = "Undo All Merges"
    $btnUndoAll.Size = New-Object System.Drawing.Size(150, 35)
    $btnUndoAll.Location = New-Object System.Drawing.Point(170, 570)
    $btnUndoAll.BackColor = [System.Drawing.Color]::Orange
    $btnUndoAll.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $rForm.Controls.Add($btnUndoAll)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.ForeColor = [System.Drawing.Color]::LightGray
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblStatus.Location = New-Object System.Drawing.Point(340, 578)
    $lblStatus.Size = New-Object System.Drawing.Size(540, 20)
    $lblStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $rForm.Controls.Add($lblStatus)

    $thumbSize = 180
    $rowHeight = $thumbSize + 60
    $rowY = 5
    $rowData = @()
    $tempDir = Join-Path $env:TEMP "WiggliteerResults"
    # Always wipe on open so stale extractions from previous runs are never shown
    if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    foreach ($item in $queue) {
        $baseName  = $item.BaseName
        $fileDir   = $item.FileDir
        $gcodePath = Join-Path $fileDir "$baseName.gcode.3mf"

        $platePath = $null
        $pickPath  = $null
        if (Test-Path $gcodePath) {
            $zip = $null
            try {
                $zip = [System.IO.Compression.ZipFile]::OpenRead($gcodePath)
            } catch { $zip = $null }

            if ($zip) {
                # Extract plate_1.png independently
                try {
                    $plateEntry = $zip.Entries | Where-Object { $_.FullName -replace "\\","/" -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1
                    if ($plateEntry) {
                        $platePath = Join-Path $tempDir "${baseName}_plate.png"
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($plateEntry, $platePath, $true)
                    }
                } catch {}

                # Extract pick_1.png independently - separate try so plate failure cannot block it
                try {
                    # Normalize entry paths to handle both / and \ separators
                    $pickEntry = $zip.Entries | Where-Object { ($_.FullName -replace "\\","/") -match "(?i)(^|/)pick_1\.png$" } | Select-Object -First 1
                    if ($pickEntry) {
                        $rawPickPath = Join-Path $tempDir "${baseName}_pick_raw.png"
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pickEntry, $rawPickPath, $true)
                        $pickPath = Join-Path $tempDir "${baseName}_pick.png"
                        Invoke-RandomizePickColors $rawPickPath $pickPath | Out-Null
                        if (-not (Test-Path $pickPath)) { $pickPath = $rawPickPath }
                        Remove-Item $rawPickPath -Force -ErrorAction SilentlyContinue
                    }
                } catch {}

                $zip.Dispose()
            }
        }

        $nestBase = $baseName.Substring(0, $baseName.Length - 4) + "Nest"
        $nestPath = Join-Path $fileDir "$nestBase.3mf"
        $hasMerge = Test-Path $nestPath

        $row = New-Object System.Windows.Forms.Panel
        $row.Location = New-Object System.Drawing.Point(0, $rowY)
        $row.Size = New-Object System.Drawing.Size(840, $rowHeight)
        $row.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
        $row.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $scroll.Controls.Add($row)

        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text = $baseName
        $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lblName.ForeColor = [System.Drawing.Color]::White
        $lblName.Location = New-Object System.Drawing.Point(5, 5)
        $lblName.Size = New-Object System.Drawing.Size(500, 22)
        $row.Controls.Add($lblName)

        $pbPlate = New-Object System.Windows.Forms.PictureBox
        $pbPlate.Size = New-Object System.Drawing.Size($thumbSize, $thumbSize)
        $pbPlate.Location = New-Object System.Drawing.Point(5, 30)
        $pbPlate.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pbPlate.BackColor = [System.Drawing.Color]::Black
        $pbPlate.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pbPlate.Cursor = [System.Windows.Forms.Cursors]::Hand
        if ($platePath -and (Test-Path $platePath)) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($platePath)
                $ms = New-Object System.IO.MemoryStream(,$bytes)
                $pbPlate.Image = [System.Drawing.Image]::FromStream($ms)
            } catch {}
        }
        $pbPlate.Tag = @{ Path = $platePath; Title = "$baseName - Plate Preview"; Role = "plate" }
        $pbPlate.Add_DoubleClick({
            $d = $this.Tag
            if ($d.Path -and (Test-Path $d.Path)) { Show-ImageViewer $d.Path $d.Title }
        })
        $row.Controls.Add($pbPlate)

        $lblPlate = New-Object System.Windows.Forms.Label
        $lblPlate.Text = "Plate Preview  (double-click to enlarge)"
        $lblPlate.Tag = "plate_lbl"
        $lblPlate.ForeColor = [System.Drawing.Color]::LightGray
        $lblPlate.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $lblPlate.Location = New-Object System.Drawing.Point(5, $($thumbSize + 33))
        $lblPlate.Size = New-Object System.Drawing.Size($thumbSize, 18)
        $lblPlate.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $row.Controls.Add($lblPlate)

        $pbPick = New-Object System.Windows.Forms.PictureBox
        $pbPick.Size = New-Object System.Drawing.Size($thumbSize, $thumbSize)
        $pbPick.Location = New-Object System.Drawing.Point($($thumbSize + 15), 30)
        $pbPick.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pbPick.BackColor = [System.Drawing.Color]::Black
        $pbPick.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pbPick.Cursor = [System.Windows.Forms.Cursors]::Hand
        if ($pickPath -and (Test-Path $pickPath)) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($pickPath)
                $ms = New-Object System.IO.MemoryStream(,$bytes)
                $pbPick.Image = [System.Drawing.Image]::FromStream($ms)
            } catch {}
        } else {
            $lblNoPick = New-Object System.Windows.Forms.Label
            $lblNoPick.Text = "No pick_1.png"
            $lblNoPick.ForeColor = [System.Drawing.Color]::DarkGray
            $lblNoPick.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $lblNoPick.Location = New-Object System.Drawing.Point(10, 80)
            $lblNoPick.Size = New-Object System.Drawing.Size($($thumbSize - 20), 20)
            $lblNoPick.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $pbPick.Controls.Add($lblNoPick)
        }
        $pbPick.Tag = @{ Path = $pickPath; Title = "$baseName - Pick / Merge Check"; Role = "pick" }
        $pbPick.Add_DoubleClick({
            $d = $this.Tag
            if ($d.Path -and (Test-Path $d.Path)) { Show-ImageViewer $d.Path $d.Title }
        })
        $row.Controls.Add($pbPick)

        $lblPick = New-Object System.Windows.Forms.Label
        $lblPick.Text = "Pick / Merge Check  (double-click to enlarge)"
        $lblPick.Tag = "pick_lbl"
        $lblPick.ForeColor = [System.Drawing.Color]::LightGray
        $lblPick.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $lblPick.Location = New-Object System.Drawing.Point($($thumbSize + 15), $($thumbSize + 33))
        $lblPick.Size = New-Object System.Drawing.Size($thumbSize, 18)
        $lblPick.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $row.Controls.Add($lblPick)

        $btnUndo = New-Object System.Windows.Forms.Button
        $btnUndo.Text = if ($hasMerge) { "Undo Merge" } else { "No Merge to Undo" }
        $btnUndo.Size = New-Object System.Drawing.Size(130, 30)
        $btnUndo.Location = New-Object System.Drawing.Point($($thumbSize * 2 + 25), 50)
        $btnUndo.BackColor = if ($hasMerge) { [System.Drawing.Color]::Orange } else { [System.Drawing.Color]::DimGray }
        $btnUndo.Enabled = $hasMerge
        $btnUndo.Tag = @{ Item = $item; Row = $row; Label = $lblName; ScriptDir = $scriptDir; StatusLabel = $lblStatus }
        $btnUndo.Add_Click({
            # Capture $this immediately - it becomes unreliable after DoEvents() calls
            $myBtn = $this
            $data  = $myBtn.Tag
            $itm   = $data.Item
            $lbl   = $data.Label
            $sLbl  = $data.StatusLabel
            $myBtn.Enabled = $false
            $myBtn.Text = "Reverting..."
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $batPath = Join-Path $data.ScriptDir "RevertMerge.bat"
                if (Test-Path $batPath) {
                    $env:WORKER_MODE = "1"
                    $proc = Start-Process -FilePath $batPath `
                        -ArgumentList "`"$($itm.InputPath)`"" `
                        -PassThru -NoNewWindow
                    while (-not $proc.HasExited) {
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 100
                    }
                }
                $lbl.Text = $itm.BaseName + "  [REVERTED]"
                $lbl.ForeColor = [System.Drawing.Color]::Orange
                $myBtn.Text = "Reverted"
                $data.Row.BackColor = [System.Drawing.Color]::FromArgb(60, 40, 10)
                $sLbl.Text = "Reverted: $($itm.BaseName)"
            } catch {
                $sLbl.Text = "Error: $_"
                $myBtn.Enabled = $true
                $myBtn.Text = "Undo Merge"
            }
        })
        $row.Controls.Add($btnUndo)

        $previewBase = $baseName -replace '(?i)[ ._-]Full$', ''
        $expectedPng = Join-Path $fileDir "${previewBase}_slicePreview.png"
        $wasGenerated = ($previews -contains $expectedPng) -or ($null -ne $platePath)
        $lblRowStatus = New-Object System.Windows.Forms.Label
        $lblRowStatus.Text = if ($wasGenerated) { "[OK] Image injected" } else { "[--] No preview found" }
        $lblRowStatus.ForeColor = if ($wasGenerated) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::DarkGray }
        $lblRowStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $lblRowStatus.Location = New-Object System.Drawing.Point($($thumbSize * 2 + 25), 90)
        $lblRowStatus.Size = New-Object System.Drawing.Size(160, 20)
        $row.Controls.Add($lblRowStatus)

        $rowData += @{ Item = $item; UndoBtn = $btnUndo }
        $rowY += $rowHeight + 8
    }

    $scroll.AutoScrollMinSize = New-Object System.Drawing.Size(840, $rowY)
    $btnKeepAll.Add_Click({ $rForm.Close() })
    $btnUndoAll.Add_Click({
        $pending = @($rowData | Where-Object { $_.UndoBtn.Enabled })
        if ($pending.Count -eq 0) { $lblStatus.Text = "No merges available to undo."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to undo all $($pending.Count) merge(s)?`n`nThis cannot be undone.",
            "Confirm Undo All",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # Call revert logic directly rather than PerformClick so $this is unambiguous
        foreach ($rd in $pending) {
            $btn  = $rd.UndoBtn
            $data = $btn.Tag
            $itm  = $data.Item
            $lbl  = $data.Label
            $btn.Enabled = $false
            $btn.Text = "Reverting..."
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $batPath = Join-Path $data.ScriptDir "RevertMerge.bat"
                if (Test-Path $batPath) {
                    $env:WORKER_MODE = "1"
                    $proc = Start-Process -FilePath $batPath `
                        -ArgumentList "`"$($itm.InputPath)`"" `
                        -PassThru -NoNewWindow
                    while (-not $proc.HasExited) {
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 100
                    }
                }
                $lbl.Text = $itm.BaseName + "  [REVERTED]"
                $lbl.ForeColor = [System.Drawing.Color]::Orange
                $btn.Text = "Reverted"
                $data.Row.BackColor = [System.Drawing.Color]::FromArgb(60, 40, 10)
                $lblStatus.Text = "Reverted: $($itm.BaseName)"
                [System.Windows.Forms.Application]::DoEvents()
            } catch {
                $lblStatus.Text = "Error reverting $($itm.BaseName): $_"
                $btn.Enabled = $true
                $btn.Text = "Undo Merge"
            }
        }
        $lblStatus.Text = "All $($pending.Count) merge(s) reverted."
    })
    $rForm.ShowDialog() | Out-Null
    Get-ChildItem -Path $tempDir -Filter "*.png" | Remove-Item -Force -ErrorAction SilentlyContinue
}

$form.ShowDialog() | Out-Null