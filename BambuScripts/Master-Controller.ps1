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
$form.Size = New-Object System.Drawing.Size(560, 650)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

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
$btnFullProcess.BackColor = [System.Drawing.Color]::LightSkyBlue
$form.Controls.Add($btnFullProcess)

$btnRevert = New-Object System.Windows.Forms.Button
$btnRevert.Text = "Revert Merge"
$btnRevert.Location = New-Object System.Drawing.Point(130, 555)
$btnRevert.Size = New-Object System.Drawing.Size(110, 35)
$btnRevert.BackColor = [System.Drawing.Color]::Orange
$form.Controls.Add($btnRevert)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(315, 555)
$btnCancel.Size = New-Object System.Drawing.Size(90, 35)
$form.Controls.Add($btnCancel)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Selected"
$btnStart.Location = New-Object System.Drawing.Point(415, 555)
$btnStart.Size = New-Object System.Drawing.Size(110, 35)
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
    $chkColors.Checked = $false
    $chkMerge.Checked  = $false
    $chkSlice.Checked  = $false
    $chkExtract.Checked = $false
    $chkImage.Checked  = $false

    $script:isRevertMode = $true
    $btnStart.PerformClick()
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
    $btnRevert.Enabled = $false

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

        foreach ($targetDir in $targetDirs) {
        if ($script:cancelRun) { break }
        Write-Log "`n--- FOLDER: $targetDir ---" "Cyan"
        $doColors  = $chkColors.Checked
        $doMerge   = $chkMerge.Checked
        $doSlice   = $chkSlice.Checked
        $doExtract = $chkExtract.Checked
        $doImage   = $chkImage.Checked

        # --- PRE-FLIGHT 'OLD' FOLDER CLEANUP ---
        $oldDirs = Get-ChildItem -Path $targetDir -Recurse -Directory | Where-Object { $_.Name -match "(?i)old" }
        if ($oldDirs) {
            $msgResult = [System.Windows.Forms.MessageBox]::Show(
                "Found folders containing the word 'old'. Do you want to delete them before processing?`n`nChoosing 'No' will keep them, but they will still be safely ignored by the engine.",
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
        # ---------------------------------------

        Write-Log "Scanning directory and subfolders for source .3mf files..." "Yellow"
        Wait-Responsive 500

        $allFiles = Get-ChildItem -Path $targetDir -Filter "*Full.3mf" -Recurse | Where-Object {
            $_.Name -notmatch "(?i)\.gcode\.3mf$" -and
            $_.FullName -notmatch "(?i)\\old"
        }

        if ($allFiles.Count -eq 0) { Write-Log "[-] No valid *Full.3mf files found." "Red" }

        # =========================================================
        # PHASE 1: INTERACTIVE PREPARATION
        # =========================================================
        $processingQueue = @()

        if ($allFiles.Count -gt 0) {
            Write-Log "`n=========================================" "Magenta"
            Write-Log " PHASE 1: INTERACTIVE PREPARATION" "White"
            Write-Log "=========================================" "Magenta"
        }

        foreach ($file in $allFiles) {
            if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }

            $baseName = $file.BaseName
            $inputName = $file.Name
            $inputPath = $file.FullName
            $fileDir = $file.DirectoryName

            Write-Log "`n=== Pre-Flight: $inputName ===" "White"

            # --- PRE-FLIGHT IMAGE CHECK ---
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

                        if ($selectedImg -ne $destImg) {
                            Copy-Item -Path $selectedImg -Destination $destImg -Force
                        }
                        Write-Log "  [+] Image selected: $(Split-Path $selectedImg -Leaf)" "LightGreen"
                    } else {
                        Write-Log "  [-] No image selected. Will use slicer fallback." "DarkGray"
                    }
                } else {
                    Write-Log "  [+] Found custom image: $($existingPng.Name)" "LightGreen"
                }
            }
            # ------------------------------

            # --- PRE-FLIGHT EXTRACTION & COLORS ---
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

            # Add to the Phase 2 Queue
            if (-not $extractFailed) {
                $processingQueue += [PSCustomObject]@{
                    File = $file
                    BaseName = $baseName
                    InputName = $inputName
                    InputPath = $inputPath
                    FileDir = $fileDir
                    TempWork = $tempWork
                }
            }
        }

        # =========================================================
        # PHASE 2: UNATTENDED PROCESSING
        # =========================================================
        if (-not $script:cancelRun -and $processingQueue.Count -gt 0) {
            Write-Log "`n=========================================" "Magenta"
            Write-Log " PHASE 2: UNATTENDED PROCESSING" "White"
            Write-Log "=========================================" "Magenta"

            # (generatedPreviews initialized before foreach - see above)

            foreach ($item in $processingQueue) {
                if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }

                # Restore variables from the Queue
                $file = $item.File
                $baseName = $item.BaseName
                $inputName = $item.InputName
                $inputPath = $item.InputPath
                $fileDir = $item.FileDir
                $tempWork = $item.TempWork

                Write-Log "`n=== Processing: $inputName ===" "White"

                $basePrefix = $baseName.Substring(0, $baseName.Length - 4)
                $nestBase   = $basePrefix + "Nest"
                $finalBase  = $basePrefix + "Final"

                # 2. MERGE WORKER
                if ($doMerge) {
                    Write-Log "  -> Merging Geometries..." "Cyan"
                    try {
                        $tempOutPath = Join-Path $fileDir "$baseName`_merged_temp.3mf"
                        $repPath = Join-Path $fileDir "$baseName`_MergeReport.txt"
                        $doColFlag = if ($doColors) { "1" } else { "0" }

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

                        $isoCommand = "& `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `"$tempSingle`" -OutputPath `"$finalPath`" *>&1"
                        Invoke-Expression $isoCommand | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                        Remove-Item -Path $tempSingle -Recurse -Force -ErrorAction SilentlyContinue

                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }

                # Cleanup Phase 1 Temp Dir
                if ($tempWork) { Remove-Item -Path $tempWork -Recurse -Force -ErrorAction SilentlyContinue }

                if ($script:cancelRun) { break }

                # 3. SLICER WORKER
                if ($doSlice) {
                    Write-Log "  -> Slicing & Exporting Gcode..." "Cyan"
                    $isolatedPath = Join-Path $fileDir "$finalBase.3mf"
                    try {
                        $command = "& `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$inputPath`" -IsolatedPath `"$isolatedPath`" *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
                if ($script:cancelRun) { break }

                # 4. EXTRACTION & IMAGE WORKER
                if ($doExtract -or $doImage) {
                    $slicedFile = Join-Path $fileDir "$baseName.gcode.3mf"
                    $singleFile = Join-Path $fileDir "$finalBase.gcode.3mf"

                    $safeToExtract = $doExtract
                    if ($safeToExtract) {
                        if (-not (Test-Path $slicedFile) -or -not (Test-Path $singleFile)) {
                            Write-Log "  [!] Missing .gcode.3mf files. Falling back to TSV-only mode to protect data." "Yellow"
                            $safeToExtract = $false
                        }
                    }

                    $extractArgs = @(
                        "-InputFile", "`"$slicedFile`"",
                        "-MasterTsvPath", "`"$(Join-Path $targetDir 'Master_Data.tsv')`"",
                        "-IndividualTsvPath", "`"$(Join-Path $fileDir "$baseName`_Data.tsv")`""
                    )

                    if (Test-Path $singleFile) { $extractArgs += "-SingleFile", "`"$singleFile`"" }
                    if ($doImage) { $extractArgs += "-GenerateImage" }
                    if (-not $safeToExtract) { $extractArgs += "-SkipExtraction" }

                    # Compute the PNG path we expect Python to write
                    $previewBaseName = $baseName -replace '(?i)[._-]Full$', ''
                    $expectedPreviewPng = Join-Path $fileDir "${previewBaseName}_slicePreview.png"

                    Write-Log "  -> Extracting Data / Generating Image..." "Cyan"
                    try {
                        $command = "& `"$scriptDir\Extract-3MFData.ps1`" $extractArgs *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }

                        # Wait up to 5 seconds for the file to appear on disk (handles Synology Drive flush lag)
                        $waitMs = 0
                        while (-not (Test-Path $expectedPreviewPng) -and $waitMs -lt 5000) {
                            Start-Sleep -Milliseconds 200
                            $waitMs += 200
                        }
                        if (Test-Path $expectedPreviewPng) {
                            $generatedPreviews.Add($expectedPreviewPng) | Out-Null
                        }

                        if ($safeToExtract -and (Test-Path $singleFile)) {
                            Remove-Item $singleFile -Force -ErrorAction SilentlyContinue
                            Write-Log "  [+] Cleaned up temporary $finalBase.gcode.3mf" "DarkGray"
                        }
                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
            }
        }
    }

    if (-not $script:cancelRun) { Write-Log "`n=== ALL TASKS COMPLETE ===" "LightGreen" }

        } # end foreach targetDir

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
    $btnRevert.Enabled = $true
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

$form.ShowDialog() | Out-Null