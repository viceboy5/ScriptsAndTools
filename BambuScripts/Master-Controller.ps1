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
$form.Size = New-Object System.Drawing.Size(560, 580)
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

# Folder Browser
$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(20, 50)
$txtPath.Size = New-Object System.Drawing.Size(410, 25)
$txtPath.Text = "" # <--- Set to totally blank!
$form.Controls.Add($txtPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(440, 49)
$btnBrowse.Size = New-Object System.Drawing.Size(85, 27)
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select Target Folder"

    # Only set the starting directory if the text box isn't empty
    if ($txtPath.Text -ne "") { $dialog.InitialDirectory = $txtPath.Text }

    $dialog.Filter = "Folders|\n"
    $dialog.AddExtension = $false
    $dialog.CheckFileExists = $false
    $dialog.DereferenceLinks = $true
    $dialog.ValidateNames = $false
    $dialog.FileName = "Select Folder"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = Split-Path $dialog.FileName
    }
})
$form.Controls.Add($btnBrowse)

# Checkboxes
$chkColors = New-Object System.Windows.Forms.CheckBox
$chkColors.Text = "1. Scan & Pick Colors"
$chkColors.Location = New-Object System.Drawing.Point(30, 90)
$chkColors.AutoSize = $true
$form.Controls.Add($chkColors)

$chkMerge = New-Object System.Windows.Forms.CheckBox
$chkMerge.Text = "2. Merge Geometries"
$chkMerge.Location = New-Object System.Drawing.Point(30, 120)
$chkMerge.AutoSize = $true
$form.Controls.Add($chkMerge)

$chkSlice = New-Object System.Windows.Forms.CheckBox
$chkSlice.Text = "3. Slice & Export Gcode"
$chkSlice.Location = New-Object System.Drawing.Point(220, 90)
$chkSlice.AutoSize = $true
$form.Controls.Add($chkSlice)

$chkExtract = New-Object System.Windows.Forms.CheckBox
$chkExtract.Text = "4. Extract Data / TSV"
$chkExtract.Location = New-Object System.Drawing.Point(220, 120)
$chkExtract.AutoSize = $true
$form.Controls.Add($chkExtract)

$chkImage = New-Object System.Windows.Forms.CheckBox
$chkImage.Text = "5. Generate Image Card"
$chkImage.Location = New-Object System.Drawing.Point(390, 90)
$chkImage.AutoSize = $true
$form.Controls.Add($chkImage)

# Smart Dependency
$updateDependencies = {
    if ($chkSlice.Checked -or $chkImage.Checked) { $chkExtract.Checked = $true }
}
$chkSlice.Add_CheckedChanged($updateDependencies)
$chkImage.Add_CheckedChanged($updateDependencies)

# --- 3. LIVE LOGGING CONSOLE ---
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 160)
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
$btnFullProcess.Location = New-Object System.Drawing.Point(20, 495)
$btnFullProcess.Size = New-Object System.Drawing.Size(100, 35)
$btnFullProcess.BackColor = [System.Drawing.Color]::LightSkyBlue
$form.Controls.Add($btnFullProcess)

$btnRevert = New-Object System.Windows.Forms.Button
$btnRevert.Text = "Revert Merge"
$btnRevert.Location = New-Object System.Drawing.Point(130, 495)
$btnRevert.Size = New-Object System.Drawing.Size(110, 35)
$btnRevert.BackColor = [System.Drawing.Color]::Orange
$form.Controls.Add($btnRevert)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(315, 495)
$btnCancel.Size = New-Object System.Drawing.Size(90, 35)
$form.Controls.Add($btnCancel)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Selected"
$btnStart.Location = New-Object System.Drawing.Point(415, 495)
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
    $txtPath.Enabled = $false
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

    $targetDir = $txtPath.Text
    $txtLog.Clear()
    Write-Log "=== AUTOMATION ENGINE ENGAGED ===" "Cyan"
    Write-Log "Target: $targetDir" "DarkGray"

    # ---------------------------------------------------------
    # ROUTE A: REVERT MODE
    # ---------------------------------------------------------
    if ($script:isRevertMode) {
        Write-Log "MODE: REVERT MERGE" "Orange"
        Write-Log "---------------------------------" "Cyan"
        Write-Log "Scanning directory for *Nest.3mf files..." "Yellow"
        Wait-Responsive 500

        $nestFiles = Get-ChildItem -Path $targetDir -Filter "*Nest.3mf"
        if ($nestFiles.Count -eq 0) { Write-Log "[-] No Nest files found to revert." "Red" }

        foreach ($file in $nestFiles) {
            if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }
            Write-Log "`n=== Reverting: $($file.Name) ===" "Orange"

            try {
                $command = "& `"$scriptDir\RevertMerge.bat`" `"$($file.FullName)`" *>&1"
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
        $doColors  = $chkColors.Checked
        $doMerge   = $chkMerge.Checked
        $doSlice   = $chkSlice.Checked
        $doExtract = $chkExtract.Checked
        $doImage   = $chkImage.Checked

        Write-Log "---------------------------------" "Cyan"
        Write-Log "Scanning directory for source .3mf files..." "Yellow"
        Wait-Responsive 500

        $allFiles = Get-ChildItem -Path $targetDir -Filter "*Full.3mf" | Where-Object {
            $_.Name -notmatch "(?i)\.gcode\.3mf$"
        }

        if ($allFiles.Count -eq 0) { Write-Log "[-] No valid *Full.3mf files found." "Red" }

        foreach ($file in $allFiles) {
            if ($script:cancelRun) { Write-Log "`n>>> OPERATION ABORTED <<<" "Red"; break }

            $baseName = $file.BaseName
            $inputName = $file.Name
            $inputPath = $file.FullName
            Write-Log "`n=== Processing: $inputName ===" "White"

            # --- PRE-FLIGHT IMAGE CHECK ---
            if ($doImage) {
                # Look for an existing custom PNG (ignoring the final output name)
                $existingPng = Get-ChildItem -Path $targetDir -Filter "*.png" | Where-Object { $_.Name -ne "$baseName.png" } | Select-Object -First 1

                if (-not $existingPng) {
                    Write-Log "  [!] No custom image found for $inputName." "Yellow"
                    Write-Log "  -> Waiting for user to select an image..." "Cyan"

                    $imgDialog = New-Object System.Windows.Forms.OpenFileDialog
                    $imgDialog.Title = "Select Custom Image for $inputName (Cancel to skip)"
                    $imgDialog.Filter = "PNG Images (*.png)|*.png|All Files (*.*)|*.*"
                    $imgDialog.InitialDirectory = $targetDir

                    # Show the dialog and keep the GUI responsive
                    if ($imgDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedImg = $imgDialog.FileName
                        $destImg = Join-Path $targetDir (Split-Path $selectedImg -Leaf)

                        # Copy the file to the working directory if it's not already there
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

            # ---------------------------------------------------------
            # PHASE 1: PREP, COLORS, AND MERGE
            # ---------------------------------------------------------
            if ($doColors -or $doMerge) {
                # Setup Temp Working Directory
                $tempWork = Join-Path $env:TEMP ("merge_work_" + [System.IO.Path]::GetRandomFileName())
                New-Item -ItemType Directory -Path $tempWork | Out-Null

                Write-Log "  -> Extracting .3mf archive..." "DarkGray"
                try {
                    [System.IO.Compression.ZipFile]::ExtractToDirectory($inputPath, $tempWork)
                } catch {
                    Write-Log "  [!] Error extracting 3MF: $_" "Red"
                    Remove-Item -Path $tempWork -Recurse -Force -ErrorAction SilentlyContinue
                    continue # Skip to the next file
                }

                # 1. COLOR WORKER
                if ($doColors) {
                    Write-Log "  -> Scanning Colors..." "Cyan"
                    try {
                        $command = "& `"$scriptDir\update_colors_worker.ps1`" -WorkDir `"$tempWork`" -FileName `"$inputName`" -OriginalZip `"$inputPath`" *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
                if ($script:cancelRun) { break }

                # 2. MERGE WORKER
                if ($doMerge) {
                    Write-Log "  -> Merging Geometries..." "Cyan"
                    try {
                        $tempOutPath = Join-Path $targetDir "$baseName`_merged_temp.3mf"
                        $repPath = Join-Path $targetDir "$baseName`_MergeReport.txt"
                        $doColFlag = if ($doColors) { "1" } else { "0" }

                        $command = "& `"$scriptDir\merge_3mf_worker.ps1`" -WorkDir `"$tempWork`" -InputPath `"$inputPath`" -OutputPath `"$tempOutPath`" -ReportPath `"$repPath`" -DoColors `"$doColFlag`" *>&1"
                        Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }

                        # Handle the Nest / Original renaming exactly like the batch script did
                        $nestName = "$baseName`Nest.3mf"
                        $nestPath = Join-Path $targetDir $nestName
                        Rename-Item -Path $inputPath -NewName $nestName -Force
                        Rename-Item -Path $tempOutPath -NewName $inputName -Force

                        # Isolate Final Object
                        Write-Log "  -> Isolating Final Object..." "Cyan"
                        $finalPath = Join-Path $targetDir "$baseName`Final.3mf"
                        if (Test-Path $finalPath) { Remove-Item $finalPath -Force }

                        $tempSingle = Join-Path $env:TEMP ("single_work_" + [System.IO.Path]::GetRandomFileName())
                        New-Item -ItemType Directory -Path $tempSingle | Out-Null
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($nestPath, $tempSingle)

                        $isoCommand = "& `"$scriptDir\isolate_final_worker.ps1`" -WorkDir `"$tempSingle`" -OutputPath `"$finalPath`" *>&1"
                        Invoke-Expression $isoCommand | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                        Remove-Item -Path $tempSingle -Recurse -Force -ErrorAction SilentlyContinue

                    } catch { Write-Log "  [!] Error: $_" "Red" }
                }
                if ($script:cancelRun) { break }

                # Cleanup main temp work dir
                Remove-Item -Path $tempWork -Recurse -Force -ErrorAction SilentlyContinue
            }

            # ---------------------------------------------------------
            # PHASE 2: SLICE, EXTRACT, AND IMAGE
            # ---------------------------------------------------------
            # 3. SLICER WORKER
            if ($doSlice) {
                Write-Log "  -> Slicing & Exporting Gcode..." "Cyan"
                $isolatedPath = Join-Path $targetDir "$baseName`Final.3mf"
                try {
                    $command = "& `"$scriptDir\slicer_automation_worker.ps1`" -InputPath `"$inputPath`" -IsolatedPath `"$isolatedPath`" *>&1"
                    Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                } catch { Write-Log "  [!] Error: $_" "Red" }
            }
            if ($script:cancelRun) { break }

            # 4. EXTRACTION & IMAGE WORKER
            if ($doExtract -or $doImage) {
                $slicedFile = Join-Path $targetDir "$baseName.gcode.3mf"
                $singleFile = Join-Path $targetDir "$baseName`Final.gcode.3mf"

                $extractArgs = @(
                    "-InputFile", "`"$slicedFile`"",
                    "-MasterTsvPath", "`"$(Join-Path $targetDir 'Master_Data.tsv')`"",
                    "-IndividualTsvPath", "`"$(Join-Path $targetDir "$baseName`_Data.tsv")`""
                )

                if (Test-Path $singleFile) { $extractArgs += "-SingleFile", "`"$singleFile`"" }
                if ($doImage) { $extractArgs += "-GenerateImage" }
                if (-not $doExtract) { $extractArgs += "-SkipExtraction" }

                Write-Log "  -> Extracting Data / Generating Image..." "Cyan"
                try {
                    $command = "& `"$scriptDir\Extract-3MFData.ps1`" $extractArgs *>&1"
                    Invoke-Expression $command | ForEach-Object { Write-Log "     $_" "LightGray"; [System.Windows.Forms.Application]::DoEvents() }
                } catch { Write-Log "  [!] Error: $_" "Red" }
            }
        }
    }

    if (-not $script:cancelRun) { Write-Log "`n=== ALL TASKS COMPLETE ===" "LightGreen" }

    # Unlock the UI
    $script:isRunning = $false
    $script:isRevertMode = $false

    $btnCancel.Enabled = $true
    $btnCancel.Text = "Close"
    $btnCancel.BackColor = [System.Drawing.Color]::LightGray

    $txtPath.Enabled = $true
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