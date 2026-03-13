# Master-Controller.ps1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- STATE VARIABLES ---
$script:isRunning = $false
$script:cancelRun = $false

# --- 1. BUILD THE MAIN WINDOW ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Wiggliteerz Master Controller"
$form.Size = New-Object System.Drawing.Size(500, 550)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$defaultFont = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Font = $defaultFont

# --- 2. UI CONTROLS (Checkboxes & Labels) ---
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Select Automation Tasks:"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

$chkColors = New-Object System.Windows.Forms.CheckBox
$chkColors.Text = "1. Scan and pick new colors"
$chkColors.Location = New-Object System.Drawing.Point(30, 50)
$chkColors.AutoSize = $true
$form.Controls.Add($chkColors)

$chkSlice = New-Object System.Windows.Forms.CheckBox
$chkSlice.Text = "2. Automated slice files"
$chkSlice.Location = New-Object System.Drawing.Point(30, 80)
$chkSlice.AutoSize = $true
$form.Controls.Add($chkSlice)

$chkImage = New-Object System.Windows.Forms.CheckBox
$chkImage.Text = "3. Generate composite image cards"
$chkImage.Location = New-Object System.Drawing.Point(250, 50)
$chkImage.AutoSize = $true
$form.Controls.Add($chkImage)

$chkExtract = New-Object System.Windows.Forms.CheckBox
$chkExtract.Text = "4. Extract data / update TSV"
$chkExtract.Location = New-Object System.Drawing.Point(250, 80)
$chkExtract.AutoSize = $true
$form.Controls.Add($chkExtract)

# Smart Checkbox Logic
$updateDependencies = {
    if ($chkSlice.Checked -or $chkImage.Checked) {
        $chkExtract.Checked = $true
    }
}
$chkSlice.Add_CheckedChanged($updateDependencies)
$chkImage.Add_CheckedChanged($updateDependencies)

# --- 3. THE LIVE LOGGING CONSOLE ---
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 120)
$txtLog.Size = New-Object System.Drawing.Size(445, 330)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LightGray
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtLog.ScrollBars = 'Vertical'
$form.Controls.Add($txtLog)

# --- HELPER FUNCTIONS ---
function Write-Log ($Message, $Color = "LightGray") {
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor = [System.Drawing.Color]::$Color
    $txtLog.AppendText("$Message`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# A non-blocking sleep so the Cancel button works instantly during delays
function Wait-Responsive($milliseconds) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $milliseconds) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
        if ($script:cancelRun) { break }
    }
    $sw.Stop()
}

# --- 4. BUTTONS ---
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Engine"
$btnStart.Location = New-Object System.Drawing.Point(345, 460)
$btnStart.Size = New-Object System.Drawing.Size(120, 35)
$btnStart.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(235, 460)
$btnCancel.Size = New-Object System.Drawing.Size(100, 35)
$form.Controls.Add($btnCancel)

# --- 5. EXECUTION LOGIC ---
$btnStart.Add_Click({
    if ($script:isRunning) { return } # Prevent double clicks

    $script:isRunning = $true
    $script:cancelRun = $false

    # Lock UI inputs, but keep the Cancel button alive!
    $chkColors.Enabled = $false
    $chkSlice.Enabled = $false
    $chkImage.Enabled = $false
    $chkExtract.Enabled = $false
    $btnStart.Enabled = $false

    # Transform Cancel button into a Stop button
    $btnCancel.Text = "Stop Engine"
    $btnCancel.BackColor = [System.Drawing.Color]::LightCoral

    $doColors  = $chkColors.Checked
    $doSlice   = $chkSlice.Checked
    $doImage   = $chkImage.Checked
    $doExtract = $chkExtract.Checked

    Write-Log "=== AUTOMATION ENGINE ENGAGED ===" "Cyan"
    Write-Log "---------------------------------" "Cyan"

    Write-Log "Scanning directory for .3mf files..." "Yellow"
    Wait-Responsive 1000

    $mockFiles = @("Bat.Full.3mf", "Ghost.Full.3mf", "Pumpkin.Full.3mf")

    foreach ($file in $mockFiles) {
        # --- THE KILL SWITCH CHECK ---
        if ($script:cancelRun) {
            Write-Log "`n>>> OPERATION ABORTED BY USER <<<" "Red"
            break
        }

        Write-Log "`nProcessing Phase 1: $file" "White"
        Wait-Responsive 500

        if ($doColors) { Write-Log "  -> Scanning colors..." "LightGray" ; Wait-Responsive 300 }
        if ($script:cancelRun) { break } # Check again after long task

        Write-Log "Processing Phase 2: $file" "White"
        if ($doSlice) { Write-Log "  -> Slicing..." "LightGray" ; Wait-Responsive 800 }
        if ($script:cancelRun) { break }

        if ($doExtract) { Write-Log "  -> Extracting Data..." "LightGray" ; Wait-Responsive 500 }
        if ($script:cancelRun) { break }

        if ($doImage) { Write-Log "  -> Generating Image Card... [DONE]" "LightGreen" ; Wait-Responsive 400 }
    }

    if (-not $script:cancelRun) {
        Write-Log "`n=== ALL TASKS COMPLETE ===" "Cyan"
    }

    # Unlock the UI and reset buttons
    $script:isRunning = $false
    $btnCancel.Enabled = $true
    $btnCancel.Text = "Close"
    $btnCancel.BackColor = [System.Drawing.Color]::LightGray

    $chkColors.Enabled = $true
    $chkSlice.Enabled = $true
    $chkImage.Enabled = $true
    $chkExtract.Enabled = $true
    $btnStart.Enabled = $true
})

$btnCancel.Add_Click({
    if ($script:isRunning) {
        Write-Log "Stopping engine... finishing current step..." "Yellow"
        $script:cancelRun = $true
        $btnCancel.Enabled = $false # Prevent spamming stop
    } else {
        $form.Close()
    }
})

# --- 6. DISPLAY THE UI ---
$form.ShowDialog() | Out-Null