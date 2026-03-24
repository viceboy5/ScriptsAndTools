Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# --- 1. Load the Official Library (Bulletproof Parser) ---
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

# --- 2. Prompt for File ---
$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Title = "Select a .3mf file to edit"
$ofd.Filter = "3MF Files (*.3mf)|*.3mf"
if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }

$targetFile = $ofd.FileName
$fileDir = Split-Path $targetFile -Parent
$baseName = (Split-Path $targetFile -Leaf) -replace '(?i)\.3mf$', ''

$parts = $baseName -split '_'
$initChar = if ($parts.Count -gt 0) { $parts[0] } else { "UNKNOWN" }
$initAdj  = if ($parts.Count -ge 3) { $parts[1] } else { "" }

# --- 3. Unpack to Temp ---
$tempWork = Join-Path $env:TEMP ("LiveCard_" + [guid]::NewGuid().ToString().Substring(0,8))
New-Item -ItemType Directory -Path $tempWork | Out-Null
[System.IO.Compression.ZipFile]::ExtractToDirectory($targetFile, $tempWork)

# --- 4. Extract Data (UNSLICED-SAFE) ---
$plateImgPath = Join-Path $tempWork "Metadata\plate_1.png"
$activeSlots = @()

$projPath   = Join-Path $tempWork "Metadata\project_settings.config"
$modSetPath = Join-Path $tempWork "Metadata\model_settings.config"
$modelFile  = (Get-ChildItem -Path $tempWork -Filter '3dmodel.model' -Recurse | Select-Object -First 1)

$SlotMap = [ordered]@{}
if (Test-Path $projPath) {
    $projContent = [System.IO.File]::ReadAllText($projPath, [System.Text.Encoding]::UTF8)
    if ($projContent -match '(?is)"filament_colou?r"\s*:\s*\[(.*?)\]') {
        $hexMatches = [regex]::Matches($matches[1], '#[0-9a-fA-F]{6,8}')
        $slotIndex = 1
        foreach ($m in $hexMatches) {
            $hexColor = $m.Value.ToUpper()
            if (-not $SlotMap.Contains($hexColor)) { $SlotMap[$hexColor] = $slotIndex.ToString() }
            $slotIndex++
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

function clr($hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { return [System.Drawing.Color]::Gray }
    if ($hex.Length -ge 9) { $hex = "#" + $hex.Substring(1,6) }
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

$cBG      = clr "#000000"; $cUI      = clr "#16171B"
$cText    = clr "#FFFFFF"; $cMuted   = clr "#A0A0A0"
$cAccent  = clr "#FFD700"; $cInput   = clr "#1E2028"
$cGrayTxt = clr "#B4B4B4"

# --- 5. THE SCALING ENGINE ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Live Card Editor - $baseName"
$form.ClientSize = New-Object System.Drawing.Size(950, 600)
$form.MinimumSize = New-Object System.Drawing.Size(800, 500)
$form.BackColor = $cUI; $form.StartPosition = "CenterScreen"

$card = New-Object System.Windows.Forms.Panel
$card.BackColor = $cBG; $card.BorderStyle = 'FixedSingle'
$form.Controls.Add($card)

# Array to hold original bounds of all card elements
$script:scaleElements = @()
function Add-ScaledElement($ctrl, $bx, $by, $bw, $bh, $bf) {
    $script:scaleElements += [PSCustomObject]@{ Ctrl = $ctrl; X = $bx; Y = $by; W = $bw; H = $bh; Font = $bf }
    $ctrl.AutoSize = $false
    $card.Controls.Add($ctrl)
}

# Image
$pbModel = New-Object System.Windows.Forms.PictureBox
$pbModel.SizeMode = 'Zoom'; $pbModel.BackColor = $cBG
if (Test-Path $plateImgPath) {
    $fs = New-Object System.IO.FileStream($plateImgPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    $pbModel.Image = [System.Drawing.Image]::FromStream($fs)
    $fs.Close()
}
Add-ScaledElement $pbModel 10 80 260 360 0

# Titles
$txtAdj = New-Object System.Windows.Forms.TextBox
$txtAdj.Text = if ($initAdj) { "($initAdj)" } else { "" }
$txtAdj.BackColor = $cBG; $txtAdj.ForeColor = $cAccent
$txtAdj.BorderStyle = 'None'; $txtAdj.TextAlign = 'Right'
Add-ScaledElement $txtAdj 320 15 180 40 18

$txtChar = New-Object System.Windows.Forms.TextBox
$txtChar.Text = $initChar
$txtChar.BackColor = $cBG; $txtChar.ForeColor = $cText
$txtChar.BorderStyle = 'None'; $txtChar.TextAlign = 'Right'
Add-ScaledElement $txtChar 60 15 250 40 18

# Dynamic Slots
$startY = 80; $boxSize = 76; $slotSpacing = $boxSize + 10; $boxX = 512 - 10 - $boxSize
$uiSlots = @()

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

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.BackColor = $cInput; $combo.ForeColor = $cText
    $combo.DropDownStyle = 'DropDown'
    $combo.AutoCompleteMode = 'SuggestAppend'
    $combo.AutoCompleteSource = 'ListItems'
    foreach ($k in $LibraryColors.Keys) { [void]$combo.Items.Add($k) }
    if ($slotData.Name) { $combo.Text = $slotData.Name } else { $combo.Text = "Select Color..." }
    $combo.Tag = $swatch

    $combo.add_TextChanged({
        param($s, $e)
        if ($LibraryColors.Contains($s.Text)) { $s.Tag.BackColor = clr $LibraryColors[$s.Text] }
        $s.SelectionStart = 0; $s.SelectionLength = 0
    }.GetNewClosure())

    # 230 width pushes it further left, guaranteeing it fits the text
    Add-ScaledElement $combo ($boxX - 240) ($startY + 15) 230 25 10

    $lblGrams = New-Object System.Windows.Forms.Label
    $lblGrams.Text = "$($slotData.Grams) g"
    $lblGrams.ForeColor = $cGrayTxt
    $lblGrams.TextAlign = 'TopRight'
    Add-ScaledElement $lblGrams ($boxX - 90) ($startY + 45) 80 25 10

    $uiSlots += [PSCustomObject]@{ OldHex = $slotData.OldHex; Combo = $combo }
    $startY += $slotSpacing
}

# --- FIX Z-ORDER ---
# Push the image to the absolute back so the dropdowns draw perfectly on top of it
$pbModel.SendToBack()

# External Tools
$lblInstructions = New-Object System.Windows.Forms.Label
$lblInstructions.Text = "LIVE 3MF EDITOR`n`nData extracted from: $baseName`n`nWhen you click save, the original .3mf will be backed up, colors will be patched, and the file will be renamed based on the canvas."
$lblInstructions.ForeColor = $cMuted; $lblInstructions.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$lblInstructions.AutoSize = $false
$form.Controls.Add($lblInstructions)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Apply & Save to 3MF"
$btnSave.BackColor = clr "#4CAF72"; $btnSave.ForeColor = clr "#FFFFFF"
$btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnSave.FlatStyle = 'Flat'; $btnSave.FlatAppearance.BorderSize = 0
$form.Controls.Add($btnSave)

# --- PERFORMANCE DEBOUNCING TIMER ---
$script:resizeTimer = New-Object System.Windows.Forms.Timer
$script:resizeTimer.Interval = 50 # Wait 50ms after dragging stops to perform the heavy font math
$script:resizeTimer.Add_Tick({
    $script:resizeTimer.Stop()
    $form.SuspendLayout()

    $availW = $form.ClientSize.Width - 320
    $availH = $form.ClientSize.Height - 40
    $sq = [math]::Min($availW, $availH)
    if ($sq -lt 300) { $sq = 300 }

    $card.Location = New-Object System.Drawing.Point(20, 20)
    $card.Size = New-Object System.Drawing.Size($sq, $sq)

    $scale = $sq / 512.0
    foreach ($item in $script:scaleElements) {
        $item.Ctrl.Location = New-Object System.Drawing.Point([int]($item.X * $scale), [int]($item.Y * $scale))

        if ($item.Ctrl -is [System.Windows.Forms.ComboBox]) {
            $item.Ctrl.Width = [int]($item.W * $scale)
        } else {
            $item.Ctrl.Size = New-Object System.Drawing.Size([int]($item.W * $scale), [int]($item.H * $scale))
        }

        if ($item.Font -gt 0) {
            $newFont = [float]($item.Font * $scale)
            if ($newFont -lt 4) { $newFont = 4 }

            # Dispose of the old font cleanly to prevent GDI Memory Leaks (lag)
            $oldFont = $item.Ctrl.Font
            $item.Ctrl.Font = New-Object System.Drawing.Font($oldFont.FontFamily, $newFont, $oldFont.Style)
            $oldFont.Dispose()
        }
    }

    $lblInstructions.Location = New-Object System.Drawing.Point(($card.Right + 20), 20)
    $lblInstructions.Size = New-Object System.Drawing.Size(280, 160)
    $btnSave.Location = New-Object System.Drawing.Point(($card.Right + 20), ($card.Bottom - 45))
    $btnSave.Size = New-Object System.Drawing.Size(280, 45)

    $form.ResumeLayout()
    $card.Invalidate()
})

$form.Add_Resize({
    # Rapidly reset the timer while the user is actively dragging the window
    $script:resizeTimer.Stop()
    $script:resizeTimer.Start()
})


# --- 6. SAVE & REPACK LOGIC ---
$btnSave.Add_Click({
    $btnSave.Text = "Saving..."
    $btnSave.Enabled = $false
    [System.Windows.Forms.Application]::DoEvents()

    $cleanChar = $txtChar.Text -replace '[^a-zA-Z0-9]', ''
    $cleanAdj  = $txtAdj.Text -replace '[^a-zA-Z0-9]', ''
    $newName = if ($cleanAdj) { "${cleanChar}_${cleanAdj}_Full.3mf" } else { "${cleanChar}_Full.3mf" }
    $newFilePath = Join-Path $fileDir $newName

    $allTextFiles = Get-ChildItem -Path $tempWork -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
    foreach ($file in $allTextFiles) {
        $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
        $modified = $false

        foreach ($slot in $uiSlots) {
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

    if ($targetFile -ne $newFilePath -and (Test-Path $targetFile)) { Rename-Item $targetFile -NewName ($baseName + "_old.3mf") -Force }
    elseif (Test-Path $newFilePath) { Remove-Item $newFilePath -Force }

    $zip = [System.IO.Compression.ZipFile]::Open($newFilePath, 'Create')
    Get-ChildItem $tempWork -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($tempWork.Length).TrimStart('\','/').Replace('\','/')
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
    }
    $zip.Dispose()

    [System.Windows.Forms.MessageBox]::Show("Successfully saved and repacked as:`n$newName", "Success")
    $form.Close()
})

$form.Add_Shown({
    $btnSave.Focus()
    # Trigger an initial resize to lock everything into place
    $form.Width = $form.Width + 1
})

$form.Add_FormClosed({
    if (Test-Path $tempWork) { Remove-Item $tempWork -Recurse -Force -ErrorAction SilentlyContinue }
})

[System.Windows.Forms.Application]::Run($form)