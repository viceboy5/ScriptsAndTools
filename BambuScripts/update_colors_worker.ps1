param(
    [string]$WorkDir,
    [string]$FileName = "Unknown File"
)
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# 1. Dynamically Load the Official Library from your CSV
$LibraryColors = [ordered]@{}
if (Test-Path $colorCsvPath) {
    $csvData = Import-Csv -Path $colorCsvPath -Header @("Hex", "Name", "R", "G", "B")
    foreach ($row in $csvData) {
        $name = if ($null -ne $row.Name) { $row.Name.Trim() } else { "" }
        $hex = if ($null -ne $row.Hex) { $row.Hex.Trim() } else { "" }

        if ([string]::IsNullOrWhiteSpace($name) -or $name -eq "N/A") { continue }
        if ($hex -match '^#[0-9a-fA-F]{6}$') { $hex += "FF" }
        if ($hex -match '^#[0-9a-fA-F]{8}$') { $LibraryColors[$name] = $hex.ToUpper() }
    }
} else {
    Write-Warning "Could not find colorNamesCSV.csv. Using fallback colors."
    $LibraryColors["Fallback Black"] = "#000000FF"
    $LibraryColors["Fallback White"] = "#FFFFFFFF"
}

# 2. Session-only memory cache auto-populated with known library colors
$SessionCache = @{}
foreach ($val in $LibraryColors.Values) { $SessionCache[$val.ToUpper()] = $val.ToUpper() }

# 3. Target project_settings.config as the absolute Source of Truth for the Palette
$SlotMap = [ordered]@{}
$projPath = Join-Path $WorkDir "Metadata\project_settings.config"

if (Test-Path $projPath) {
    $projContent = [System.IO.File]::ReadAllText($projPath, [System.Text.Encoding]::UTF8)

    # Isolate strictly the filament_colour array (ignores extruder_colour entirely)
    if ($projContent -match '(?is)"filament_colou?r"\s*:\s*\[(.*?)\]') {
        $arrayContent = $matches[1]

        # Extract each hex color inside the brackets in order
        $hexMatches = [regex]::Matches($arrayContent, '#[0-9a-fA-F]{6,8}')
        $slotIndex = 1

        foreach ($m in $hexMatches) {
            $hexColor = $m.Value.ToUpper()
            if (-not $SlotMap.Contains($hexColor)) {
                $SlotMap[$hexColor] = $slotIndex.ToString()
            }
            $slotIndex++
        }
    } else {
        Write-Host "Could not find 'filament_colour' array in project_settings.config!" -ForegroundColor Yellow
        exit
    }
} else {
    Write-Host "Could not find project_settings.config in the Metadata folder!" -ForegroundColor Red
    exit
}

# 3.5. ROBUST DETECTOR: Find exactly which slots are actively used
$UsedSlots = New-Object System.Collections.Generic.HashSet[string]
$UsedSlots.Add("1") | Out-Null # Slot 1 is always the fallback default

$modSetPath = Join-Path $WorkDir "Metadata\model_settings.config"
if (Test-Path $modSetPath) {
    try {
        # Use a true XML parser so we don't trip over attribute formatting/ordering
        [xml]$modXml = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
        foreach ($node in $modXml.SelectNodes('//metadata[contains(@key, "extruder")]')) {
            $val = $node.GetAttribute('value')
            if (-not [string]::IsNullOrWhiteSpace($val)) { $UsedSlots.Add($val) | Out-Null }
        }
    } catch {
        # Failsafe broad regex just in case the XML is technically malformed
        $modContent = [System.IO.File]::ReadAllText($modSetPath, [System.Text.Encoding]::UTF8)
        $extMatches = [regex]::Matches($modContent, '(?i)extruder[^>]*?"(\d+)"')
        foreach ($m in $extMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
    }
}

# Also scan the core 3D mesh file to catch any complex multi-color painted components
$modelFile = (Get-ChildItem -Path $WorkDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
if ($modelFile -and (Test-Path $modelFile)) {
    try {
        $modelContent = [System.IO.File]::ReadAllText($modelFile)
        $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
        foreach ($m in $matMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
    } catch {}
}

# Filter down to only the slots that are actually in use on the plate
$ActiveSlotMap = [ordered]@{}
foreach ($hex in $SlotMap.Keys) {
    if ($UsedSlots.Contains($SlotMap[$hex])) {
        $ActiveSlotMap[$hex] = $SlotMap[$hex]
    }
}

if ($ActiveSlotMap.Count -eq 0) {
    Write-Host "No active filament colors found to process." -ForegroundColor Yellow
    exit
}

# 4. UI Function
function Show-ColorPicker([string]$UnknownHex, [string]$SlotId) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Mapping Colors: $FileName"
    $form.Size = New-Object System.Drawing.Size(400, 240)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $slotText = if ([string]::IsNullOrWhiteSpace($SlotId)) { "(Unknown Slot)" } else { "(Filament Slot $SlotId)" }

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "File: $FileName`n`nRogue Hex: $UnknownHex $slotText`nPlease map this to a color from your library:"
    $label.Location = New-Object System.Drawing.Point(15, 10)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    $FormatHex = {
        param([string]$h)
        if ($h.Length -eq 9) { return "#" + $h.Substring(7,2) + $h.Substring(1,6) }
        return $h
    }

    $lblOrig = New-Object System.Windows.Forms.Label
    $lblOrig.Text = "Original:"
    $lblOrig.Location = New-Object System.Drawing.Point(15, 75)
    $lblOrig.AutoSize = $true
    $form.Controls.Add($lblOrig)

    $swatchOrig = New-Object System.Windows.Forms.Panel
    $swatchOrig.Location = New-Object System.Drawing.Point(15, 95)
    $swatchOrig.Size = New-Object System.Drawing.Size(40, 40)
    $swatchOrig.BorderStyle = 'Fixed3D'
    try { $swatchOrig.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $UnknownHex)) } catch {}
    $form.Controls.Add($swatchOrig)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point(70, 105)
    $combo.Size = New-Object System.Drawing.Size(200, 20)
    $combo.DropDownStyle = 'DropDownList'
    foreach ($key in $LibraryColors.Keys) { $combo.Items.Add($key) | Out-Null }
    $form.Controls.Add($combo)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New:"
    $lblNew.Location = New-Object System.Drawing.Point(285, 75)
    $lblNew.AutoSize = $true
    $form.Controls.Add($lblNew)

    $swatchNew = New-Object System.Windows.Forms.Panel
    $swatchNew.Location = New-Object System.Drawing.Point(285, 95)
    $swatchNew.Size = New-Object System.Drawing.Size(40, 40)
    $swatchNew.BorderStyle = 'Fixed3D'
    $form.Controls.Add($swatchNew)

    $combo.add_SelectedIndexChanged({
        $selHex = $LibraryColors[$combo.SelectedItem]
        try { $swatchNew.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $selHex)) } catch {}
    })

    if ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Map Color"
    $btn.Location = New-Object System.Drawing.Point(235, 155)
    $btn.Size = New-Object System.Drawing.Size(100, 25)
    $btn.DialogResult = 'OK'
    $form.Controls.Add($btn)
    $form.AcceptButton = $btn

    if ($form.ShowDialog() -eq 'OK' -and $null -ne $combo.SelectedItem) {
        return $LibraryColors[$combo.SelectedItem].ToUpper()
    }
    return $UnknownHex.ToUpper()
}

# 5. Trigger UI for unmapped, ACTIVE colors ONLY
foreach ($hex in @($ActiveSlotMap.Keys)) {
    # Check both 7-char and 9-char variations against the cache
    $checkHex = $hex
    if ($checkHex.Length -eq 7) { $checkHex += "FF" }

    if (-not $SessionCache.Contains($checkHex) -and -not $SessionCache.Contains($hex)) {
        $mappedHex = Show-ColorPicker -UnknownHex $hex -SlotId $ActiveSlotMap[$hex]
        $SessionCache[$hex] = $mappedHex
    }
}

# 6. Apply Replacements safely across all files (Only touches verified active colors)
$allTextFiles = Get-ChildItem -Path $WorkDir -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
foreach ($file in $allTextFiles) {
    $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
    $modified = $false

    foreach ($oldHex in $ActiveSlotMap.Keys) {
        $newHex = if ($SessionCache.Contains($oldHex)) { $SessionCache[$oldHex] } else { $SessionCache[$oldHex + "FF"] }

        if ($null -ne $newHex -and $newHex -ne $oldHex) {
            if ($content -match "(?i)$oldHex") {
                $content = $content -ireplace [regex]::Escape($oldHex), $newHex
                $modified = $true
            }
        }
    }
    if ($modified) {
        [System.IO.File]::WriteAllText($file.FullName, $content, (New-Object System.Text.UTF8Encoding($false)))
    }
}