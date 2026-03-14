param(
    [string]$WorkDir,
    [string]$FileName = "Unknown File",
    [string]$OriginalZip = "",
    [switch]$ForceEditAll
)
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# 1. Dynamically Load the Official Library from your CSV (RGB EDITION)
$LibraryColors = [ordered]@{}
if (Test-Path $colorCsvPath) {
    $csvData = Import-Csv -Path $colorCsvPath -Header @("Name", "R", "G", "B")
    foreach ($row in $csvData) {
        $name = if ($null -ne $row.Name) { $row.Name.Trim() } else { "" }
        if ([string]::IsNullOrWhiteSpace($name) -or $name -match '(?i)^name$' -or $name -eq "N/A") { continue }
        try {
            $r = [int]$row.R; $g = [int]$row.G; $b = [int]$row.B
            $hex = "#{0:X2}{1:X2}{2:X2}FF" -f $r, $g, $b
            $LibraryColors[$name] = $hex
        } catch { continue }
    }
} else {
    $LibraryColors["Fallback Black"] = "#000000FF"
    $LibraryColors["Fallback White"] = "#FFFFFFFF"
}

# 2. Session-only memory cache
$SessionCache = @{}
foreach ($val in $LibraryColors.Values) { $SessionCache[$val.ToUpper()] = $val.ToUpper() }

# 3. Find Active Slots
$SlotMap = [ordered]@{}
$projPath = Join-Path $WorkDir "Metadata\project_settings.config"
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
    } else { exit }
} else { exit }

$UsedSlots = New-Object System.Collections.Generic.HashSet[string]
$UsedSlots.Add("1") | Out-Null

$modSetPath = Join-Path $WorkDir "Metadata\model_settings.config"
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

$modelFile = (Get-ChildItem -Path $WorkDir -Filter '3dmodel.model' -Recurse | Select-Object -First 1).FullName
if ($modelFile -and (Test-Path $modelFile)) {
    try {
        $modelContent = [System.IO.File]::ReadAllText($modelFile)
        $matMatches = [regex]::Matches($modelContent, '(?i)materialid="(\d+)"')
        foreach ($m in $matMatches) { $UsedSlots.Add($m.Groups[1].Value) | Out-Null }
    } catch {}
}

$ActiveSlotMap = [ordered]@{}
foreach ($hex in $SlotMap.Keys) {
    if ($UsedSlots.Contains($SlotMap[$hex])) { $ActiveSlotMap[$hex] = $SlotMap[$hex] }
}

if ($ActiveSlotMap.Count -eq 0) { exit }

# 4. Determine if we need to show the UI
$hasUnknowns = $false
foreach ($hex in $ActiveSlotMap.Keys) {
    $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
    if (-not $SessionCache.Contains($checkHex) -and -not $SessionCache.Contains($hex)) {
        $hasUnknowns = $true
        break
    }
}

# 5. MULTI-COLOR UI
if ($ForceEditAll -or $hasUnknowns) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Color Mapping: $FileName"
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(15)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Map the active filaments for this file:"
    $lblTitle.AutoSize = $true
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(15, 15)
    $form.Controls.Add($lblTitle)

    $yOffset = 55
    $dropdowns = @{}

    $FormatHex = {
        param([string]$h)
        if ($h.Length -eq 9) { return "#" + $h.Substring(7,2) + $h.Substring(1,6) }
        if ($h.Length -eq 7) { return $h }
        return "#FFFFFF"
    }

    foreach ($hex in $ActiveSlotMap.Keys) {
        $slotId = $ActiveSlotMap[$hex]
        $checkHex = if ($hex.Length -eq 7) { $hex + "FF" } else { $hex }
        $isUnknown = (-not $SessionCache.Contains($checkHex) -and -not $SessionCache.Contains($hex))

        # Original Color Swatch (Made larger)
        $pnlOrig = New-Object System.Windows.Forms.Panel
        $pnlOrig.Size = New-Object System.Drawing.Size(35, 35)
        $pnlOrig.Location = New-Object System.Drawing.Point(15, $yOffset)
        $pnlOrig.BorderStyle = 'FixedSingle'
        try { $pnlOrig.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $hex)) } catch {}
        $form.Controls.Add($pnlOrig)

        # Slot Label & Status (Larger font, shifted right)
        $lblY = $yOffset + 8
        $lblSlot = New-Object System.Windows.Forms.Label
        $statusText = if ($isUnknown) { "[UNKNOWN]" } else { "Matched" }
        $lblSlot.Text = "Slot $slotId - $statusText"
        if ($isUnknown) { $lblSlot.ForeColor = [System.Drawing.Color]::Red }
        $lblSlot.AutoSize = $true
        $lblSlot.Location = New-Object System.Drawing.Point(60, $lblY)
        $lblSlot.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $form.Controls.Add($lblSlot)

        # Dropdown (Restored large font, max items, type-to-search, and wider size)
        $comboY = $yOffset + 4
        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.Location = New-Object System.Drawing.Point(220, $comboY)
        $combo.Size = New-Object System.Drawing.Size(300, 28)
        $combo.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $combo.MaxDropDownItems = 25
        $combo.DropDownStyle = 'DropDown'
        $combo.AutoCompleteMode = 'SuggestAppend'
        $combo.AutoCompleteSource = 'ListItems'
        foreach ($key in $LibraryColors.Keys) { $combo.Items.Add($key) | Out-Null }

        # Pre-select matching name
        $matchedName = $null
        foreach ($key in $LibraryColors.Keys) {
            if ($LibraryColors[$key] -eq $checkHex) { $matchedName = $key; break }
        }

        if ($null -ne $matchedName) {
            $combo.SelectedItem = $matchedName
        } elseif ($combo.Items.Count -gt 0) {
            $combo.SelectedIndex = 0
        }
        $form.Controls.Add($combo)

        # New Color Swatch (Made larger)
        $pnlNew = New-Object System.Windows.Forms.Panel
        $pnlNew.Size = New-Object System.Drawing.Size(35, 35)
        $pnlNew.Location = New-Object System.Drawing.Point(535, $yOffset)
        $pnlNew.BorderStyle = 'FixedSingle'
        try {
            $initHex = $LibraryColors[$combo.Text]
            $pnlNew.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $initHex))
        } catch {}
        $form.Controls.Add($pnlNew)

        # Live Update Swatch when typing
        $combo.add_TextChanged({
            param($sender, $e)
            if ($LibraryColors.Contains($sender.Text)) {
                $selHex = $LibraryColors[$sender.Text]
                try { $pnlNew.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $selHex)) } catch {}
            } else {
                $pnlNew.BackColor = [System.Drawing.SystemColors]::Control
            }
        }.GetNewClosure())

        $dropdowns[$hex] = $combo
        $yOffset += 55 # Increased spacing between rows for the larger UI elements
    }

    # Save Button (Made larger and shifted to align right)
    $btnY = $yOffset + 10
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save mapped colors"
    $btnSave.Size = New-Object System.Drawing.Size(180, 40)
    $btnSave.Location = New-Object System.Drawing.Point(390, $btnY)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnSave)
    $form.AcceptButton = $btnSave

    $btnSave.add_Click({
        $allValid = $true
        foreach ($combo in $dropdowns.Values) {
            if (-not $LibraryColors.Contains($combo.Text)) { $allValid = $false; break }
        }

        if ($allValid) {
            $form.DialogResult = 'OK'
            $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("One or more typed colors are invalid.", "Error", 'OK', 'Warning')
        }
    })

    # Invisible spacer to pad the bottom of the window
    $botY = $yOffset + 60
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Location = New-Object System.Drawing.Point(0, $botY)
    $pnlBottom.Size = New-Object System.Drawing.Size(10, 10)
    $form.Controls.Add($pnlBottom)

    if ($form.ShowDialog() -eq 'OK') {
        # Lock in the user's choices to the SessionCache
        foreach ($hex in $dropdowns.Keys) {
            $selName = $dropdowns[$hex].Text
            $newHex = $LibraryColors[$selName].ToUpper()
            $SessionCache[$hex] = $newHex
        }
    }
}

# 6. Apply Replacements
$allTextFiles = Get-ChildItem -Path $WorkDir -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config|json)$' }
$zip = $null

try {
    if ($OriginalZip -ne "" -and (Test-Path $OriginalZip)) {
        $zip = [System.IO.Compression.ZipFile]::Open($OriginalZip, 'Update')
    }

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
            if ($null -ne $zip) {
                $relPath = $file.FullName.Substring($WorkDir.Length).TrimStart('\','/').Replace('\','/')
                $entry = $zip.GetEntry($relPath)
                if ($null -ne $entry) { $entry.Delete() }

                $newEntry = $zip.CreateEntry($relPath)
                $stream = $newEntry.Open()
                $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
                $writer.Write($content)
                $writer.Flush()
                $writer.Close()
                $stream.Close()
            }
        }
    }
} finally {
    if ($null -ne $zip) { $zip.Dispose() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}