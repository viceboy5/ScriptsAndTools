param(
    [string]$WorkDir,
    [string]$FileName = "Unknown File",
    [string]$OriginalZip = ""
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
    # Notice the new headers: Name, R, G, B
    $csvData = Import-Csv -Path $colorCsvPath -Header @("Name", "R", "G", "B")
    foreach ($row in $csvData) {
        $name = if ($null -ne $row.Name) { $row.Name.Trim() } else { "" }

        # Skip invalid or header rows
        if ([string]::IsNullOrWhiteSpace($name) -or $name -match '(?i)^name$' -or $name -eq "N/A") { continue }

        try {
            $r = [int]$row.R
            $g = [int]$row.G
            $b = [int]$row.B
            # {0:X2} converts the number to a 2-digit uppercase Hex value
            $hex = "#{0:X2}{1:X2}{2:X2}FF" -f $r, $g, $b
            $LibraryColors[$name] = $hex
        } catch {
            # Quietly skip rows with broken/missing RGB numbers
            continue
        }
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

    if ($projContent -match '(?is)"filament_colou?r"\s*:\s*\[(.*?)\]') {
        $arrayContent = $matches[1]
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
    if ($UsedSlots.Contains($SlotMap[$hex])) {
        $ActiveSlotMap[$hex] = $SlotMap[$hex]
    }
}

if ($ActiveSlotMap.Count -eq 0) {
    Write-Host "No active filament colors found to process." -ForegroundColor Yellow
    exit
}

# --- NEW: PROMPT FOR EDIT MODE ---
$formPrompt = New-Object System.Windows.Forms.Form
$formPrompt.Text = "Color Check Mode"
$formPrompt.Size = New-Object System.Drawing.Size(400, 160)
$formPrompt.StartPosition = 'CenterScreen'
$formPrompt.FormBorderStyle = 'FixedDialog'
$formPrompt.MaximizeBox = $false
$formPrompt.TopMost = $true

$lblPrompt = New-Object System.Windows.Forms.Label
$lblPrompt.Text = "How would you like to process the active colors for:`n$FileName?"
$lblPrompt.Location = New-Object System.Drawing.Point(20, 20)
$lblPrompt.AutoSize = $true
$lblPrompt.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$formPrompt.Controls.Add($lblPrompt)

$btnUnknowns = New-Object System.Windows.Forms.Button
$btnUnknowns.Text = "Fix Unknowns Only"
$btnUnknowns.Location = New-Object System.Drawing.Point(20, 70)
$btnUnknowns.Size = New-Object System.Drawing.Size(150, 35)
$btnUnknowns.DialogResult = 'No' # Maps to "Unknowns Only"
$formPrompt.Controls.Add($btnUnknowns)

$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = "Edit ALL Active Colors"
$btnAll.Location = New-Object System.Drawing.Point(200, 70)
$btnAll.Size = New-Object System.Drawing.Size(160, 35)
$btnAll.DialogResult = 'Yes' # Maps to "Edit All"
$formPrompt.Controls.Add($btnAll)

$forceEditAll = ($formPrompt.ShowDialog() -eq 'Yes')
$formPrompt.Dispose()
# ---------------------------------

# 4. UPGRADED UI FUNCTION
function Show-ColorPicker([string]$UnknownHex, [string]$SlotId) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Mapping Colors: $FileName"
    $form.Size = New-Object System.Drawing.Size(550, 320)
    $form.MinimumSize = New-Object System.Drawing.Size(450, 280)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true

    # Enable resizing
    $form.FormBorderStyle = 'Sizable'

    $slotText = if ([string]::IsNullOrWhiteSpace($SlotId)) { "(Unknown Slot)" } else { "(Filament Slot $SlotId)" }

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "File: $FileName`n`nRogue Hex: $UnknownHex $slotText`nPlease select or type a color from your library:"
    $label.Location = New-Object System.Drawing.Point(15, 15)
    # Removing fixed size and using AutoSize so the bigger font doesn't get cut off
    $label.AutoSize = $true
    $label.Anchor = 'Top, Left, Right'
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($label)

    $FormatHex = {
        param([string]$h)
        if ($h.Length -eq 9) { return "#" + $h.Substring(7,2) + $h.Substring(1,6) }
        return $h
    }

    $lblOrig = New-Object System.Windows.Forms.Label
    $lblOrig.Text = "Original:"
    $lblOrig.Location = New-Object System.Drawing.Point(15, 95)
    $lblOrig.AutoSize = $true
    $lblOrig.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($lblOrig)

    $swatchOrig = New-Object System.Windows.Forms.Panel
    $swatchOrig.Location = New-Object System.Drawing.Point(15, 120)
    $swatchOrig.Size = New-Object System.Drawing.Size(55, 55)
    $swatchOrig.BorderStyle = 'Fixed3D'
    try { $swatchOrig.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $UnknownHex)) } catch {}
    $form.Controls.Add($swatchOrig)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point(85, 135)
    $combo.Size = New-Object System.Drawing.Size(320, 28)
    $combo.Anchor = 'Top, Left, Right'
    $combo.Font = New-Object System.Drawing.Font("Segoe UI", 11)

    # UI Upgrades: Tall dropdown list, Type-to-search enabled
    $combo.MaxDropDownItems = 25
    $combo.DropDownStyle = 'DropDown'
    $combo.AutoCompleteMode = 'SuggestAppend'
    $combo.AutoCompleteSource = 'ListItems'

    foreach ($key in $LibraryColors.Keys) { $combo.Items.Add($key) | Out-Null }
    $form.Controls.Add($combo)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New:"
    $lblNew.Location = New-Object System.Drawing.Point(425, 95)
    $lblNew.AutoSize = $true
    $lblNew.Anchor = 'Top, Right'
    $lblNew.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($lblNew)

    $swatchNew = New-Object System.Windows.Forms.Panel
    $swatchNew.Location = New-Object System.Drawing.Point(425, 120)
    $swatchNew.Size = New-Object System.Drawing.Size(55, 55)
    $swatchNew.BorderStyle = 'Fixed3D'
    $swatchNew.Anchor = 'Top, Right'
    $form.Controls.Add($swatchNew)

    # Dynamic color update as you type
    $combo.add_TextChanged({
        if ($LibraryColors.Contains($combo.Text)) {
            $selHex = $LibraryColors[$combo.Text]
            try { $swatchNew.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $selHex)) } catch {}
        } else {
            # FIX: Correctly pull the default Windows gray background so it doesn't crash on invalid text
            $swatchNew.BackColor = [System.Drawing.SystemColors]::Control
        }
    })

    # --- NEW: Reverse-lookup to pre-select the known name ---
    $matchedName = $null
    foreach ($key in $LibraryColors.Keys) {
        if ($LibraryColors[$key] -eq $UnknownHex.ToUpper()) {
            $matchedName = $key
            break
        }
    }

    if ($null -ne $matchedName) {
        $combo.SelectedItem = $matchedName
    } elseif ($combo.Items.Count -gt 0) {
        $combo.SelectedIndex = 0
    }
    # --------------------------------------------------------

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Map Color"
    $btn.Location = New-Object System.Drawing.Point(380, 220)
    $btn.Size = New-Object System.Drawing.Size(130, 35)
    $btn.Anchor = 'Bottom, Right'
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btn)
    $form.AcceptButton = $btn

    # Validate the typed text on click
    $btn.add_Click({
        if ($LibraryColors.Contains($combo.Text)) {
            $form.DialogResult = 'OK'
            $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select or type a valid color name from the list.", "Invalid Color", 'OK', 'Warning')
        }
    })

    if ($form.ShowDialog() -eq 'OK') {
        return $LibraryColors[$combo.Text].ToUpper()
    }
    return $UnknownHex.ToUpper()
}

# 5. Trigger UI based on user choice
foreach ($hex in @($ActiveSlotMap.Keys)) {
    $checkHex = $hex
    if ($checkHex.Length -eq 7) { $checkHex += "FF" }

    # If ForceEditAll is true, OR the color is missing from the cache, trigger the UI
    if ($forceEditAll -or (-not $SessionCache.Contains($checkHex) -and -not $SessionCache.Contains($hex))) {

        # We temporarily remove it from the cache so the UI knows it needs a new map
        if ($forceEditAll) { $SessionCache.Remove($checkHex); $SessionCache.Remove($hex) }

        $mappedHex = Show-ColorPicker -UnknownHex $hex -SlotId $ActiveSlotMap[$hex]
        $SessionCache[$hex] = $mappedHex
    }
}

# 6. Apply Replacements safely across both the extracted folder AND the original zip
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
    $zip = $null
    Remove-Variable zip -ErrorAction SilentlyContinue
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}