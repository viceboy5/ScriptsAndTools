param(
    [string]$WorkDir
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  COLOR STANDARDIZATION INTERCEPTOR (EXACT ORIGINAL WORKING CODE)
# ════════════════════════════════════════════════════════════════════════════════
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# 1. Dynamically Load the Official Library from your CSV
$LibraryColors = [ordered]@{}
if (Test-Path $colorCsvPath) {
    # Import CSV assuming columns: Hex, Name, R, G, B
    $csvData = Import-Csv -Path $colorCsvPath -Header @("Hex", "Name", "R", "G", "B")
    foreach ($row in $csvData) {
        $name = if ($null -ne $row.Name) { $row.Name.Trim() } else { "" }
        $hex = if ($null -ne $row.Hex) { $row.Hex.Trim() } else { "" }

        # Skip invalid or N/A rows
        if ([string]::IsNullOrWhiteSpace($name) -or $name -eq "N/A") { continue }

        # Auto-Append 'FF' alpha channel for Bambu Studio if it's a standard 6-char hex
        if ($hex -match '^#[0-9a-fA-F]{6}$') { $hex += "FF" }

        if ($hex -match '^#[0-9a-fA-F]{8}$') {
            $LibraryColors[$name] = $hex.ToUpper()
        }
    }
} else {
    Write-Warning "Could not find colorNamesCSV.csv. Using fallback colors."
    $LibraryColors["Fallback Black"] = "#000000FF"
    $LibraryColors["Fallback White"] = "#FFFFFFFF"
}

# 2. Session-only memory cache (wipes clean after this run)
$SessionCache = @{}

# Auto-add library colors to session cache so they pass validation silently
foreach ($val in $LibraryColors.Values) { $SessionCache[$val.ToUpper()] = $val.ToUpper() }

# 3. Pre-scan XML configs to map Hex Codes to their Filament Slot IDs
$SlotMap = @{}

$setPath = Join-Path $WorkDir "Metadata\model_settings.config"
if (Test-Path $setPath) {
    try {
        [xml]$cfg = [System.IO.File]::ReadAllText($setPath)
        foreach ($f in $cfg.SelectNodes('//filament')) {
            $id = $f.GetAttribute('id')
            $c = $f.SelectSingleNode('metadata[@key="color"]')
            if ($null -ne $id -and $null -ne $c) {
                $SlotMap[$c.GetAttribute('value').ToUpper()] = $id
            }
        }
    } catch {}
}

$slcPath = Join-Path $WorkDir "Metadata\slice_info.config"
if (Test-Path $slcPath) {
    try {
        [xml]$cfg = [System.IO.File]::ReadAllText($slcPath)
        foreach ($f in $cfg.SelectNodes('//filament')) {
            $id = $f.GetAttribute('id')
            $c = $f.GetAttribute('color')
            if ($null -ne $id -and $null -ne $c) {
                $SlotMap[$c.ToUpper()] = $id
            }
        }
    } catch {}
}

function Show-ColorPicker([string]$UnknownHex, [string]$SlotId) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Color Standardization Required"
    $form.Size = New-Object System.Drawing.Size(380, 220)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $slotText = if ([string]::IsNullOrWhiteSpace($SlotId)) { "(Unknown Slot)" } else { "(Filament Slot $SlotId)" }

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Rogue Hex: $UnknownHex $slotText`nPlease map this to a color from your library:"
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    # Helper: Convert Bambu #RRGGBBAA to standard #AARRGGBB for WinForms UI
    $FormatHex = {
        param([string]$h)
        if ($h.Length -eq 9) { return "#" + $h.Substring(7,2) + $h.Substring(1,6) }
        return $h
    }

    # --- ORIGINAL COLOR SWATCH ---
    $lblOrig = New-Object System.Windows.Forms.Label
    $lblOrig.Text = "Original:"
    $lblOrig.Location = New-Object System.Drawing.Point(15, 60)
    $lblOrig.AutoSize = $true
    $form.Controls.Add($lblOrig)

    $swatchOrig = New-Object System.Windows.Forms.Panel
    $swatchOrig.Location = New-Object System.Drawing.Point(15, 80)
    $swatchOrig.Size = New-Object System.Drawing.Size(40, 40)
    $swatchOrig.BorderStyle = 'Fixed3D'
    try { $swatchOrig.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $UnknownHex)) } catch {}
    $form.Controls.Add($swatchOrig)

    # --- DROPDOWN MENU ---
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point(70, 90)
    $combo.Size = New-Object System.Drawing.Size(200, 20)
    $combo.DropDownStyle = 'DropDownList'
    foreach ($key in $LibraryColors.Keys) { $combo.Items.Add($key) | Out-Null }
    $form.Controls.Add($combo)

    # --- DYNAMIC NEW COLOR SWATCH ---
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New:"
    $lblNew.Location = New-Object System.Drawing.Point(285, 60)
    $lblNew.AutoSize = $true
    $form.Controls.Add($lblNew)

    $swatchNew = New-Object System.Windows.Forms.Panel
    $swatchNew.Location = New-Object System.Drawing.Point(285, 80)
    $swatchNew.Size = New-Object System.Drawing.Size(40, 40)
    $swatchNew.BorderStyle = 'Fixed3D'
    $form.Controls.Add($swatchNew)

    # Event Listener: Update 'New' swatch dynamically when dropdown is changed
    $combo.add_SelectedIndexChanged({
        $selHex = $LibraryColors[$combo.SelectedItem]
        try { $swatchNew.BackColor = [System.Drawing.ColorTranslator]::FromHtml((&$FormatHex $selHex)) } catch {}
    })

    if ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 } # Force first selection

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Map Color"
    $btn.Location = New-Object System.Drawing.Point(235, 140)
    $btn.Size = New-Object System.Drawing.Size(100, 25)
    $btn.DialogResult = 'OK'
    $form.Controls.Add($btn)
    $form.AcceptButton = $btn

    if ($form.ShowDialog() -eq 'OK' -and $null -ne $combo.SelectedItem) {
        return $LibraryColors[$combo.SelectedItem].ToUpper()
    }
    return $UnknownHex.ToUpper()
}

# 4. Scan & Clean files BEFORE the rest of the script loads the XML DOM
$colorFiles = Get-ChildItem -Path $WorkDir -Recurse -File | Where-Object { $_.Name -match '\.(xml|model|config)$' }

foreach ($file in $colorFiles) {
    $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
    $modified = $false

    # Search for all Bambu hex strings
    $matches = [regex]::Matches($content, '#[0-9a-fA-F]{6,8}\b')
    $uniqueHexes = $matches.Value | Select-Object -Unique

    foreach ($hex in $uniqueHexes) {
        $upperHex = $hex.ToUpper()

        # If completely unknown, prompt the UI
        if (-not $SessionCache.Contains($upperHex)) {
            $slot = if ($SlotMap.Contains($upperHex)) { $SlotMap[$upperHex] } else { "" }
            $mappedHex = Show-ColorPicker -UnknownHex $upperHex -SlotId $slot
            $SessionCache[$upperHex] = $mappedHex
        }

        # If session cache dictates an update, perform exact string replacement
        if ($SessionCache[$upperHex] -ne $upperHex) {
            $content = $content -ireplace [regex]::Escape($hex), $SessionCache[$upperHex]
            $modified = $true
        }
    }

    if ($modified) {
        [System.IO.File]::WriteAllText($file.FullName, $content, (New-Object System.Text.UTF8Encoding($false)))
    }
}