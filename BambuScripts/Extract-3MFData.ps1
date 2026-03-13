param (
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$InputFile,

    [string]$SingleFile = "",

    [switch]$ConsoleOnly,

    [string]$MasterTsvPath = "",
    [string]$IndividualTsvPath = "",

    [switch]$GenerateImage,
    [switch]$SkipExtraction
)

if ($ConsoleOnly) {
    Write-Host "Processing $InputFile (Console Output Only)..." -ForegroundColor Cyan
} else {
    Write-Host "Processing $InputFile (Formatting for Google Sheets)..." -ForegroundColor Cyan
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# --- 1. Load the Reverse-Lookup Color Dictionary (RGB EDITION) ---
$LibraryNames = @{}
$NameToHex = @{}
if (Test-Path $colorCsvPath) {
    $csvLines = Get-Content -Path $colorCsvPath
    foreach ($line in $csvLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split ','

        # Ensure we have at least Name, R, G, B
        if ($parts.Count -ge 4) {
            $name = $parts[0].Replace('"','').Trim()
            if ($name -match '(?i)^name$' -or $name -eq "N/A" -or $name -eq "") { continue }

            try {
                $r = [int]$parts[1].Replace('"','').Trim()
                $g = [int]$parts[2].Replace('"','').Trim()
                $b = [int]$parts[3].Replace('"','').Trim()

                $rawHex = "{0:X2}{1:X2}{2:X2}FF" -f $r, $g, $b
                $LibraryNames[$rawHex] = $name

                # --- NEW: Save Hex for TSV Reversal ---
                $NameToHex[$name] = "#" + $rawHex
            } catch { continue }
        }
    }
} else {
    Write-Warning "Could not find colorNamesCSV.csv for reverse lookup."
}

# --- 2. Define the Native C# Engine ---
$csharpCode = @"
using System;
using System.IO;
using System.Globalization;

public class GcodeAnalyzer {
    public double FlushGrams { get; set; }
    public double TowerGrams { get; set; }
    public string PrintTime { get; set; }
    public int ColorChanges { get; set; }

    public void Analyze(Stream gcodeStream, double gramsPerMm) {
        using (StreamReader reader = new StreamReader(gcodeStream)) {
            string line;
            double flushE = 0;
            double towerE = 0;
            bool inFlush = false;
            bool inTower = false;
            bool isRelative = false;
            double currentE = 0;
            double maxE = 0;

            this.PrintTime = "Not found";
            this.ColorChanges = 0;

            while ((line = reader.ReadLine()) != null) {
                line = line.TrimStart();
                if (line.Length == 0) continue;

                char c = line[0];

                if (c == 'G') {
                    if (line.StartsWith("G1 ") || line.StartsWith("G0 ")) {
                        int eIdx = line.IndexOf(" E");
                        if (eIdx > -1) {
                            int startIdx = eIdx + 2;
                            int endIdx = line.IndexOf(' ', startIdx);
                            if (endIdx == -1) endIdx = line.IndexOf(';', startIdx);
                            if (endIdx == -1) endIdx = line.Length;

                            string eStr = line.Substring(startIdx, endIdx - startIdx);
                            double eVal;
                            if (double.TryParse(eStr, NumberStyles.Any, CultureInfo.InvariantCulture, out eVal)) {
                                if (isRelative) currentE += eVal;
                                else currentE = eVal;

                                if (currentE > maxE) {
                                    double delta = currentE - maxE;
                                    if (inFlush) flushE += delta;
                                    else if (inTower) towerE += delta;
                                    maxE = currentE;
                                }
                            }
                        }
                    }
                    else if (line.StartsWith("G92")) {
                        int eIdx = line.IndexOf(" E");
                        if (eIdx > -1) {
                            int startIdx = eIdx + 2;
                            int endIdx = line.IndexOf(' ', startIdx);
                            if (endIdx == -1) endIdx = line.IndexOf(';', startIdx);
                            if (endIdx == -1) endIdx = line.Length;

                            string eStr = line.Substring(startIdx, endIdx - startIdx);
                            double eVal;
                            if (double.TryParse(eStr, NumberStyles.Any, CultureInfo.InvariantCulture, out eVal)) {
                                currentE = eVal;
                                maxE = currentE;
                            }
                        } else {
                            currentE = 0;
                            maxE = 0;
                        }
                    }
                }
                else if (c == ';') {
                    if (line.StartsWith("; FLUSH_START")) inFlush = true;
                    else if (line.StartsWith("; FLUSH_END")) inFlush = false;
                    else if (line.StartsWith("; WIPE_TOWER_START")) inTower = true;
                    else if (line.StartsWith("; WIPE_TOWER_END")) inTower = false;
                    else if (line.StartsWith("; TYPE:")) {
                        if (line.Contains("Wipe tower") || line.Contains("Prime tower")) inTower = true;
                        else inTower = false;
                    }
                    else if (line.Contains("total estimated time:")) {
                        int colIdx = line.IndexOf("total estimated time:");
                        if (colIdx > -1) {
                            string pt = line.Substring(colIdx + 21).Trim();
                            int semiIdx = pt.IndexOf(';');
                            if (semiIdx > -1) pt = pt.Substring(0, semiIdx).Trim();
                            this.PrintTime = pt;
                        }
                    }
                    else if (line.Contains("estimated printing time") && this.PrintTime == "Not found") {
                        int eqIdx = line.IndexOf('=');
                        if (eqIdx > -1) {
                            string pt = line.Substring(eqIdx + 1).Trim();
                            int semiIdx = pt.IndexOf(';');
                            if (semiIdx > -1) pt = pt.Substring(0, semiIdx).Trim();
                            this.PrintTime = pt;
                        }
                    }
                }
                else if (c == 'M') {
                    if (line.StartsWith("M83")) isRelative = true;
                    else if (line.StartsWith("M82")) isRelative = false;
                    else if (line.StartsWith("M620 S")) this.ColorChanges++;
                }
            }

            this.FlushGrams = flushE * gramsPerMm;
            this.TowerGrams = towerE * gramsPerMm;
        }
    }
}
"@

if (-not ("GcodeAnalyzer" -as [type])) { Add-Type -TypeDefinition $csharpCode -Language CSharp }
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------------
# INITIALIZE VARIABLES
# ---------------------------------------------------------------------------------
$projectName = ((Split-Path $InputFile -Leaf) -replace '\.gcode\.3mf$', '')
$timeAdd = 0
$filData = @(
    @{ g = 0; color = ""; rawHex = "" }, @{ g = 0; color = ""; rawHex = "" },
    @{ g = 0; color = ""; rawHex = "" }, @{ g = 0; color = ""; rawHex = "" }, @{ g = 0; color = ""; rawHex = "" }
)
$outputValues = @()
$h = 0; $m = 0; $modelGrams = 0; $objCount = 0; $actualColorSwaps = 0
$singlePrintTimeStr = "N/A"
$analyzer = $null


# ---------------------------------------------------------------------------------
# DATA COLLECTION (HEAVY EXTRACTION OR FAST TSV READ)
# ---------------------------------------------------------------------------------
if (-not $SkipExtraction) {
    try {
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($InputFile)

        $totalGrams = 0.0; $totalMeters = 0.0; $objCount = 0

        # --- 3. Extract Metadata (XML) ---
        $configEntry = $zipArchive.Entries | Where-Object { $_.FullName -replace '\\', '/' -match "Metadata/slice_info\.config$" }
        if ($configEntry) {
            $configStream = $configEntry.Open()
            $configReader = [System.IO.StreamReader]::new($configStream)
            $configContent = $configReader.ReadToEnd()
            $configReader.Close()

            try {
                [xml]$xml = $configContent

                foreach ($i in 1..4) {
                    $node = $xml.SelectSingleNode("//filament[@id='$i']")
                    if ($node) {
                        $weight = 0.0
                        [double]::TryParse($node.used_g, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$weight) | Out-Null

                        if ($weight -gt 0) {
                            $filData[$i].g = [math]::Round($weight, 2)

                            # Normalize hex code to pure alphanumeric for the dictionary match
                            $rawHex = $node.color.Replace('"','').Trim().ToUpper()
                            if ($rawHex.StartsWith('#')) { $rawHex = $rawHex.Substring(1) }

                            if ($rawHex.Length -eq 6) { $rawHex += "FF" }

                            $filData[$i].rawHex = "#" + $rawHex

                            if ($LibraryNames.Contains($rawHex)) {
                                $filData[$i].color = $LibraryNames[$rawHex]
                            } else {
                                $filData[$i].color = "#" + $rawHex
                            }
                        }
                    }
                }

                # --- STRICT Pre-Merge Object Count Logic ---
                $objs = $xml.SelectNodes('//plate/object')
                if ($objs) {
                    foreach ($o in $objs) {
                        $objName = $o.name
                        if ($objName -match '(?i)text|version') { continue }
                        if ($objName -match 'MergedGroup_(\d+)$') { $objCount += [int]$matches[1] }
                        else { $objCount += 1 }
                    }
                }
            } catch { Write-Warning "Could not parse XML properly." }

            $gramMatches = [regex]::Matches($configContent, 'used_g="([0-9.]+)"')
            foreach ($match in $gramMatches) { $totalGrams += [double]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) }
            $meterMatches = [regex]::Matches($configContent, 'used_m="([0-9.]+)"')
            foreach ($match in $meterMatches) { $totalMeters += [double]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) }
        }

        $gramsPerMm = 0.0
        if ($totalMeters -gt 0) { $gramsPerMm = $totalGrams / ($totalMeters * 1000) }

        # --- 4. Pass the torch to C# (Full Plate) ---
        $gcodeEntry = $zipArchive.Entries | Where-Object { $_.FullName -like "*.gcode" } | Select-Object -First 1
        $analyzer = New-Object GcodeAnalyzer
        if ($gcodeEntry) {
            $gcodeStream = $gcodeEntry.Open()
            $analyzer.Analyze($gcodeStream, $gramsPerMm)
        }
        $zipArchive.Dispose()

        # --- 5. Crunch Final Math & Format ---
        $modelGrams = $totalGrams - $analyzer.FlushGrams - $analyzer.TowerGrams

        $d = 0; $h = 0; $m = 0
        if ($analyzer.PrintTime -match '(\d+)d') { $d = [int]$matches[1] }
        if ($analyzer.PrintTime -match '(\d+)h') { $h = [int]$matches[1] }
        if ($analyzer.PrintTime -match '(\d+)m') { $m = [int]$matches[1] }
        if ($analyzer.PrintTime -match '(\d+)s' -and [int]$matches[1] -ge 30) { $m++ }

        $h += ($d * 24)
        if ($m -ge 60) { $m -= 60; $h++ }

        $totalMinutes = ($h * 60) + $m
        $actualColorSwaps = [math]::Max(0, ($analyzer.ColorChanges - 1))

        # --- 6. Extract Single Object Time (If Provided) ---
        if ($SingleFile -ne "" -and (Test-Path $SingleFile)) {
            try {
                $singleArchive = [System.IO.Compression.ZipFile]::OpenRead($SingleFile)
                $singleGcode = $singleArchive.Entries | Where-Object { $_.FullName -like "*.gcode" } | Select-Object -First 1
                $singleAnalyzer = New-Object GcodeAnalyzer

                if ($singleGcode) {
                    $singleStream = $singleGcode.Open()
                    $singleAnalyzer.Analyze($singleStream, 0)
                    $singlePrintTimeStr = $singleAnalyzer.PrintTime

                    $sd = 0; $sh = 0; $sm = 0
                    if ($singleAnalyzer.PrintTime -match '(\d+)d') { $sd = [int]$matches[1] }
                    if ($singleAnalyzer.PrintTime -match '(\d+)h') { $sh = [int]$matches[1] }
                    if ($singleAnalyzer.PrintTime -match '(\d+)m') { $sm = [int]$matches[1] }
                    if ($singleAnalyzer.PrintTime -match '(\d+)s' -and [int]$matches[1] -ge 30) { $sm++ }

                    $sh += ($sd * 24)
                    if ($sm -ge 60) { $sm -= 60; $sh++ }

                    $singleTotalMinutes = ($sh * 60) + $sm

                    if ($objCount -gt 1) {
                        $timeAdd = [math]::Round(($totalMinutes - $singleTotalMinutes) / ($objCount - 1), 2)
                    }
                }
                $singleArchive.Dispose()
            } catch {
                Write-Warning "Failed to read Single Object file for Time Add/Wig math."
                if ($null -ne $singleArchive) { $singleArchive.Dispose() }
            }
        }

        $outputValues = @(
            $projectName, "", (Get-Date).ToString("M/d/yyyy"),
            $h, $m,
            $(if ($filData[1].g -gt 0) { $filData[1].g } else { 0 }), $filData[1].color,
            $(if ($filData[2].g -gt 0) { $filData[2].g } else { 0 }), $filData[2].color,
            $(if ($filData[3].g -gt 0) { $filData[3].g } else { 0 }), $filData[3].color,
            $(if ($filData[4].g -gt 0) { $filData[4].g } else { 0 }), $filData[4].color,
            "", $actualColorSwaps, $objCount, [math]::Round($modelGrams, 2), "", $timeAdd
        )
    } catch {
        Write-Error "Failed to process .gcode.3mf archive: $_"
        if ($null -ne $zipArchive) { $zipArchive.Dispose() }
    }
} else {
    # --- FAST TSV-ONLY MODE ---
    if ($IndividualTsvPath -ne "" -and (Test-Path $IndividualTsvPath)) {
        Write-Host "  -> Skipping data extraction. Loading values from TSV..." -ForegroundColor Cyan
        try {
            $existingData = Get-Content $IndividualTsvPath | Select-Object -Last 1
            $cols = $existingData -split "`t"

            if ($cols.Count -ge 19) {
                # Rebuild filament array from TSV columns
                for ($i = 1; $i -le 4; $i++) {
                    $gIdx = 5 + (($i - 1) * 2)
                    $cIdx = $gIdx + 1
                    if ([double]::TryParse($cols[$gIdx], [ref]$null)) { $filData[$i].g = [double]$cols[$gIdx] }

                    $filData[$i].color = $cols[$cIdx]
                    if ($filData[$i].color -ne "") {
                        if ($filData[$i].color.StartsWith("#")) {
                            $filData[$i].rawHex = $filData[$i].color
                        } else {
                            $filData[$i].rawHex = $NameToHex[$filData[$i].color]
                        }
                    }
                }
                if ([double]::TryParse($cols[18], [ref]$null)) { $timeAdd = [double]$cols[18] }
            }
        } catch {
            Write-Warning "Failed to read data from TSV."
        }
    } else {
        Write-Warning "  -> [!] No TSV found to load data from. Image will be blank."
    }
}

# ---------------------------------------------------------------------------------
# 7. AUTO-GENERATE COMPOSITE IMAGE
# ---------------------------------------------------------------------------------
if ($GenerateImage) {
    try {
        $pyScript = Join-Path $scriptDir "generate_image_worker.py"
        if (Test-Path $pyScript) {
            Write-Host "  -> Generating Composite Card... " -ForegroundColor Cyan -NoNewline

            $inputFolder = Split-Path $InputFile -Parent
            $outImg = Join-Path $inputFolder "$projectName.png"

            # Python strips [._-]Full from the name and appends _slicePreview.png
            # e.g. projectName "Bat.Full" -> pyBaseName "Bat" -> "Bat_slicePreview.png"
            $pyBaseName = $projectName -replace '(?i)[._-]Full$', ''
            $expectedPng = Join-Path $inputFolder "${pyBaseName}_slicePreview.png"

            $sourceImg = ""
            $isTemp = $false

            # 1. Search the folder for a custom PNG (never pick up a previously generated slicePreview)
            $customPng = Get-ChildItem -Path $inputFolder -Filter "*.png" |
                         Where-Object { $_.Name -ne "$projectName.png" -and $_.Name -notlike "*_slicePreview.png" } |
                         Select-Object -First 1

            if ($customPng) {
                $sourceImg = $customPng.FullName
            } else {
                # 2. SILENT FALLBACK to internal plate_1.png (No prompts!)
                $sourceImg = Join-Path $env:TEMP "$projectName_plate_1.png"
                $isTemp = $true

                # Protect against the archive not existing in TSV-Only mode
                if (Test-Path $InputFile) {
                    $archive = [System.IO.Compression.ZipFile]::OpenRead($InputFile)
                    $plateEntry = $archive.Entries | Where-Object { $_.FullName -replace '\\', '/' -match "(?i)Metadata/plate_1\.png$" } | Select-Object -First 1

                    if ($plateEntry) {
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($plateEntry, $sourceImg, $true)
                    } else {
                        $sourceImg = "" # Failsafe
                    }
                    $archive.Dispose()
                } else {
                    $sourceImg = ""
                }
            }

            if (Test-Path $sourceImg) {
                # Build Python arguments
                $pyArgs = @($pyScript, "--name", "`"$projectName`"", "--time", "`"$timeAdd`"", "--img", "`"$sourceImg`"", "--out", "`"$outImg`"", "--colors")
                foreach ($i in 1..4) {
                    if ($filData[$i].g -gt 0) {
                        $pyArgs += "`"$($filData[$i].color)|$($filData[$i].rawHex)|$($filData[$i].g)`""
                    }
                }

                $pyLog = Join-Path $env:TEMP "python_error.log"
                # Call Python to build the image
                $proc = Start-Process -FilePath "python" -ArgumentList $pyArgs -Wait -NoNewWindow -PassThru -RedirectStandardError $pyLog

                if (Test-Path $expectedPng) {
                    Write-Host "[DONE]" -ForegroundColor Green
                    Remove-Item $pyLog -Force -ErrorAction SilentlyContinue
                } else {
                    Write-Host "[FAILED]" -ForegroundColor Red
                    if (Test-Path $pyLog) {
                        Write-Host "      PYTHON ERROR:" -ForegroundColor Yellow
                        Get-Content $pyLog | Write-Host -ForegroundColor DarkRed
                        Remove-Item $pyLog -Force -ErrorAction SilentlyContinue
                    }
                }

                if ($isTemp) { Remove-Item $sourceImg -Force -ErrorAction SilentlyContinue }
            } else {
                Write-Host "[SKIPPED - No Image Found]" -ForegroundColor DarkGray
            }
        }
    } catch {
        Write-Host "[CRASHED]" -ForegroundColor Red
        Write-Host "     POWERSHELL EXCEPTION: $_" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------------
# 8. CONSOLE OUTPUT & TSV SAVING (Only runs during Full Extraction)
# ---------------------------------------------------------------------------------
if (-not $SkipExtraction) {
    Write-Host "`n--- Console Output Verification ---" -ForegroundColor Green
    Write-Host "Raw File Time:    $(if ($analyzer) { $analyzer.PrintTime } else { 'N/A' })"
    Write-Host "Single Obj Time:  $singlePrintTimeStr"
    Write-Host "Converted Time:   $h Hours, $m Minutes"
    Write-Host "Pre-Merge Count:  $objCount Objects"
    Write-Host "Model Filament:   $([math]::Round($modelGrams, 2))g"
    Write-Host "Color Swaps:      $actualColorSwaps"
    Write-Host "Added Time/Wig:   $timeAdd Minutes"
    Write-Host "-----------------------------------"

    if ($ConsoleOnly) {
        Write-Host "Data was NOT saved to the TSV files (ConsoleOnly flag used)." -ForegroundColor Yellow
    } else {
        $tsvLine = $outputValues -join "`t"

        # Fallback if no paths are provided
        if (-not $MasterTsvPath -and -not $IndividualTsvPath) {
            $MasterTsvPath = Join-Path (Split-Path $InputFile -Parent) "ExtractionResults.tsv"
        }

        # --- UPSERT LOGIC FOR MASTER TSV ---
        if ($MasterTsvPath) {
            if (Test-Path $MasterTsvPath) {
                $lines = @(Get-Content $MasterTsvPath)
                $found = $false

                for ($i = 0; $i -lt $lines.Count; $i++) {
                    if ($lines[$i] -match "^$([regex]::Escape($projectName))\t") {
                        $lines[$i] = $tsvLine
                        $found = $true
                        break
                    }
                }

                if (-not $found) { $lines += $tsvLine }
                Set-Content -Path $MasterTsvPath -Value $lines
            } else {
                Set-Content -Path $MasterTsvPath -Value $tsvLine
            }
            Write-Host "Success! Upserted to Master: $MasterTsvPath" -ForegroundColor Green
        }

        # --- STRICT OVERWRITE FOR INDIVIDUAL TSV ---
        if ($IndividualTsvPath) {
            Set-Content -Path $IndividualTsvPath -Value $tsvLine
            Write-Host "Success! Saved Individual: $IndividualTsvPath" -ForegroundColor Green
        }
    }
}