param (
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$InputFile,

    [string]$SingleFile = "",

    [switch]$ConsoleOnly,

    [string]$MasterTsvPath = "",
    [string]$IndividualTsvPath = ""
)

if ($ConsoleOnly) {
    Write-Host "Processing $InputFile (Console Output Only)..." -ForegroundColor Cyan
} else {
    Write-Host "Processing $InputFile (Formatting for Google Sheets)..." -ForegroundColor Cyan
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# --- 1. Load the Reverse-Lookup Color Dictionary (BULLETPROOF EDITION) ---
$LibraryNames = @{}
if (Test-Path $colorCsvPath) {
    # Use raw Get-Content so we don't trip over Import-Csv header injection or quoting bugs
    $csvLines = Get-Content -Path $colorCsvPath
    foreach ($line in $csvLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split ','
        if ($parts.Count -ge 2) {
            $rawHex = $parts[0].Replace('"','').Trim().ToUpper()
            $name = $parts[1].Replace('"','').Trim()

            # Skip the header row if it exists
            if ($name -match '(?i)^name$' -or $name -eq "N/A" -or $name -eq "") { continue }

            # Strip the '#' symbol so we are comparing pure alphanumeric strings
            if ($rawHex.StartsWith('#')) { $rawHex = $rawHex.Substring(1) }

            # Force 8-character (Alpha) compliance
            if ($rawHex.Length -eq 6) { $rawHex += "FF" }

            if ($rawHex.Length -eq 8) {
                $LibraryNames[$rawHex] = $name
            }
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

try {
    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($InputFile)

    $totalGrams = 0.0; $totalMeters = 0.0; $objCount = 0

    $filData = @(
        @{ g = 0; color = "" }, @{ g = 0; color = "" },
        @{ g = 0; color = "" }, @{ g = 0; color = "" }, @{ g = 0; color = "" }
    )

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

                        if ($LibraryNames.Contains($rawHex)) {
                            $filData[$i].color = $LibraryNames[$rawHex]
                        } else {
                            # Fallback to the raw hex with the # put back so it looks normal in Sheets
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

                    # Only parse numbers that explicitly follow our safe "MergedGroup_" tag
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
    $timeAdd = 0
    $singlePrintTimeStr = "N/A"

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

    $projectName = ((Split-Path $InputFile -Leaf) -replace '\.gcode\.3mf$', '')

    $outputValues = @(
        $projectName,
        "",
        (Get-Date).ToString("M/d/yyyy"),
        $h,
        $m,
        $(if ($filData[1].g -gt 0) { $filData[1].g } else { 0 }),
        $filData[1].color,
        $(if ($filData[2].g -gt 0) { $filData[2].g } else { 0 }),
        $filData[2].color,
        $(if ($filData[3].g -gt 0) { $filData[3].g } else { 0 }),
        $filData[3].color,
        $(if ($filData[4].g -gt 0) { $filData[4].g } else { 0 }),
        $filData[4].color,
        "",
        $actualColorSwaps,
        $objCount,
        [math]::Round($modelGrams, 2),
        "",
        $timeAdd
    )

    Write-Host "`n--- Console Output Verification ---" -ForegroundColor Green
    Write-Host "Raw File Time:    $($analyzer.PrintTime)"
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
                # Read all existing lines
                $lines = @(Get-Content $MasterTsvPath)
                $found = $false

                # Scan for an exact match of the Project Name in the first column
                for ($i = 0; $i -lt $lines.Count; $i++) {
                    if ($lines[$i] -match "^$([regex]::Escape($projectName))\t") {
                        $lines[$i] = $tsvLine
                        $found = $true
                        break
                    }
                }

                # If not found, append to the bottom
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

} catch {
    Write-Error "Failed to process .gcode.3mf archive: $_"
    if ($null -ne $zipArchive) { $zipArchive.Dispose() }
}