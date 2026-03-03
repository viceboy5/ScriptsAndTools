param (
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$InputFile,

    [switch]$ConsoleOnly
)

Write-Host "Processing $InputFile (Formatting for Google Sheets)..." -ForegroundColor Cyan

# Define the script directory to look for the color mapping CSV
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$colorCsvPath = Join-Path $scriptDir "colorNamesCSV.csv"

# --- 1. Define the Native C# Engine ---
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
                
                // Track Machine Movements
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
                // Track Metadata Blocks
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
                        int colIdx = line.IndexOf("time:");
                        if (colIdx > -1) {
                            string pt = line.Substring(colIdx + 5).Trim();
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
                // Track Extruder Modes & Color Changes
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

if (-not ("GcodeAnalyzer" -as [type])) {
    Add-Type -TypeDefinition $csharpCode -Language CSharp
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($InputFile)

    $totalGrams = 0.0
    $totalMeters = 0.0
    $objCount = 0
    
    # Initialize array for up to 4 filaments
    $filData = @(
        @{ g = 0; color = "" }, @{ g = 0; color = "" }, 
        @{ g = 0; color = "" }, @{ g = 0; color = "" }, @{ g = 0; color = "" } 
    )

    # --- 2. Extract Metadata (XML) ---
    $configEntry = $zipArchive.Entries | Where-Object { $_.FullName -replace '\\', '/' -match "Metadata/slice_info\.config$" }
    if ($configEntry) {
        $configStream = $configEntry.Open()
        $configReader = [System.IO.StreamReader]::new($configStream)
        $configContent = $configReader.ReadToEnd()
        $configReader.Close()
        
        try {
            [xml]$xml = $configContent
            
            # Extract Filament Weights and Colors
            foreach ($i in 1..4) {
                $node = $xml.SelectSingleNode("//filament[@id='$i']")
                if ($node) {
                    $weight = 0.0
                    [double]::TryParse($node.used_g, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$weight) | Out-Null
                    
                    if ($weight -gt 0) {
                        $filData[$i].g = [math]::Round($weight, 2)
                        $hexColor = $node.color
                        
                        if (Test-Path $colorCsvPath) {
                            $mappedName = Select-String -Path $colorCsvPath -Pattern "(?i)$hexColor" | Select-Object -First 1
                            if ($mappedName) {
                                $hexColor = $mappedName.Line.Split(',')[1].Trim()
                            }
                        }
                        $filData[$i].color = $hexColor
                    }
                }
            }

            # --- Pre-Merge Object Count Logic ---
            $objs = $xml.SelectNodes('//plate/object')
            if ($objs) {
                foreach ($o in $objs) {
                    $objName = $o.name
                    
                    # Ignore known text/version anomalies entirely
                    if ($objName -match '(?i)text|version') { continue }
                    
                    # Check for merged groups (e.g., "Object12") and grab the trailing number
                    if ($objName -match '(\d+)$') {
                        $objCount += [int]$matches[1]
                    } else {
                        $objCount += 1
                    }
                }
            }
        } catch {
            Write-Warning "Could not parse XML properly, some details may be blank."
        }

        # Calculate total density for the C# engine
        $gramMatches = [regex]::Matches($configContent, 'used_g="([0-9.]+)"')
        foreach ($match in $gramMatches) {
            $totalGrams += [double]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        $meterMatches = [regex]::Matches($configContent, 'used_m="([0-9.]+)"')
        foreach ($match in $meterMatches) {
            $totalMeters += [double]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }

    $gramsPerMm = 0.0
    if ($totalMeters -gt 0) {
        $gramsPerMm = $totalGrams / ($totalMeters * 1000)
    }

    # --- 3. Pass the torch to C# ---
    $gcodeEntry = $zipArchive.Entries | Where-Object { $_.FullName -like "*.gcode" } | Select-Object -First 1
    $analyzer = New-Object GcodeAnalyzer

    if ($gcodeEntry) {
        $gcodeStream = $gcodeEntry.Open()
        $analyzer.Analyze($gcodeStream, $gramsPerMm)
    }

    $zipArchive.Dispose()

    # --- 4. Crunch Final Math & Format ---
    $modelGrams = $totalGrams - $analyzer.FlushGrams - $analyzer.TowerGrams

    # Parse the Print Time into Days, Hours, and Minutes
    $d = 0; $h = 0; $m = 0
    if ($analyzer.PrintTime -match '(\d+)d') { $d = [int]$matches[1] }
    if ($analyzer.PrintTime -match '(\d+)h') { $h = [int]$matches[1] }
    if ($analyzer.PrintTime -match '(\d+)m') { $m = [int]$matches[1] }
    if ($analyzer.PrintTime -match '(\d+)s' -and [int]$matches[1] -ge 30) { $m++ }
    
    # Convert days to hours and add to the total
    $h += ($d * 24)

    if ($m -ge 60) { $m -= 60; $h++ }

    # Calculate Time Add/Wig
    $timeAdd = 0
    if ($objCount -gt 0) {
        $totalMinutes = ($h * 60) + $m
        $timeAdd = [math]::Round($totalMinutes / $objCount, 2)
    }

    # Build Array of Values directly in order
    $outputValues = @(
        ((Split-Path $InputFile -Leaf) -replace '\.gcode\.3mf$', ''),
        "", # Theme (Blank)
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
        "", # Replaced Waste Grams with a blank space for the checkbox
        $analyzer.ColorChanges,
        $objCount,
        $timeAdd,
        [math]::Round($modelGrams, 2)
    )

    $tsvPath = Join-Path (Split-Path $InputFile -Parent) "ExtractionResults.tsv"
    
    # Bypass Export-Csv entirely, output strictly as tab-separated values without quotes
    $tsvLine = $outputValues -join "`t"
    Add-Content -Path $tsvPath -Value $tsvLine

    Write-Host "`n--- Console Output Verification ---" -ForegroundColor Green
    Write-Host "Raw File Time:   $($analyzer.PrintTime)"
    Write-Host "Converted Time:  $h Hours, $m Minutes"
    Write-Host "Pre-Merge Count: $objCount Objects"
    Write-Host "Model Filament:  $([math]::Round($modelGrams, 2))g"
    Write-Host "-----------------------------------"

    if ($ConsoleOnly) {
        Write-Host "Data was NOT saved to the TSV file (ConsoleOnly flag used)." -ForegroundColor Yellow
    } else {
        $tsvPath = Join-Path (Split-Path $InputFile -Parent) "ExtractionResults.tsv"
        $tsvLine = $outputValues -join "`t"
        Add-Content -Path $tsvPath -Value $tsvLine
        Write-Host "Success! Formatted data appended to: $tsvPath" -ForegroundColor Green
    }

} catch {
    Write-Error "Failed to process .gcode.3mf archive: $_"
    if ($null -ne $zipArchive) { $zipArchive.Dispose() }
}