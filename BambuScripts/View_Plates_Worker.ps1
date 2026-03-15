$ErrorActionPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a clean temporary directory for the images
$tempDir = Join-Path $env:TEMP "BambuPlates"

# Clean out images from previous runs so you only see the current batch
if (Test-Path $tempDir) { Remove-Item -Path "$tempDir\*" -Force -Recurse }
if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }

$filesToProcess = @()

# Sort out the dragged-and-dropped items
foreach ($path in $args) {
    if (Test-Path -LiteralPath $path -PathType Container) {
        Write-Host "Scanning folder for *Full.gcode.3mf files..." -ForegroundColor Cyan
        # If it's a folder, ONLY grab the Full.gcode.3mf files
        $filesToProcess += @(Get-ChildItem -LiteralPath $path -Filter "*Full.gcode.3mf" -File -Recurse)
    } elseif (Test-Path -LiteralPath $path -PathType Leaf) {
        # If it's a direct file, process it regardless of name
        $filesToProcess += Get-Item -LiteralPath $path
    }
}

if ($filesToProcess.Count -eq 0) {
    Write-Host "No matching files found to process." -ForegroundColor Yellow
    return
}

$extractedCount = 0

# Process the files
foreach ($file in $filesToProcess) {
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)

        # Search the archive for plate_1.png (case insensitive)
        $entry = $zip.Entries | Where-Object { $_.FullName -match '(?i)plate_1\.png$' }

        if ($entry) {
            # Name the image after the 3MF file so you know which is which
            $safeName = $file.BaseName + "_plate.png"
            $outDest = Join-Path $tempDir $safeName

            # If by some chance the exact same filename exists, append a random string
            if (Test-Path $outDest) {
                $safeName = $file.BaseName + "_" + [guid]::NewGuid().ToString().Substring(0,4) + "_plate.png"
                $outDest = Join-Path $tempDir $safeName
            }

            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $outDest, $true)

            Write-Host "Extracted: $safeName" -ForegroundColor Green
            $extractedCount++
        } else {
            Write-Host "Skipped: No plate_1.png found in $($file.Name)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "Error reading $($file.Name)" -ForegroundColor Red
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

# Open the folder once at the end if we successfully extracted anything
if ($extractedCount -gt 0) {
    Write-Host "`nOpening folder containing $extractedCount images..." -ForegroundColor Cyan
    Invoke-Item $tempDir
}