$ErrorActionPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a clean temporary directory for the images
$tempDir = Join-Path $env:TEMP "BambuPicks"
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

# Process the files
foreach ($file in $filesToProcess) {
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)

        # Search the archive for pick_1.png (case insensitive)
        $entry = $zip.Entries | Where-Object { $_.FullName -match '(?i)pick_1\.png$' }

        if ($entry) {
            # Add a random 4-character suffix so Windows Photo Viewer doesn't lock the file on repeat runs
            $safeName = $file.BaseName + "_" + [guid]::NewGuid().ToString().Substring(0,4) + "_pick.png"
            $outDest = Join-Path $tempDir $safeName

            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $outDest, $true)

            Write-Host "Opened: $safeName" -ForegroundColor Green
            # This triggers Windows to instantly open the image in your default photo viewer
            Invoke-Item $outDest
        } else {
            Write-Host "Skipped: No pick_1.png found in $($file.Name)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "Error reading $($file.Name)" -ForegroundColor Red
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}