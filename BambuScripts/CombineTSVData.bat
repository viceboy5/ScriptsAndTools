@echo off
title TSV Combiner
color 0B

:: Check if a folder was dropped
if "%~1"=="" (
    echo ===================================================
    echo ERROR: No folder provided!
    echo Please drag and drop a folder directly onto this .bat file.
    echo ===================================================
    pause
    exit /b
)

:: Ensure the dropped item is a directory
if not exist "%~1\" (
    echo ===================================================
    echo ERROR: The dropped item is not a folder.
    echo Please drag and drop a FOLDER, not a file.
    echo ===================================================
    pause
    exit /b
)

set "TARGET_DIR=%~1"
set "OUTPUT_FILE=%~1\Combined_Data.tsv"

echo Scanning for .tsv files in: %TARGET_DIR%
echo ===================================================

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$folder = '%TARGET_DIR%'; " ^
    "$output = '%OUTPUT_FILE%'; " ^
    "$hasHeaders = $false; " ^
    "$files = Get-ChildItem -Path $folder -Filter '*.tsv' -Recurse | Where-Object { $_.FullName -ne $output } | Sort-Object FullName; " ^
    "if ($files.Count -eq 0) { Write-Host 'No .tsv files found in the folder or subfolders.' -ForegroundColor Yellow; exit }; " ^
    "Write-Host \"Found $($files.Count) .tsv files. Merging sequentially...`n\" -ForegroundColor Cyan; " ^
    "$outStream = [System.IO.StreamWriter]::new($output, $false, [System.Text.Encoding]::UTF8); " ^
    "$headerWritten = $false; " ^
    "$totalLines = 0; " ^
    "foreach ($file in $files) { " ^
    "    $lines = @(Get-Content -Path $file.FullName); " ^
    "    if ($lines.Count -eq 0) { Write-Host \"  [EMPTY]  $($file.Name)\" -ForegroundColor DarkGray; continue }; " ^
    "    $startIndex = 0; " ^
    "    if ($hasHeaders) { " ^
    "        if (-not $headerWritten) { " ^
    "            $outStream.WriteLine($lines[0]); " ^
    "            $headerWritten = $true; " ^
    "            $totalLines++; " ^
    "        } " ^
    "        $startIndex = 1; " ^
    "    } " ^
    "    $linesAdded = 0; " ^
    "    for ($i = $startIndex; $i -lt $lines.Count; $i++) { " ^
    "        $outStream.WriteLine($lines[$i]); " ^
    "        $linesAdded++; " ^
    "        $totalLines++; " ^
    "    } " ^
    "    Write-Host \"  [+$linesAdded lines] $($file.Name)\" -ForegroundColor Gray; " ^
    "} " ^
    "$outStream.Dispose(); " ^
    "Write-Host \"`n===================================================\" -ForegroundColor Cyan; " ^
    "Write-Host \"Success! Wrote $totalLines total lines to:\" -ForegroundColor Green; " ^
    "Write-Host $output -ForegroundColor White; "

echo.
pause