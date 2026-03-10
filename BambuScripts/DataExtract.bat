@echo off
setlocal
TITLE 3MF Data Extractor

:: 1. Check if a file or folder was actually dropped
if "%~1"=="" (
    echo [ERROR] No file or folder detected.
    echo Please drag and drop a .gcode.3mf file or a master folder onto this batch file.
    echo.
    pause
    exit /b
)

set "TARGET=%~1"

:: 2. Determine where the Master TSV should be saved
if exist "%TARGET%\" (
    :: If it's a folder, save it at the root of that folder
    set "MASTER_TSV=%~f1\Master_ExtractionResults.tsv"
) else (
    :: If it's a file, save it in the same directory as the file
    set "MASTER_TSV=%~dp1Master_ExtractionResults.tsv"
)

:: 3. Process the file or folder
if exist "%TARGET%\" (
    echo [INFO] Folder detected. Searching recursively for .gcode.3mf files...
    echo Master file will be saved to: %MASTER_TSV%
    echo.

    :: Recursively search for any file ending in .gcode.3mf
    for /r "%TARGET%" %%F in (*.gcode.3mf) do (
        echo ---------------------------------------------------------
        echo Processing: %%~nxF
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "%%~fF" -MasterTsvPath "%MASTER_TSV%"
    )
) else (
    echo [INFO] Single file detected. Passing to PowerShell...
    echo Master file will be saved to: %MASTER_TSV%
    echo ---------------------------------------------------------
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "%TARGET%" -MasterTsvPath "%MASTER_TSV%"
)

echo.
echo =====================================================================
echo All tasks complete! Master file updated.
pause