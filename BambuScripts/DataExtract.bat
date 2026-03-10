@echo off
setlocal enabledelayedexpansion
TITLE 3MF Modular Data Extractor

:: Define Bambu Studio Path
set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"

:: 1. Check if a file or folder was actually dropped
if "%~1"=="" (
    echo [ERROR] No file or folder detected.
    pause
    exit /b
)

:: Strip trailing slash if present for consistent formatting
set "TARGET=%~1"
if "!TARGET:~-1!"=="\" set "TARGET=!TARGET:~0,-1!"

:: 2. Determine where the Master TSV should be saved
if exist "%TARGET%\" (
    set "MASTER_TSV=%TARGET%\Master_ExtractionResults.tsv"
    set "IS_FOLDER=1"
) else (
    set "MASTER_TSV=%~dp1Master_ExtractionResults.tsv"
    set "IS_FOLDER=0"
)

:: 3. Route to Processing Subroutine
if "!IS_FOLDER!"=="1" (
    echo [INFO] Folder detected. Searching recursively using DIR...
    echo Master TSV: !MASTER_TSV!
    echo.

    :: 'dir /s /b' bypasses the nasty bugs FOR /R has with spaces in variable paths
    for /f "delims=" %%F in ('dir /s /b "%TARGET%\*.gcode.3mf" 2^>nul') do (
        set "FILE_NAME=%%~nxF"

        :: Only process files that DO NOT contain "Final" (We only want the Full plates)
        if /I "!FILE_NAME:Final=!"=="!FILE_NAME!" (
            call :ProcessFile "%%~fF"
        )
    )
) else (
    echo [INFO] Single file detected.
    echo Master TSV: !MASTER_TSV!
    echo.
    call :ProcessFile "%TARGET%"
)

echo.
echo =====================================================================
echo All tasks complete! Master file updated.
pause
exit /b

:: ---------------------------------------------------------
:: Processing Subroutine
:: ---------------------------------------------------------
:ProcessFile
set "INPUT_GCODE=%~1"
set "INPUT_DIR=%~dp1"
set "INPUT_NAME=%~nx1"

echo ---------------------------------------------------------
echo Processing: !INPUT_NAME!

:: Deduce paths based on naming conventions (Full -> Final)
set "SINGLE_3MF_NAME=!INPUT_NAME:Full.gcode.3mf=Final.3mf!"
set "SINGLE_3MF_PATH=!INPUT_DIR!!SINGLE_3MF_NAME!"
set "SINGLE_GCODE_PATH=!INPUT_DIR!!INPUT_NAME:Full=Final!"

:: Check if the base Final.3mf exists to slice
if exist "!SINGLE_3MF_PATH!" (
    echo   [FOUND] Matching isolated file: !SINGLE_3MF_NAME!
    echo   [ACTION] Slicing for Time Add calculations... Please wait.

    :: Wait for Bambu Studio to slice
    "!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SINGLE_GCODE_PATH!" "!SINGLE_3MF_PATH!" >nul 2>&1

    if exist "!SINGLE_GCODE_PATH!" (
        echo   [SUCCESS] Slicing complete. Extracting data...
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!INPUT_GCODE!" -SingleFile "!SINGLE_GCODE_PATH!" -MasterTsvPath "!MASTER_TSV!"

        echo   [CLEANUP] Deleting temporary !SINGLE_GCODE_PATH!...
        del /f /q "!SINGLE_GCODE_PATH!"
    ) else (
        echo   [ERROR] Bambu Studio failed to generate the Final.gcode.3mf file.
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!INPUT_GCODE!" -MasterTsvPath "!MASTER_TSV!"
    )
) else (
    echo   [SKIP] No matching !SINGLE_3MF_NAME! found in directory.
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!INPUT_GCODE!" -MasterTsvPath "!MASTER_TSV!"
)
exit /b