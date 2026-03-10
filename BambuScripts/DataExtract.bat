@echo off
setlocal enabledelayedexpansion
TITLE 3MF Modular Data Extractor

:: Define Bambu Studio Path
set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"

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
if exist "!TARGET!\" (
    set "MASTER_TSV=%~f1\Master_ExtractionResults.tsv"
) else (
    set "MASTER_TSV=%~dp1Master_ExtractionResults.tsv"
)

:: 3. Route to Processing Subroutine
if exist "!TARGET!\" (
    echo [INFO] Folder detected. Searching recursively for .gcode.3mf files...
    echo Master TSV: !MASTER_TSV!
    echo.
    for /r "!TARGET!" %%F in (*.gcode.3mf) do (
        set "FILE_NAME=%%~nxF"

        :: Skip processing if it's a leftover "Final" file
        if /I "!FILE_NAME:Final=!"=="!FILE_NAME!" (
            call :ProcessFile "%%~fF"
        )
    )
) else (
    echo [INFO] Single file detected.
    echo Master TSV: !MASTER_TSV!
    echo.
    call :ProcessFile "!TARGET!"
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
    echo   [ACTION] Slicing !SINGLE_3MF_NAME! for Time Add calculations...
    "!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SINGLE_GCODE_PATH!" "!SINGLE_3MF_PATH!" >nul 2>&1

    if exist "!SINGLE_GCODE_PATH!" (
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!INPUT_GCODE!" -SingleFile "!SINGLE_GCODE_PATH!" -MasterTsvPath "!MASTER_TSV!"

        echo   [CLEANUP] Deleting temporary !SINGLE_GCODE_PATH!...
        del /f /q "!SINGLE_GCODE_PATH!"
    ) else (
        echo   [ERROR] Slicing failed. Proceeding without single-object Time Add data.
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!INPUT_GCODE!" -MasterTsvPath "!MASTER_TSV!"
    )
) else (
    echo   [SKIP] No matching Final.3mf found. Proceeding without single-object Time Add data.
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!INPUT_GCODE!" -MasterTsvPath "!MASTER_TSV!"
)
exit /b