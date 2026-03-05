@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo Usage: Drag and drop a single .3mf file onto this script.
    pause
    exit /b 1
)

set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTBASE=%~n1"
set "OUTPUT=!INPUTDIR!!INPUTBASE!.gcode.3mf"

set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"

echo --------------------------------------------------
echo BAMBU STUDIO LIVE-FEED SLICER
echo --------------------------------------------------
echo Target File : !INPUT!
echo Output      : !OUTPUT!
echo --------------------------------------------------
echo.

REM 1. Clear out any stuck background instances
echo [1/2] Ensuring background instances are closed...
taskkill /im bambu-studio.exe /f >nul 2>&1
timeout /t 2 >nul

REM 2. Run the slicer and pipe the text directly to the console
echo.
echo [2/2] Slicing Plate 1 (Bypassing safety checks)...
echo --------------------------------------------------

REM --no-check forces the slicer to ignore boundary/conflict warnings.
REM --min-save strips out the 3D models so it acts as a read-only sliced file.
"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!OUTPUT!" "!INPUT!" | findstr "^"

echo --------------------------------------------------
echo.
echo Slicing stream closed.
echo.

if exist "!OUTPUT!" (
    echo SUCCESS: Sliced file was created!
) else (
    echo FAILED: The output file is missing. Read the log above to see why.
)
echo --------------------------------------------------
pause
endlocal