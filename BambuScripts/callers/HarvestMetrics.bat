@echo off
setlocal enabledelayedexpansion
:: ============================================================
:: HarvestMetrics.bat  -  build the production metrics dataset
::
:: DRAG AND DROP onto this file:
::   - Any folder(s): searched RECURSIVELY for every design folder
::     (a folder with a Full.gcode.3mf). You can drop several at once.
::
:: Appends each NEW design's FULL metric set to:
::     BambuScripts\data\production_metrics.csv
:: Designs already in that CSV are skipped (parsed only once). To start
:: the file over, run the worker manually with --overwrite.
:: ============================================================

:: --- locate a real Python (the WindowsApps "python"/"py" aliases are dead stubs) ---
set "PYEXE="
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
if not defined PYEXE if exist "%LOCALAPPDATA%\Python\bin\python.exe" set "PYEXE=%LOCALAPPDATA%\Python\bin\python.exe"
if not defined PYEXE set "PYEXE=python"

if "%~1"=="" (
    echo.
    echo   Drag production folders onto this .bat to harvest them into
    echo   data\production_metrics.csv  - full gcode parse, this can take a while.
    echo.
    pause
    exit /b
)

set "SCRIPT=%~dp0..\workers\design_metrics_worker.py"
echo.
"!PYEXE!" "!SCRIPT!" %* --full --csv production_metrics.csv --select

echo.
pause
