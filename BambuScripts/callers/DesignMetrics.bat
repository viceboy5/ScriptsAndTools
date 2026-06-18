@echo off
setlocal enabledelayedexpansion
:: ============================================================
:: DesignMetrics.bat  -  per-design metrics readout
::
:: DRAG AND DROP onto this file:
::   - A design FOLDER, or a PARENT folder (searched recursively for every
::     design folder beneath it - any folder with a Full.gcode.3mf).
::   - You can drop several folders/files at once.
::
:: One design -> full readout.  Many -> a compact one-line-per-design summary.
:: This .bat lives in callers\ and calls workers\design_metrics_worker.py.
:: ============================================================

:: --- locate a real Python (the WindowsApps "python"/"py" aliases are dead stubs) ---
set "PYEXE="
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
if not defined PYEXE if exist "%LOCALAPPDATA%\Python\bin\python.exe" set "PYEXE=%LOCALAPPDATA%\Python\bin\python.exe"
if not defined PYEXE if exist "%ProgramFiles%\Python313\python.exe" set "PYEXE=%ProgramFiles%\Python313\python.exe"
if not defined PYEXE set "PYEXE=python"

if "%~1"=="" (
    echo.
    echo   Drag a design FOLDER onto this .bat file to see its metrics.
    echo   ^(using: !PYEXE!^)
    echo.
    pause
    exit /b
)

set "SCRIPT=%~dp0..\workers\design_metrics_worker.py"
echo.
"!PYEXE!" "!SCRIPT!" %*

echo.
pause
