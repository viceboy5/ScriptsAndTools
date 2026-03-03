@echo off
TITLE 3MF Data Extractor

:: Check if a file was actually dropped
if "%~1"=="" (
    echo [ERROR] No file detected.
    echo Please drag and drop a .gcode.3mf file directly onto this batch file.
    echo.
    pause
    exit /b
)

echo [INFO] Passing file to PowerShell...
:: %~dp0 ensures it looks for the PS1 script in the exact same folder the .bat file lives in
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "%~1"

echo.
pause