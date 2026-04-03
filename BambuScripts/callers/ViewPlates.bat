@echo off
setlocal

:: Check if a file/folder was dropped
if "%~1"=="" (
    echo [ERROR] Please drag and drop 3MF files or a folder onto this script.
    pause
    exit /b
)

echo Launching Plate Viewer...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\ViewPlates_worker.ps1" %*

echo.
pause