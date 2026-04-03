@echo off
setlocal

:: Check if a file/folder was dropped
if "%~1"=="" (
    echo [ERROR] Please drag and drop 3MF files or a folder onto this script.
    pause
    exit /b
)

echo Launching Viewer...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\View_Picks_Worker.ps1" %*

echo.
pause