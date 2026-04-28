@echo off
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo Please drag and drop files or folders onto this script.
    pause
    exit /b
)

:: Capture script directory before any shift moves %0
set "scriptdir=%~dp0"

:: Write all dropped paths to a temp file to avoid quoting issues
set "tmpfile=%TEMP%\convert_high_inputs.txt"
if exist "%tmpfile%" del "%tmpfile%"

:collect
if "%~1"=="" goto run
echo %~1>>"%tmpfile%"
shift /1
goto :collect

:run
powershell -NoProfile -ExecutionPolicy Bypass -File "%scriptdir%..\workers\ConvertToHighRes.ps1" -InputFile "%tmpfile%"
pause
