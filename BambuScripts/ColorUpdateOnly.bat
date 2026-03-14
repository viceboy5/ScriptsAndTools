@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo STANDALONE COLOR UPDATE WORKER
echo ===================================================

:: Check if a folder was dropped onto the script
if "%~1"=="" (
    echo [ERROR] Please drag and drop a MASTER FOLDER onto this script.
    pause
    exit /b
)

:: FIX: Standard variable expansion for the path string
set "masterDir=%~1"
if "%masterDir:~-1%"=="\" set "masterDir=%masterDir:~0,-1%"

:: Get the directory where this .bat file lives to reliably find the .ps1 worker
set "scriptDir=%~dp0"

echo Scanning for *Full.3mf files in: %masterDir%
echo ---------------------------------------------------

:: FIX: Use %masterDir% here so the FOR loop can actually read the path
for /r "%masterDir%" %%F in (*Full.3mf) do (
    set "targetFile=%%~fF"
    set "fileName=%%~nxF"

    :: Failsafe: Ignore anything that is already sliced (.gcode.3mf)
    echo !fileName! | findstr /i "\.gcode\.3mf$" >nul
    if errorlevel 1 (
        echo.
        echo === Checking Colors: !fileName! ===

        :: 1. Create a unique temporary directory
        set "localTemp=%TEMP%\ColorCheck_!RANDOM!"
        mkdir "!localTemp!"

        :: 2. Extract the .3mf to the temp folder using Windows native tar
        tar.exe -xf "!targetFile!" -C "!localTemp!" >nul 2>&1

        if errorlevel 1 (
            echo [ERROR] Failed to extract !fileName! - File might be in use.
        ) else (
            :: 3. Call the PowerShell worker
            powershell.exe -NoProfile -ExecutionPolicy Bypass -File "!scriptDir!update_colors_worker.ps1" -WorkDir "!localTemp!" -FileName "!fileName!" -OriginalZip "!targetFile!"
        )

        :: 4. Clean up the temporary reading directory
        rd /s /q "!localTemp!"
    )
)

echo.
echo ===================================================
echo All color updates complete!
pause
exit /b