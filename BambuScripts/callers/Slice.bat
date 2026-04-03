@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo Usage: Drag and drop one or more .3mf files or folders onto this script.
    pause
    exit /b 1
)

set "WORKER=%~dp0..\workers\slicer_automation_worker.ps1"

echo --------------------------------------------------
echo BAMBU STUDIO BATCH SLICER
echo --------------------------------------------------
echo.

:: Collect all .3mf files from every dropped item (file or folder)
set "COUNT=0"

:collect_loop
if "%~1"=="" goto run_slice

if exist "%~1\" (
    echo [ Scanning folder: %~nx1 ]
    for /R "%~1" %%F in (*.3mf) do (
        echo "%%F" | findstr /i /v "\.gcode\.3mf" >nul
        if not errorlevel 1 (
            set /a COUNT+=1
            set "FILE_!COUNT!=%%F"
        )
    )
) else (
    echo "%~1" | findstr /i /v "\.gcode\.3mf" >nul
    if not errorlevel 1 (
        set /a COUNT+=1
        set "FILE_!COUNT!=%~1"
    )
)

shift
goto collect_loop

:run_slice
if %COUNT%==0 (
    echo [!] No valid .3mf files found to slice.
    pause
    exit /b 1
)

echo [ Found %COUNT% file(s) to slice ]
echo.

:: Build PowerShell array string
set "PS_PATHS="
for /L %%I in (1,1,%COUNT%) do (
    if defined PS_PATHS (
        set "PS_PATHS=!PS_PATHS!, '!FILE_%%I!'"
    ) else (
        set "PS_PATHS='!FILE_%%I!'"
    )
)

powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& '%WORKER%' -InputPaths @(%PS_PATHS%)"

echo.
echo --------------------------------------------------
echo All done!
pause