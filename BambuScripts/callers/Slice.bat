@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo Usage: Drag and drop one or more .3mf files or folders onto this script.
    pause
    exit /b 1
)

set "WORKER=%~dp0..\workers\Slice_worker.ps1"

echo --------------------------------------------------
echo BAMBU STUDIO BATCH SLICER
echo --------------------------------------------------
echo.

:: Collect dropped items (files and/or folders) as a PS array.
:: Folders are passed as-is and expanded inside Slice_worker.ps1 - this avoids
:: cmd.exe's ~8191 char command-line limit when a folder contains many files.
set "COUNT=0"
set "PS_PATHS="

:collect_loop
if "%~1"=="" goto run_slice

set /a COUNT+=1
if defined PS_PATHS (
    set "PS_PATHS=!PS_PATHS!, '%~1'"
) else (
    set "PS_PATHS='%~1'"
)

shift
goto collect_loop

:run_slice
if %COUNT%==0 (
    echo [!] Nothing to slice.
    pause
    exit /b 1
)

echo [ Passing %COUNT% item(s) to slicer ]
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& '%WORKER%' -InputPaths @(%PS_PATHS%)"

echo.
echo --------------------------------------------------
echo All done!
pause
