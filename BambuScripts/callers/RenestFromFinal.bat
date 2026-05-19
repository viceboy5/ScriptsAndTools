@echo off
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo.
    echo   RenestFromFinal
    echo   ---------------------------------------------------------------
    echo   Drag and drop one or more Final.3mf files onto this script.
    echo.
    echo   For each file the script will automatically locate a sibling
    echo   _Nest.3mf or _Full.3mf in the same folder to read plate
    echo   transforms from, then produce a _Renest.3mf alongside it.
    echo.
    echo   Multiple files are confirmed once, then processed in sequence.
    echo.
    pause
    exit /b
)

set "worker=%~dp0..\workers\RenestFromFinal_worker.ps1"

:: ---- Collect all dropped file paths ----------------------------------------
:: Shift-loop handles quoted paths with spaces correctly.
set /a filecount=0
:collect
if "%~1"=="" goto collected
set /a filecount+=1
set "file!filecount!=%~1"
shift
goto collect
:collected

:: ---- Single file: let the worker show its own prompt -----------------------
if !filecount! equ 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%worker%" -FinalPath "!file1!"
    echo.
    pause
    exit /b
)

:: ---- Multiple files: show a batch summary and confirm once -----------------
echo.
echo ================================================================
echo   RenestFromFinal  ^|  Batch Mode  ^|  !filecount! files queued
echo ================================================================
echo.
for /l %%I in (1,1,!filecount!) do echo   %%I. !file%%I!
echo.
echo   Each Final will be matched to its sibling Nest or Full file.
echo   Output will be saved as _Renest.3mf alongside each Final.
echo.
set /p batchconfirm=  Proceed with all !filecount! files? [Y/N]:
if /i "!batchconfirm!" neq "y" (
    echo.
    echo   Cancelled.
    echo.
    pause
    exit /b
)
echo.

:: ---- Process each file in sequence -----------------------------------------
set /a done=0
set /a failed=0
for /l %%I in (1,1,!filecount!) do (
    echo ----------------------------------------------------------------
    echo   File %%I of !filecount!: !file%%I!
    echo ----------------------------------------------------------------
    powershell -NoProfile -ExecutionPolicy Bypass -File "%worker%" -FinalPath "!file%%I!" -NoConfirm
    if !errorlevel! equ 0 (set /a done+=1) else (set /a failed+=1 & echo   WARNING: worker reported an error for file %%I.)
    echo.
)

echo ================================================================
if !failed! equ 0 (
    echo   All !filecount! files completed successfully.
) else (
    echo   Completed: !done! succeeded,  !failed! failed.
)
echo ================================================================
echo.
pause
