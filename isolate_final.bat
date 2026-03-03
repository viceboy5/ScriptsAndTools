@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo [ERROR] Please drag and drop a .3mf file onto this batch script.
    pause
    exit /b 1
)

:: --- Set up File Paths ---
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTBASE=%~n1"

:: Clean up the name. If it ends in 'Full', change it to 'Final'. Otherwise, just append '_Final'.
set "FINALBASE=!INPUTBASE:Full=Final!"
if "!FINALBASE!"=="!INPUTBASE!" set "FINALBASE=!INPUTBASE!_Final"

set "FINAL_PATH=!INPUTDIR!!FINALBASE!.3mf"
set "WORK_DIR=%TEMP%\isolate_work_%RANDOM%"

echo =======================================================
echo   SINGLE OBJECT ISOLATION TEST
echo =======================================================
echo Target: !INPUTBASE!
echo.

:: --- Step 1: Unzip ---
echo [1/2] Unzipping original archive...
mkdir "!WORK_DIR!" 2>nul
powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUT!', '!WORK_DIR!')" >nul 2>&1

:: --- Step 2: Isolate Center Object ---
echo [2/2] Isolating center object and generating Final.3mf...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0isolate_final_worker.ps1" -WorkDir "!WORK_DIR!" -OutputPath "!FINAL_PATH!"

if not exist "!FINAL_PATH!" (
    echo   [!] ERROR: Failed to generate Final.3mf.
    goto cleanup
)

:cleanup
echo.
echo Cleaning up temp folders...
rmdir /s /q "!WORK_DIR!" 2>nul

echo =======================================================
echo Finished! !FINALBASE!.3mf has been created successfully.
pause