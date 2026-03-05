@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo [ERROR] Please drag and drop a .3mf file or folder onto this batch script.
    pause
    exit /b 1
)

echo =======================================================
echo   SINGLE OBJECT ISOLATION TEST
echo =======================================================
echo.

:process_loop
if "%~1"=="" goto finish

if exist "%~1\" (
    echo [ Scanning Directory: %~nx1 ]
    for /R "%~1" %%F in (*.3mf) do (
        REM Skip files that are already Final or Gcode
        echo "%%F" | findstr /i /v "Final\.3mf \.gcode\.3mf" >nul
        if not errorlevel 1 call :isolate_target "%%F"
    )
) else (
    call :isolate_target "%~1"
)

shift
goto process_loop

:isolate_target
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTBASE=%~n1"

:: Clean up the name
set "FINALBASE=!INPUTBASE:Full=Final!"
if "!FINALBASE!"=="!INPUTBASE!" set "FINALBASE=!INPUTBASE!_Final"

set "FINAL_PATH=!INPUTDIR!!FINALBASE!.3mf"
set "WORK_DIR=%TEMP%\isolate_work_%RANDOM%"

echo Target: !INPUTBASE!

:: --- Step 1: Unzip ---
echo   [1/2] Unzipping original archive...
mkdir "!WORK_DIR!" 2>nul
powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUT!', '!WORK_DIR!')" >nul 2>&1

:: --- Step 2: Isolate Center Object ---
echo   [2/2] Isolating center object and generating Final.3mf...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0isolate_final_worker.ps1" -WorkDir "!WORK_DIR!" -OutputPath "!FINAL_PATH!"

if not exist "!FINAL_PATH!" (
    echo   [!] ERROR: Failed to generate Final.3mf.
) else (
    echo   [+] Success: !FINALBASE!.3mf created.
)

:: Cleanup
rmdir /s /q "!WORK_DIR!" 2>nul
echo.
exit /b

:finish
echo =======================================================
echo Finished! All files processed.
pause