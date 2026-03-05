@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo Usage: Drag and drop a .3mf file or folder onto this script.
    pause
    exit /b 1
)

set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"

echo --------------------------------------------------
echo BAMBU STUDIO LIVE-FEED SLICER
echo --------------------------------------------------
echo [ Ensuring background instances are closed... ]
taskkill /im bambu-studio.exe /f >nul 2>&1
timeout /t 2 >nul
echo --------------------------------------------------
echo.

:process_loop
if "%~1"=="" goto finish

if exist "%~1\" (
    echo [ Scanning Directory: %~nx1 ]
    for /R "%~1" %%F in (*.3mf) do (
        REM Prevent slicing files that are already sliced
        echo "%%F" | findstr /i /v "\.gcode\.3mf" >nul
        if not errorlevel 1 call :slice_target "%%F"
    )
) else (
    call :slice_target "%~1"
)

shift
goto process_loop

:slice_target
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTBASE=%~n1"
set "OUTPUT=!INPUTDIR!!INPUTBASE!.gcode.3mf"

echo Target File : !INPUT!
echo Slicing Plate 1 (Bypassing safety checks)...

"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!OUTPUT!" "!INPUT!" | findstr "^"
echo.
exit /b

:finish
echo --------------------------------------------------
echo Slicing complete!
pause