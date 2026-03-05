@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo [ERROR] Please drag and drop a processed file or a folder onto this script.
    pause
    exit /b 1
)

echo ==========================================
echo          3MF PIPELINE REVERTER
echo ==========================================
echo.

:process_loop
if "%~1"=="" goto finish

REM Check if the dropped item is a folder (the trailing \ tests for a directory)
if exist "%~1\" (
    echo [ Scanning Directory and Subfolders: %~nx1 ]
    echo.
    REM Recursively hunt for any _Nest.3mf files in this folder and below
    for /R "%~1" %%F in (*_Nest.3mf) do (
        call :revert_target "%%F"
    )
) else (
    REM It is just a standard file
    call :revert_target "%~1"
)

shift
goto process_loop

:revert_target
set "FILE_PATH=%~1"
set "INPUTDIR=%~dp1"
set "INPUTBASE=%~n1"

REM Extract the core project name
set "CORE=!INPUTBASE:_Full=!"
set "CORE=!CORE:_Nest=!"
set "CORE=!CORE:_Final=!"

echo [ Reverting: !CORE! ]

REM 1. Delete the generated junk
if exist "!INPUTDIR!!CORE!_Full.gcode.3mf" (
    echo   [-] Deleting !CORE!_Full.gcode.3mf...
    del /f /q "!INPUTDIR!!CORE!_Full.gcode.3mf"
)

if exist "!INPUTDIR!!CORE!_Full.3mf" (
    echo   [-] Deleting generated !CORE!_Full.3mf...
    del /f /q "!INPUTDIR!!CORE!_Full.3mf"
)

if exist "!INPUTDIR!!CORE!_Full_Data.tsv" (
    echo   [-] Deleting !CORE!_Full_Data.tsv...
    del /f /q "!INPUTDIR!!CORE!_Full_Data.tsv"
)

REM 2. Restore the original file
if exist "!INPUTDIR!!CORE!_Nest.3mf" (
    echo   [+] Restoring !CORE!_Nest.3mf to !CORE!_Full.3mf...
    ren "!INPUTDIR!!CORE!_Nest.3mf" "!CORE!_Full.3mf"
) else (
    echo   [!] Could not find !CORE!_Nest.3mf to restore.
)
echo.
exit /b

:finish
echo ==========================================
echo Revert complete! Your Full.3mf files are restored.
pause