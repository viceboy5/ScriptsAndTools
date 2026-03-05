@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo [ERROR] Please drag and drop a processed file ^(like _Full.3mf or _Nest.3mf^) onto this script.
    pause
    exit /b 1
)

echo ==========================================
echo          3MF PIPELINE REVERTER
echo ==========================================
echo.

:process_loop
if "%~1"=="" goto finish

set "INPUTDIR=%~dp1"
set "INPUTBASE=%~n1"

REM Extract the core project name by stripping any pipeline suffixes
set "CORE=!INPUTBASE:_Full=!"
set "CORE=!CORE:_Nest=!"
set "CORE=!CORE:_Final=!"

echo [ Reverting Project: !CORE! ]

REM 1. Delete the merged and sliced files
if exist "!INPUTDIR!!CORE!_Full.gcode.3mf" (
    echo   [-] Deleting !CORE!_Full.gcode.3mf...
    del /f /q "!INPUTDIR!!CORE!_Full.gcode.3mf"
)

if exist "!INPUTDIR!!CORE!_Full.3mf" (
    echo   [-] Deleting merged !CORE!_Full.3mf...
    del /f /q "!INPUTDIR!!CORE!_Full.3mf"
)

REM Delete the individual TSV data file
if exist "!INPUTDIR!!CORE!_Full_Data.tsv" (
    echo   [-] Deleting !CORE!_Full_Data.tsv...
    del /f /q "!INPUTDIR!!CORE!_Full_Data.tsv"
)

REM 2. Rename Nest back to Full
if exist "!INPUTDIR!!CORE!_Nest.3mf" (
    echo   [+] Restoring !CORE!_Nest.3mf back to !CORE!_Full.3mf...
    ren "!INPUTDIR!!CORE!_Nest.3mf" "!CORE!_Full.3mf"
) else (
    echo   [!] Could not find !CORE!_Nest.3mf to restore.
)

echo.
shift
goto process_loop

:finish
echo ==========================================
echo Revert complete! Your Full.3mf is restored and Final files were kept.
pause