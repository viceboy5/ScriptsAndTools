@echo off
setlocal EnableDelayedExpansion

set "PAUSE_AT_END=1"
if "!WORKER_MODE!"=="1" set "PAUSE_AT_END=0"

if "%~1"=="" (
    echo [ERROR] Please drag and drop a processed file or a folder onto this script.
    if !PAUSE_AT_END!==1 pause
    exit /b 1
)

if !PAUSE_AT_END!==1 (
    echo ==========================================
    echo          3MF PIPELINE REVERTER
    echo ==========================================
    echo.
)

:process_loop
if "%~1"=="" goto finish

REM Check if the dropped item is a folder (the trailing \ tests for a directory)
if exist "%~1\" (
    if !PAUSE_AT_END!==1 echo [ Scanning Directory and Subfolders: %~nx1 ] & echo.
    REM Using *Nest.3mf acts as a wildcard, catching _Nest, .Nest, and " Nest"
    for /R "%~1" %%F in (*Nest.3mf) do (
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

REM Dynamically extract the separator (e.g., _, ., or Space) and the core name
set "SEP=!INPUTBASE:~-5,1!"
set "BASE_PREFIX=!INPUTBASE:~0,-5!"

REM Rebuild the file names using the detected separator
set "FULLBASE=!BASE_PREFIX!!SEP!Full"
set "NESTBASE=!BASE_PREFIX!!SEP!Nest"
set "FINALBASE=!BASE_PREFIX!!SEP!Final"

echo [ Reverting: !FULLBASE! ]

REM 1. Delete the generated junk files
if exist "!INPUTDIR!!FULLBASE!.gcode.3mf" (
    echo   [-] Deleting !FULLBASE!.gcode.3mf...
    del /f /q "!INPUTDIR!!FULLBASE!.gcode.3mf"
)

if exist "!INPUTDIR!!FULLBASE!.3mf" (
    echo   [-] Deleting generated !FULLBASE!.3mf...
    del /f /q "!INPUTDIR!!FULLBASE!.3mf"
)

if exist "!INPUTDIR!!FULLBASE!_Data.tsv" (
    echo   [-] Deleting !FULLBASE!_Data.tsv...
    del /f /q "!INPUTDIR!!FULLBASE!_Data.tsv"
)

if exist "!INPUTDIR!!FINALBASE!.3mf" (
    echo   [-] Deleting !FINALBASE!.3mf...
    del /f /q "!INPUTDIR!!FINALBASE!.3mf"
)

if exist "!INPUTDIR!!FINALBASE!.gcode.3mf" (
    echo   [-] Deleting !FINALBASE!.gcode.3mf...
    del /f /q "!INPUTDIR!!FINALBASE!.gcode.3mf"
)

REM 2. Restore the original file
if exist "!INPUTDIR!!NESTBASE!.3mf" (
    echo   [+] Restoring !NESTBASE!.3mf to !FULLBASE!.3mf...
    ren "!INPUTDIR!!NESTBASE!.3mf" "!FULLBASE!.3mf"
) else (
    echo   [!] Could not find !NESTBASE!.3mf to restore.
)
exit /b

:finish
if !PAUSE_AT_END!==1 (
    echo ==========================================
    echo Revert complete! Your Full.3mf files are restored.
    pause
)
exit /b