@echo off
setlocal enabledelayedexpansion

echo =====================================================================
echo BAMBU BATCH IMAGE MOVE + LOGGING
echo =====================================================================

:: 1. Check if a master folder was dropped [cite: 1]
if "%~1"=="" (
    echo [ERROR] Please drag and drop the MASTER FOLDER onto this script.
    pause
    exit /b
)

set "masterDir=%~1"
set "logFile=%masterDir%\ProcessLog.txt"

echo Processing started: %date% %time% > "%logFile%"
echo --------------------------------------------------------- >> "%logFile%"

:: 2. Recursively search for *_Full.gcode.3mf files
for /r "%masterDir%" %%F in (*_Full.gcode.3mf) do (
    set "targetFile=%%~fF"
    set "workingDir=%%~dpF"
    set "originalFullName=%%~nxF"
    
    :: Look for any PNG in the same subfolder [cite: 3]
    set "newPng="
    for %%P in ("!workingDir!\*.png") do (
        set "newPng=%%~fP"
    )

    echo.
    echo ---------------------------------------------------------
    echo Folder: !workingDir!

    if not defined newPng (
        echo [SKIP] No PNG found for !originalFullName!
        echo [SKIP] No PNG found in: !workingDir! >> "%logFile%"
    ) else (
        echo [MOVING] !newPng!
        call :ProcessAndMove "!targetFile!" "!newPng!" "!workingDir!" "!originalFullName!"
        echo [SUCCESS] Processed !originalFullName! >> "%logFile%"
    )
)

echo.
echo --------------------------------------------------------- >> "%logFile%"
echo Processing finished: %date% %time% >> "%logFile%"
echo =====================================================================
echo All tasks complete! Log saved to: %logFile%
pause
exit /b

:: ---------------------------------------------------------
:: Processing Subroutine
:: ---------------------------------------------------------
:ProcessAndMove
set "thisFile=%~1"
set "thisPng=%~2"
set "thisDir=%~3"
set "thisName=%~4"

:: Create unique local temp folder to avoid sync locks [cite: 4]
set "localTemp=%TEMP%\BambuReplace_%RANDOM%"
mkdir "%localTemp%"

:: Rename to .zip to allow tar extraction [cite: 4]
set "zipName=%thisName:.gcode.3mf=.zip%"
rename "%thisFile%" "%zipName%"

:: Extract to local %TEMP% [cite: 4]
tar.exe -xf "%thisDir%\%zipName%" -C "%localTemp%"

:: Move image INTO the metadata folder (Only replacing plate_1.png) 
set "metadataDir=%localTemp%\Metadata"
if exist "%metadataDir%" (
    move /y "%thisPng%" "%metadataDir%\plate_1.png" >nul
)

:: Re-zip and move back to original location
del "%thisDir%\%zipName%"
pushd "%localTemp%"
tar.exe -a -cf "%thisDir%\%zipName%" *
popd

:: Final Rename back to 3MF
rename "%thisDir%\%zipName%" "%thisName%"

:: Cleanup local buffer [cite: 6]
rd /s /q "%localTemp%"
exit /b