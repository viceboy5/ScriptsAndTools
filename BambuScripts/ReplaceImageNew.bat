@echo off
setlocal enabledelayedexpansion

echo =====================================================================
echo BAMBU BATCH IMAGE MOVE + LOGGING
echo =====================================================================

if "%~1"=="" (
    echo [ERROR] Please drag and drop the MASTER FOLDER onto this script.
    pause
    exit /b
)

:: Safely capture the dropped folder path with standard expansion
set "masterDir=%~1"
if "%masterDir:~-1%"=="\" set "masterDir=%masterDir:~0,-1%"
set "logFile=%masterDir%\ProcessLog.txt"

echo Processing started: %date% %time% > "%logFile%"
echo --------------------------------------------------------- >> "%logFile%"

:: MUST use standard percentage signs for the FOR /R loop!
for /r "%masterDir%" %%F in (*Full.gcode.3mf) do (
    set "targetFile=%%~fF"
    set "workingDir=%%~dpF"
    set "originalFullName=%%~nxF"

    :: Strip all variations of Full.gcode.3mf
    set "strippedName=%%~nxF"
    set "strippedName=!strippedName:_Full.gcode.3mf=!"
    set "strippedName=!strippedName:.Full.gcode.3mf=!"
    :: NEW: Strip the version with a space
    set "strippedName=!strippedName: Full.gcode.3mf=!"

    set "newPng="
    if exist "!workingDir!!strippedName!_slicePreview.png" (
        set "newPng=!workingDir!!strippedName!_slicePreview.png"
    )

    echo.
    echo ---------------------------------------------------------
    echo Folder: !workingDir!
    echo Base name: !strippedName!

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

:ProcessAndMove
set "thisFile=%~1"
set "thisPng=%~2"
set "thisDir=%~3"
set "thisName=%~4"

set "localTemp=%TEMP%\BambuReplace_%RANDOM%"
mkdir "%localTemp%"

set "zipName=%thisName:.gcode.3mf=.zip%"
rename "%thisFile%" "%zipName%"
if errorlevel 1 (
    echo [ERROR] Could not rename %thisFile%
    rd /s /q "%localTemp%"
    exit /b
)

tar.exe -xf "%thisDir%%zipName%" -C "%localTemp%"
if errorlevel 1 (
    echo [ERROR] tar extraction failed
    rename "%thisDir%%zipName%" "%thisName%"
    rd /s /q "%localTemp%"
    exit /b
)

set "metadataDir="
for /f "delims=" %%D in ('dir /b /s /ad "%localTemp%" 2^>nul ^| findstr /i "\\Metadata$"') do (
    set "metadataDir=%%D"
)

if defined metadataDir (
    move /y "%thisPng%" "%metadataDir%\plate_1.png" >nul
)

del "%thisDir%%zipName%"
pushd "%localTemp%"
tar.exe -a -cf "%thisDir%%zipName%" *
popd

rename "%thisDir%%zipName%" "%thisName%"

rd /s /q "%localTemp%"
exit /b