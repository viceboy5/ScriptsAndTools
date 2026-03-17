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
    set "strippedName=!strippedName: Full.gcode.3mf=!"
    :: Fallback: strip bare .gcode.3mf for names with no separator (e.g. BeaverFull)
    set "strippedName=!strippedName:.gcode.3mf=!"

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
exit /b

:ProcessAndMove
set "thisFile=%~1"
set "thisPng=%~2"
set "thisDir=%~3"
set "thisName=%~4"

set "localTemp=%TEMP%\BambuReplace_%RANDOM%"
mkdir "%localTemp%"
set "zipName=%thisName:.gcode.3mf=.zip%"
set "retryCount=0"

:RenameRetry
set /a "retryCount+=1"
if %retryCount% gtr 15 (
    echo [ERROR] Gave up after 15 attempts: %thisName%
    rd /s /q "%localTemp%"
    exit /b 1
)

:: Step 1 - Rename to .zip
echo [DEBUG] Attempt %retryCount% - renaming to zip...
rename "%thisFile%" "%zipName%"
if errorlevel 1 (
    echo [WAIT] Rename failed - file locked, waiting 3s...
    timeout /t 3 /nobreak >nul
    goto RenameRetry
)

:: Step 2 - Probe with tar -tf to confirm content is accessible
echo [DEBUG] Probing archive readability...
tar.exe -tf "%thisDir%%zipName%" >nul 2>&1
if errorlevel 1 (
    echo [WAIT] Cloud stub not ready, triggering download and waiting 10s...
    rename "%thisDir%%zipName%" "%thisName%"
    type "%thisFile%" >nul 2>&1
    timeout /t 10 /nobreak >nul
    goto RenameRetry
)

:: Step 3 - Extract
echo [DEBUG] Extracting...
tar.exe -xf "%thisDir%%zipName%" -C "%localTemp%" 2>&1
if errorlevel 1 (
    echo [ERROR] tar extraction failed: %thisName%
    rename "%thisDir%%zipName%" "%thisName%"
    rd /s /q "%localTemp%"
    exit /b 1
)

:: Step 4 - Find Metadata folder and replace plate_1.png
echo [DEBUG] Locating Metadata folder...
set "metadataDir="
for /f "delims=" %%D in ('dir /b /s /ad "%localTemp%" 2^>nul ^| findstr /i "\\Metadata$"') do (
    set "metadataDir=%%D"
)

if defined metadataDir (
    move /y "%thisPng%" "%metadataDir%\plate_1.png"
    if errorlevel 1 (
        echo [ERROR] Failed to move PNG into archive: %thisPng%
    ) else (
        echo [OK] Replaced plate_1.png in: %metadataDir%
    )
) else (
    echo [WARN] Metadata folder not found. Archive contents:
    dir /b /s "%localTemp%"
)

:: Step 5 - Re-zip and rename back
echo [DEBUG] Re-compressing...
del "%thisDir%%zipName%"
pushd "%localTemp%"
tar.exe -a -cf "%thisDir%%zipName%" * 2>&1
popd
if errorlevel 1 (
    echo [ERROR] tar re-compression failed: %thisName%
    rd /s /q "%localTemp%"
    exit /b 1
)

echo [DEBUG] Renaming back to 3MF...
rename "%thisDir%%zipName%" "%thisName%"
if errorlevel 1 (
    echo [ERROR] Final rename back to 3MF failed: %thisName%
    rd /s /q "%localTemp%"
    exit /b 1
)

echo [INJECTED] %thisName%
rd /s /q "%localTemp%"
exit /b 0