@echo off
setlocal enabledelayedexpansion

echo =====================================================================
echo BAMBU 3MF OVERWRITE (v2026.01 Release)
echo =====================================================================

:: 1. Handle Drag and Drop
if "%~1"=="" (
    echo [ERROR] Please drag and drop a .3mf file onto this script icon.
    pause & exit /b
)

set "full_path=%~1"
set "workingDir=%~dp1"
set "originalFullName=%~nx1"
set "baseName=%~n1"
set "zipName=%baseName%.zip"

:: 2. Create Temp Workspace
set "tempExtract=%temp%\bambu_extract_%random%"
if exist "%tempExtract%" rd /s /q "%tempExtract%"
mkdir "%tempExtract%"

echo Processing: %originalFullName%

:: 3. Extract using tar.exe
:: Temporarily rename the original .3mf to .zip so tar recognizes the format
rename "%full_path%" "%zipName%"
echo Extracting...
tar.exe -xf "%workingDir%%zipName%" -C "%tempExtract%"

if errorlevel 1 (
    echo [ERROR] Extraction failed. Restoring original...
    rename "%workingDir%%zipName%" "%originalFullName%"
    pause & exit /b
)

:: --- [LOGIC START] ---

:: A. Count Objects
set "configFile=%tempExtract%\Metadata\model_settings.config"
set "objCount=0"
if exist "%configFile%" (
    for /f %%A in ('find /c "<object" ^< "%configFile%"') do set "objCount=%%A"
)
echo Found !objCount! objects.

:: B. Process layer_heights_profile.txt
set "profileFile=%tempExtract%\Metadata\layer_heights_profile.txt"
set "tempProfile=%tempExtract%\Metadata\profile_new.txt"

if exist "%profileFile%" (
    echo Updating profile for !objCount! objects...

    if exist "%tempProfile%" del "%tempProfile%"

    set "dataString="

    :: Read ONLY the first line and split at first |
    for /f "usebackq tokens=1* delims=|" %%A in ("%profileFile%") do (
        if not defined dataString (
            set "dataString=%%B"
        )
    )

    if not defined dataString (
        echo [ERROR] Could not extract profile data.
        pause & exit /b
    )

    for /l %%N in (1,1,!objCount!) do (
        >>"%tempProfile%" echo object_id=%%N^|!dataString!
    )

    move /y "%tempProfile%" "%profileFile%" >nul
)

:: --- [LOGIC END] ---

:: 4. Re-zip and Replace Original
echo Finalizing...

:: Create the new archive as a temporary file first
set "tempZip=%workingDir%final_temp.zip"
pushd "%tempExtract%"
tar.exe -a -cf "%tempZip%" *
popd

:: Replace the original by moving the temp zip over the renamed original
:: This effectively deletes the old .zip and saves the new one as the original .3mf
move /y "%tempZip%" "%workingDir%%originalFullName%" >nul

:: Cleanup the original renamed zip if it still exists (it shouldn't after move /y)
if exist "%workingDir%%zipName%" del "%workingDir%%zipName%"

:: 5. Cleanup
rd /s /q "%tempExtract%"

echo.
echo All tasks complete! 
echo Overwritten: %originalFullName%
pause
