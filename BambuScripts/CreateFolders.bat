@echo off
setlocal enabledelayedexpansion

:: Check if any files were dropped
if "%~1"=="" (
    echo Please drag and drop files onto this script.
    pause
    exit /b
)

echo Organizing files...
echo.

:: Loop through every file dropped onto the script
for %%I in (%*) do (
    :: Check if the dropped item is a file (skips folders)
    if not exist "%%~I\" (

        :: %%~dpI is the drive and path (e.g., C:\Your\Folder\)
        :: %%~nI is the file name without the extension
        set "targetFolder=%%~dpI%%~nI"

        :: Create the new folder if it doesn't already exist
        if not exist "!targetFolder!" (
            mkdir "!targetFolder!"
        )

        :: Move the file into its new folder
        move "%%~I" "!targetFolder!\" >nul
        echo Moved: %%~nxI -^> !targetFolder!
    ) else (
        echo Skipped folder: %%~nxI
    )
)

echo.
echo All done!
pause