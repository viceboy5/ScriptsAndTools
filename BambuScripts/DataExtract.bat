@echo off
TITLE 3MF Data Extractor

:: Check if a file or folder was actually dropped
if "%~1"=="" (
    echo [ERROR] No file or folder detected.
    echo Please drag and drop a .gcode.3mf file or a master folder onto this batch file.
    echo.
    pause
    exit /b
)

:: Check if the dropped item is a directory
if exist "%~1\" (
    echo [INFO] Folder detected. Searching recursively for .gcode.3mf files...
    echo.

    :: Recursively search for any file ending in .gcode.3mf
    for /r "%~1" %%F in (*.gcode.3mf) do (
        echo ---------------------------------------------------------
        echo Processing: %%~nxF
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "%%~fF"
    )
) else (
    :: It's a single file
    echo [INFO] Single file detected. Passing to PowerShell...
    echo ---------------------------------------------------------
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "%~1"
)

echo.
echo =====================================================================
echo All tasks complete!
pause