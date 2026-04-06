@echo off
setlocal enabledelayedexpansion

:: Check if something was dragged onto the script
if "%~1"=="" (
    echo Please drag and drop files or folders onto this script.
    pause
    exit /b
)

echo Scanning for "Nest" (case-insensitive) and renaming to "Full"...
echo -----------------------------------------------------------

:Loop
if "%~1"=="" goto End

:: Check if the input is a directory or a file
if exist "%~1\" (
    :: It's a folder: Search recursively inside it
    pushd "%~1"
    for /r %%F in (*Nest*) do (
        call :RenameProcess "%%F"
    )
    popd
) else (
    :: It's a single file: Check it directly
    echo "%~nx1" | findstr /i "Nest" >nul
    if !errorlevel! == 0 call :RenameProcess "%~1"
)

shift
goto Loop

:RenameProcess
set "fullpath=%~1"
:: Use PowerShell to perform the case-insensitive replacement on the filename only
powershell -NoProfile -Command ^
    "$path = '%fullpath%'; $dir = Split-Path $path; $name = Split-Path $path -Leaf; " ^
    "$newName = $name -ireplace 'Nest', 'Full'; " ^
    "if ($name -ne $newName) { Rename-Item -Path $path -NewName $newName -ErrorAction SilentlyContinue }"
goto :eof

:End
echo -----------------------------------------------------------
echo Done!
pause