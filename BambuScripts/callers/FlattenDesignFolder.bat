@echo off
setlocal EnableDelayedExpansion

if "%~1"=="" (
    echo.
    echo  No folders were dropped onto this script.
    echo  Drag one or more Design Folders onto this file to use it.
    echo.
    pause
    exit /b
)

set "TMPLIST=%TEMP%\flatten_%RANDOM%.txt"
type nul > "!TMPLIST!"

set /a TOTAL_FILES=0
set /a TOTAL_CONFLICTS=0
set /a FOLDER_COUNT=0
set /a VALID_FOLDERS=0

echo.
echo ============================================================
echo  Scanning folders...
echo ============================================================

:scan_next
if "%~1"=="" goto scan_done

set "D=%~1"
shift

if not exist "!D!\" (
    echo.
    echo  SKIPPED ^(not a folder^): !D!
    goto scan_next
)

set /a FOLDER_COUNT+=1
echo.
echo  [!FOLDER_COUNT!] !D!

set /a F=0
set /a C=0

rem PASS 1: Scan for files and conflicts
for /d %%S in ("!D!\*") do (
    for /f "delims=" %%X in ('dir /a-d /s /b "%%S" 2^>nul') do (
        set /a F+=1
        if exist "!D!\%%~nxX" (
            echo      CONFLICT FOUND: %%~nxX  ^(already exists at root^)
            set /a C+=1
        )
    )
)

rem Logic: Only add to the to-do list if there are files AND zero conflicts
if !C! gtr 0 (
    echo      SKIPPING ENTIRE FOLDER: !C! conflict^(s^) detected.
    set /a TOTAL_CONFLICTS+=!C!
) else if !F!==0 (
    echo      Nothing to do - no files found in sub-folders.
) else (
    echo      !F! file^(s^) ready to move.
    echo !D!>>"!TMPLIST!"
    set /a TOTAL_FILES+=!F!
    set /a VALID_FOLDERS+=1
)

goto scan_next
:scan_done

:: ============================================================
::  Summary + single confirmation
:: ============================================================
echo.
echo ============================================================
echo  Summary: !VALID_FOLDERS! folder(s) ready  ^|  !TOTAL_FILES! file(s) to move  ^|  !TOTAL_CONFLICTS! conflict(s) found
echo ============================================================
echo.

if !TOTAL_FILES!==0 (
    echo  Nothing to move. Exiting.
    del "!TMPLIST!" 2>nul
    echo.
    pause
    exit /b
)

set /p "YN= Proceed with flattening the !VALID_FOLDERS! clean folder(s)? (Y/N): "
if /i not "!YN!"=="Y" (
    echo.
    echo  Cancelled. No files were moved.
    del "!TMPLIST!" 2>nul
    echo.
    pause
    exit /b
)

echo.
echo ============================================================
echo  Moving files...
echo ============================================================

set /a TOTAL_MOVED=0
set /a IDX=0

rem PASS 2: Move files (Only processes clean folders, so no conflict checks needed here!)
for /f "usebackq delims=" %%L in ("!TMPLIST!") do (
    set "D=%%L"
    set /a IDX+=1
    echo.
    echo  [!IDX!] !D!

    set /a MOVED=0

    for /d %%S in ("!D!\*") do (
        for /f "delims=" %%X in ('dir /a-d /s /b "%%S" 2^>nul') do (
            move "%%X" "!D!\" >nul
            echo      Moved: %%~nxX
            set /a MOVED+=1
            set /a TOTAL_MOVED+=1
        )
    )

    rem Clean up empty subfolders
    for /f "delims=" %%E in ('dir /ad /s /b "!D!" 2^>nul ^| sort /r') do (
        rd "%%E" 2>nul && echo      Removed empty folder: %%~nxE
    )

    echo      Done - !MOVED! moved.
)

del "!TMPLIST!" 2>nul

echo.
echo ============================================================
echo  Finished - !TOTAL_MOVED! file(s) successfully moved across !VALID_FOLDERS! folder(s).
echo ============================================================
echo.
pause