@echo off
setlocal EnableDelayedExpansion

:: ============================================================
::  FlattenDesignFolder.bat
::  Drag one or more Design Folders onto this script.
::  All files inside sub-folders are moved up to the
::  Design Folder root.  Empty sub-folders are removed after.
:: ============================================================

if "%~1"=="" (
    echo.
    echo  No folders were dropped onto this script.
    echo  Drag one or more Design Folders onto this file to use it.
    echo.
    pause
    exit /b
)

set "TMPLIST=%TEMP%\flatten_%RANDOM%.txt"
del "%TMPLIST%" 2>nul

set /a TOTAL_FILES=0
set /a TOTAL_CONFLICTS=0
set /a FOLDER_COUNT=0

echo.
echo ============================================================
echo  Scanning folders...
echo ============================================================

:: ============================================================
::  PASS 1 - Scan, count files, list conflicts
:: ============================================================
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
echo !D!>> "!TMPLIST!"

echo.
echo  [!FOLDER_COUNT!] !D!

set /a F=0
set /a C=0

for /d %%S in ("!D!\*") do (
    for /r "%%S" %%X in (*) do (
        set /a F+=1
        set /a TOTAL_FILES+=1
        if exist "!D!\%%~nxX" (
            echo      CONFLICT: %%~nxX  ^(already exists at root^)
            set /a C+=1
            set /a TOTAL_CONFLICTS+=1
        )
    )
)

if !F!==0 (
    echo      Nothing to do - no files found in sub-folders.
) else (
    echo      !F! file(s) to move, !C! conflict(s)
)

goto scan_next
:scan_done

:: ============================================================
::  Summary + single confirmation
:: ============================================================
echo.
echo ============================================================
echo  Summary: !FOLDER_COUNT! folder(s)  ^|  !TOTAL_FILES! file(s) to move  ^|  !TOTAL_CONFLICTS! conflict(s) ^(will be skipped^)
echo ============================================================
echo.

if !TOTAL_FILES!==0 (
    echo  Nothing to move. Exiting.
    del "!TMPLIST!" 2>nul
    echo.
    pause
    exit /b
)

set /p "YN= Proceed with all folders? (Y/N): "
if /i not "!YN!"=="Y" (
    echo.
    echo  Cancelled. No files were moved.
    del "!TMPLIST!" 2>nul
    echo.
    pause
    exit /b
)

:: ============================================================
::  PASS 2 - Move files and remove empty sub-folders
:: ============================================================
echo.
echo ============================================================
echo  Moving files...
echo ============================================================

set /a TOTAL_MOVED=0
set /a TOTAL_SKIPPED=0
set /a IDX=0

for /f "usebackq delims=" %%L in ("!TMPLIST!") do (
    set "D=%%L"
    set /a IDX+=1
    echo.
    echo  [!IDX!] !D!

    set /a MOVED=0
    set /a SKIPPED=0

    for /d %%S in ("!D!\*") do (
        for /r "%%S" %%X in (*) do (
            if exist "!D!\%%~nxX" (
                echo      SKIPPED: %%~nxX
                set /a SKIPPED+=1
                set /a TOTAL_SKIPPED+=1
            ) else (
                move "%%X" "!D!\" >nul
                echo      Moved: %%~nxX
                set /a MOVED+=1
                set /a TOTAL_MOVED+=1
            )
        )
    )

    for /d /r "!D!" %%E in (*) do rd "%%E" 2>nul && echo      Removed folder: %%~nxE

    echo      Done - !MOVED! moved, !SKIPPED! skipped.
)

del "!TMPLIST!" 2>nul

echo.
echo ============================================================
echo  Finished - !TOTAL_MOVED! file(s) moved, !TOTAL_SKIPPED! skipped across !FOLDER_COUNT! folder(s).
echo ============================================================
echo.
pause
