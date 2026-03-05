@echo off
setlocal EnableDelayedExpansion

set "SCRIPT=%~dp0merge_3mf_worker.ps1"
set "SCRIPT_SINGLE=%~dp0isolate_single_worker.ps1"
set "ERRORS=0"
set "PROCESSED=0"
set "TOTAL=0"

if "%~1"=="" (
    echo Usage: drag .3mf files and/or folders onto this batch file.
    pause
    exit /b 1
)

REM --- Derive Report Name and Location ---
for %%I in ("%~1") do set "TARGET_NAME=%%~nxI"
set "TARGET_NAME=!TARGET_NAME:.3mf=!"
set "TARGET_NAME=!TARGET_NAME:.3MF=!"

REM Check if the dropped item is a folder or a file
if exist "%~1\" (
    set "REPORT_DIR=%~f1\"
) else (
    set "REPORT_DIR=%~dp1"
)
set "REPORT=!REPORT_DIR!!TARGET_NAME!_MergeReport.txt"
set "MASTER_DATA=!REPORT_DIR!!TARGET_NAME!_Design_Data.tsv"

set "FILELIST=%TEMP%\merge_3mf_list_%RANDOM%.txt"
set "SUCCESS_LIST=%TEMP%\merge_success_%RANDOM%.txt"
set "FAIL_LIST=%TEMP%\merge_fail_%RANDOM%.txt"

REM --- Pre-Cleanup: Scorch any existing temp_3mf_extract folders ---
echo --------------------------------------------------
echo Cleaning up phantom folders...
for /d /r "%~dp1" %%d in (temp_3mf_extract) do (
    if exist "%%d" (
        rmdir /s /q "%%d" 2>nul
    )
)
echo Clean up complete.
echo --------------------------------------------------

REM Clear old temp lists
if exist "%FILELIST%" del "%FILELIST%"
if exist "%SUCCESS_LIST%" del "%SUCCESS_LIST%"
if exist "%FAIL_LIST%" del "%FAIL_LIST%"

REM --- Collect files ---
:collect_loop
if "%~1"=="" goto begin_process
if exist "%~1\" (
    for /r "%~1" %%F in (*.3mf) do (
        set "_n=%%~nF"
        set "_p=%%~dpF"
        if /I "!_p:old=!" == "!_p!" (
            if /I "!_n:~-4!"=="full" (
                echo %%F>> "%FILELIST%"
                set /a TOTAL+=1
            )
        )
    )
) else if exist "%~1" (
    echo %~1>> "%FILELIST%"
    set /a TOTAL+=1
)
shift
goto collect_loop

:begin_process
if !TOTAL! == 0 (
    echo No matching files found.
    if exist "%FILELIST%" del "%FILELIST%"
    pause
    exit /b 1
)

echo Found !TOTAL! file(s) to process.
echo.

REM --- Prompt for Slicing ---
set /p "CHOICE_SLICE=Do you want to automatically slice and extract data? (Y/N): "
if /I "!CHOICE_SLICE!"=="Y" (
    set "DO_SLICE=1"
    echo Automated slicing and data extraction is ENABLED.
) else (
    set "DO_SLICE=0"
    echo Automated slicing is DISABLED.
)
echo.

REM --- Process each file ---
for /f "usebackq delims=" %%F in ("%FILELIST%") do call :process_one "%%F"
del "%FILELIST%" 2>nul

REM --- Build Final Report ---
goto skip_report

echo Generating Final Report...
(
    echo MERGE 3MF SESSION REPORT
    echo Date/Time: %DATE% %TIME%
    echo Target: !TARGET_NAME!
    echo Total files queued: !TOTAL!
    echo --------------------------------------------------
    echo.
    echo [ SUCCESSFUL FILES ]
) > "%REPORT%"

if exist "%SUCCESS_LIST%" ( type "%SUCCESS_LIST%" >> "%REPORT%" ) else ( echo None. >> "%REPORT%" )

(
    echo.
    echo [ UNSUCCESSFUL FILES ]
) >> "%REPORT%"

if exist "%FAIL_LIST%" ( type "%FAIL_LIST%" >> "%REPORT%" ) else ( echo None. >> "%REPORT%" )

(
    echo.
    echo --------------------------------------------------
    echo Succeeded: !PROCESSED! ^| Failed: !ERRORS! ^| Total: !TOTAL!
) >> "%REPORT%"
:skip_report

del "%SUCCESS_LIST%" 2>nul
del "%FAIL_LIST%" 2>nul

echo -------------------------------------------
echo Done. Succeeded: !PROCESSED!   Failed: !ERRORS!
if exist "!MASTER_DATA!" echo Master Data saved: !MASTER_DATA!
echo -------------------------------------------
pause
endlocal
goto :eof

REM --- Subroutine ---
:process_one
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTNAME=%~nx1"
set "INPUTBASE=%~n1"
set "TEMPOUT=%~dp1%~n1_merged_temp.3mf"

REM Clean Base Name generation (replaces "Full" with "Nest" and "Final")
set "NESTBASE=!INPUTBASE:~0,-4!Nest"
set "FINALBASE=!INPUTBASE:~0,-4!Final"

set "NESTNAME=!NESTBASE!.3mf"
set "FINAL_PATH=!INPUTDIR!!FINALBASE!.3mf"

for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORKERLOG=%TEMP%\merge_log_%%T.txt"
for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORKERRPT=%TEMP%\merge_rpt_%%T.txt"
for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORK=%TEMP%\merge_work_%%T"

set /a _IDX=PROCESSED+ERRORS+1
echo --------------------------------------------------
echo [!_IDX!/!TOTAL!] Processing: !INPUTNAME!
mkdir "!WORK!" 2>nul

REM 1. Merge Extract
powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUT!', '!WORK!')" >> "!WORKERLOG!" 2>&1
if errorlevel 1 (
    echo   ERROR: Extract failed.
    echo - !INPUTNAME! [Extraction Failed] >> "%FAIL_LIST%"
    set /a ERRORS+=1
    goto cleanup
)

REM 2. Run Worker
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkDir "!WORK!" -InputPath "!INPUT!" -OutputPath "!TEMPOUT!" -ReportPath "!WORKERRPT!" >> "!WORKERLOG!" 2>&1
if errorlevel 1 (
    echo   ERROR: Merge script failed.
    echo - !INPUTNAME! [Script Error] >> "%FAIL_LIST%"
    set /a ERRORS+=1
    del "!TEMPOUT!" 2>nul
    goto cleanup
)

REM 3. Handle Merge Success
ren "!INPUT!" "!NESTNAME!"
ren "!TEMPOUT!" "!INPUTNAME!"

REM 4. --- CREATE SINGLE OBJECT (FINAL) FILE ---
echo   Generating !FINALBASE!.3mf ^(Overwriting if exists^)...

if exist "!FINAL_PATH!" del /f /q "!FINAL_PATH!"
set "WORK_SINGLE=%TEMP%\single_work_%RANDOM%"
mkdir "!WORK_SINGLE!" 2>nul

powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUTDIR!!NESTNAME!', '!WORK_SINGLE!')" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0isolate_final_worker.ps1" -WorkDir "!WORK_SINGLE!" -OutputPath "!FINAL_PATH!"
rmdir /s /q "!WORK_SINGLE!" 2>nul

if not exist "!FINAL_PATH!" (
    echo   [!] WARNING: The Final.3mf file failed to generate.
)

REM --- AUTOMATED SLICING AND EXTRACTION ---
if "!DO_SLICE!"=="0" goto skip_slicing

set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"
set "SLICED_OUT=!INPUTDIR!!INPUTBASE!.gcode.3mf"
set "SLICED_FINAL_TEMP=!INPUTDIR!!FINALBASE!.gcode.3mf"

echo   Slicing Plate 1 ^(Merged Plate^)... Please wait...
"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SLICED_OUT!" "!INPUTDIR!!INPUTNAME!" >nul 2>&1

if not exist "!SLICED_OUT!" (
    echo   WARNING: Slicing failed on merged plate.
    goto skip_slicing
)

echo   Slicing Plate 1 ^(Isolated Single Object^)... Please wait...
"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SLICED_FINAL_TEMP!" "!FINAL_PATH!" >nul 2>&1

echo   Extracting design data and calculating Added Time...
if exist "!SLICED_FINAL_TEMP!" (
    set "INDIVIDUAL_DATA=!INPUTDIR!!INPUTBASE!_Data.tsv"
    
    REM Pass BOTH files and BOTH output paths to the extractor
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!SLICED_OUT!" -SingleFile "!SLICED_FINAL_TEMP!" -MasterTsvPath "!MASTER_DATA!" -IndividualTsvPath "!INDIVIDUAL_DATA!"
    
    echo   Deleting temporary Final.gcode.3mf...
    del "!SLICED_FINAL_TEMP!" /q
) else (
    echo   WARNING: Isolated object failed to slice. Running normal extraction...
    set "INDIVIDUAL_DATA=!INPUTDIR!!INPUTBASE!_Data.tsv"
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!SLICED_OUT!" -MasterTsvPath "!MASTER_DATA!" -IndividualTsvPath "!INDIVIDUAL_DATA!"
)

:skip_slicing
REM Pluck data directly from the worker report using findstr
set "FINAL_COUNT=Unknown"
set "LONE_COUNT=Unknown"

if exist "!WORKERRPT!" (
    for /f "tokens=2 delims=:" %%A in ('findstr /I "Final" "!WORKERRPT!"') do (
        set "VAL=%%A"
        set "FINAL_COUNT=!VAL: =!"
    )
    for /f "tokens=2 delims=:" %%A in ('findstr /I "Lone" "!WORKERRPT!"') do (
        set "VAL=%%A"
        for /f "tokens=1" %%B in ("!VAL!") do set "LONE_COUNT=%%B"
    )
)

(
    echo - !INPUTNAME!
    echo     Final Object Count: !FINAL_COUNT!
    echo     Lone Object Count: !LONE_COUNT!
    echo.
) >> "%SUCCESS_LIST%"

echo   OK --^> Merge Success.
set /a PROCESSED+=1

:cleanup
rmdir /s /q "!WORK!" 2>nul
del "!WORKERLOG!" 2>nul
del "!WORKERRPT!" 2>nul
goto :eof