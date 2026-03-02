@echo off
setlocal EnableDelayedExpansion

set "SCRIPT=%~dp0merge_3mf_worker.ps1"
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
    REM It is a folder: save inside it
    set "REPORT_DIR=%~f1\"
) else (
    REM It is a file: save in the same directory as the file
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

REM Clear old temp lists just in case
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
set /p "CHOICE_SLICE=Do you want to automatically slice and export the merged files? (Y/N): "
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

REM --- Build Final Report (Grouped to prevent empty files) ---
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

if exist "%SUCCESS_LIST%" (
    type "%SUCCESS_LIST%" >> "%REPORT%"
) else (
    echo None. >> "%REPORT%"
)

(
    echo.
    echo [ UNSUCCESSFUL FILES ]
) >> "%REPORT%"

if exist "%FAIL_LIST%" (
    type "%FAIL_LIST%" >> "%REPORT%"
) else (
    echo None. >> "%REPORT%"
)

(
    echo.
    echo --------------------------------------------------
    echo Succeeded: !PROCESSED! ^| Failed: !ERRORS! ^| Total: !TOTAL!
) >> "%REPORT%"

REM --- Final Cleanup ---
del "%SUCCESS_LIST%" 2>nul
del "%FAIL_LIST%" 2>nul

echo -------------------------------------------
echo Done. Succeeded: !PROCESSED!   Failed: !ERRORS!
echo Report saved to: !REPORT!
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
set "NESTNAME=!INPUTBASE:~0,-4!Nest.3mf"

for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORKERLOG=%TEMP%\merge_log_%%T.txt"
for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORKERRPT=%TEMP%\merge_rpt_%%T.txt"
for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORK=%TEMP%\merge_work_%%T"

set /a _IDX=PROCESSED+ERRORS+1
echo --------------------------------------------------
echo [!_IDX!/!TOTAL!] Processing: !INPUTNAME!
mkdir "!WORK!" 2>nul

REM Extract
powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUT!', '!WORK!')" >> "!WORKERLOG!" 2>&1
if errorlevel 1 (
    echo   ERROR: Extract failed.
    echo - !INPUTNAME! [Extraction Failed] >> "%FAIL_LIST%"
    set /a ERRORS+=1
    goto cleanup
)

REM Run Worker
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkDir "!WORK!" -InputPath "!INPUT!" -OutputPath "!TEMPOUT!" -ReportPath "!WORKERRPT!" >> "!WORKERLOG!" 2>&1
if errorlevel 1 (
    echo   ERROR: Merge script failed.
    echo - !INPUTNAME! [Script Error] >> "%FAIL_LIST%"
    set /a ERRORS+=1
    del "!TEMPOUT!" 2>nul
    goto cleanup
)

REM Handle Success
ren "!INPUT!" "!NESTNAME!"
ren "!TEMPOUT!" "!INPUTNAME!"

REM --- AUTOMATED SLICING AND EXTRACTION ---
if "!DO_SLICE!"=="0" goto skip_slicing

set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"
set "SLICED_OUT=!INPUTDIR!!INPUTBASE!.gcode.3mf"

echo   Slicing Plate 1 ^(Bypassing safety checks^)...

"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SLICED_OUT!" "!INPUTDIR!!INPUTNAME!" | findstr "^"

if not exist "!SLICED_OUT!" (
    echo   WARNING: Slicing failed. Check the live log output above.
    goto skip_slicing
)

echo   Slicing complete: !INPUTBASE!.gcode.3mf
echo   Extracting design data...

set "data_dir=!WORK!\data_extract"
mkdir "!data_dir!"
pushd "!data_dir!"
tar -xf "!SLICED_OUT!" "Metadata/plate_1.gcode" "Metadata/slice_info.config" >nul 2>&1
popd

set "config_file=!data_dir!\Metadata\slice_info.config"
set "extracted_gcode=!data_dir!\Metadata\plate_1.gcode"

if not exist "!config_file!" (
    echo   WARNING: Could not extract metadata from sliced file.
    goto skip_slicing
)

for /f "usebackq" %%d in (`powershell -NoProfile -Command "(Get-Date).ToString('M/d/yyyy')"`) do set "today_date=%%d"

for /L %%i in (1,1,4) do (
    set "fil%%i_g=0"
    set "fil%%i_color=0"
)

for /f "usebackq tokens=1-9" %%A in (`powershell -NoProfile -Command "$xml=[xml](Get-Content '!config_file!'); $out=@(); 1..4 | %% { $id=$_; $n=$xml.SelectSingleNode('//filament[@id=' + $id + ']'); if($n){ $out += $n.used_g; $out += $n.color } else { $out += '0'; $out += '0' } }; $objs=$xml.SelectNodes('//plate/object'); $c=0; if($objs){ foreach($o in $objs){ if($o.name -eq 'Assembly'){ $c+=2 } else { $c+=1 } } }; $out += $c; $out -join ' '"`) do (
    set "fil1_g=%%A"
    set "fil1_color=%%B"
    set "fil2_g=%%C"
    set "fil2_color=%%D"
    set "fil3_g=%%E"
    set "fil3_color=%%F"
    set "fil4_g=%%G"
    set "fil4_color=%%H"
    set "obj_count=%%I"
)

REM Search CSV for the current hex code and update the variable
for /L %%i in (1,1,4) do (
    set "current_hex=!fil%%i_color!"
    if exist "%~dp0colorNamesCSV.csv" (
        for /f "tokens=2 delims=," %%A in ('findstr /I /C:"!current_hex!" "%~dp0colorNamesCSV.csv" 2^>nul') do (
            set "fil%%i_color=%%A"
        )
    )
)

set "color_change=0"
if exist "!extracted_gcode!" (
    for /f "usebackq" %%C in (`powershell -NoProfile -Command "(Select-String -Path '!extracted_gcode!' -Pattern 'M620 S').Count"`) do set "color_change=%%C"
)

for /f "usebackq tokens=1,2" %%H in (`powershell -NoProfile -Command "$line = Select-String -Path '!extracted_gcode!' -Pattern '; total estimated time:'; if($line){ $str = $line.Line; if($str -match '(\d+)d') { $d=[int]$matches[1] } else {$d=0}; if($str -match '(\d+)h') { $h=[int]$matches[1] } else {$h=0}; if($str -match '(\d+)m') { $m=[int]$matches[1] } else {$m=0}; if($str -match '(\d+)s' -and [int]$matches[1] -ge 30) { $m++ }; if($m -ge 60){ $m=0; $h++ }; $th=($d*24)+$h; \"$th $m\" } else { \"0 0\" }"`) do (
    set "total_h=%%H"
    set "total_m=%%I"
)

REM Build the TSV string and write to files
set "ps_cmd=\"!INPUTBASE!\" + \"`t`t\" + \"!today_date!\" + \"`t\" + \"!total_h!\" + \"`t\" + \"!total_m!\" + \"`t\" + \"!fil1_g!\" + \"`t\" + \"!fil1_color!\" + \"`t\" + \"!fil2_g!\" + \"`t\" + \"!fil2_color!\" + \"`t\" + \"!fil3_g!\" + \"`t\" + \"!fil3_color!\" + \"`t\" + \"!fil4_g!\" + \"`t\" + \"!fil4_color!\" + \"`t`t\" + \"!color_change!\" + \"`t\" + \"!obj_count!\""

set "INDIVIDUAL_DATA=!INPUTDIR!!INPUTBASE!_Data.tsv"

powershell -NoProfile -Command "$str = !ps_cmd!; $str | Out-File -FilePath '!INDIVIDUAL_DATA!' -Encoding UTF8"
powershell -NoProfile -Command "$str = !ps_cmd!; $str | Out-File -FilePath '!MASTER_DATA!' -Append -Encoding UTF8"

echo   Data saved to !INPUTBASE!_Data.tsv and Master_Design_Data.tsv

:skip_slicing

REM Pluck data directly from the worker report using findstr
set "FINAL_COUNT=Unknown"
set "LONE_COUNT=Unknown"

if exist "!WORKERRPT!" (
    for /f "tokens=2 delims=:" %%A in ('findstr /I "Final" "!WORKERRPT!"') do (
        set "VAL=%%A"
        REM strip spaces
        set "FINAL_COUNT=!VAL: =!"
    )
    for /f "tokens=2 delims=:" %%A in ('findstr /I "Lone" "!WORKERRPT!"') do (
        set "VAL=%%A"
        REM Grab just the first word (number) before the parentheticals
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