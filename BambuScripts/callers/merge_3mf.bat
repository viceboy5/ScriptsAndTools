@echo off
setlocal EnableDelayedExpansion

set "SCRIPT=%~dp0..\workers\merge_3mf_worker.ps1"
set "PREP_ERRORS=0"
set "PREP_PROCESSED=0"
set "PREP_SKIPPED=0"
set "SLICE_ERRORS=0"
set "SLICE_PROCESSED=0"
set "TOTAL=0"

if "%~1"=="" exit /b 1

:: Derive Target Name
for %%I in ("%~1") do set "TARGET_NAME=%%~nxI"
set "TARGET_NAME=!TARGET_NAME:.3mf=!"
set "TARGET_NAME=!TARGET_NAME:.3MF=!"

if exist "%~1\" ( set "REPORT_DIR=%~f1\" ) else ( set "REPORT_DIR=%~dp1" )
set "MASTER_DATA=!REPORT_DIR!!TARGET_NAME!_Design_Data.tsv"
set "FILELIST=%TEMP%\merge_3mf_list_%RANDOM%.txt"
set "SLICELIST=%TEMP%\slice_3mf_list_%RANDOM%.txt"

:: --- PRE-FLIGHT CLEANUP WORKER ---
:: We strip the trailing backslash so it doesn't accidentally escape the quote mark!
set "CLEAN_DIR=!REPORT_DIR:~0,-1!"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\cleanup_old_worker.ps1" -TargetDir "!CLEAN_DIR!"

:: Silent Cleanup
for /d /r "%~dp1" %%d in (temp_3mf_extract) do ( if exist "%%d" rmdir /s /q "%%d" 2>nul )

if exist "%FILELIST%" del "%FILELIST%"
if exist "%SLICELIST%" del "%SLICELIST%"

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
if !TOTAL!==0 exit /b 1

set /p "CHOICE_SLICE=Automated slice files? (Y/N): "
if /I "!CHOICE_SLICE!"=="Y" ( set "DO_SLICE=1" ) else ( set "DO_SLICE=0" )

set /p "CHOICE_COLORS=Scan and pick new colors? (Y/N): "
if /I "!CHOICE_COLORS!"=="Y" ( set "DO_COLORS=1" ) else ( set "DO_COLORS=0" )

:: --- IMAGE LOGIC ---
set /p "CHOICE_IMAGE=Generate composite image cards? (Y/N): "
set "GEN_IMAGE_SWITCH="
if /I "!CHOICE_IMAGE!"=="Y" ( set "GEN_IMAGE_SWITCH=-GenerateImage" )

:: --- EXTRACTION LOGIC ---
set /p "CHOICE_DATA=Extract data / update TSV? (Y/N): "
set "DO_EXTRACT=0"
if /I "!CHOICE_DATA!"=="Y" ( set "DO_EXTRACT=1" )
if "!DO_SLICE!"=="1" ( set "DO_EXTRACT=1" )

:: Master switch: Do we need to call PowerShell in Phase 2 at all?
set "CALL_PS1=0"
if "!DO_EXTRACT!"=="1" set "CALL_PS1=1"
if not "!GEN_IMAGE_SWITCH!"=="" set "CALL_PS1=1"
echo.
echo ==============================================================
echo PHASE 1: PREPARATION ^& COLOR MAPPING (Requires User Input)
echo ==============================================================
for /f "usebackq delims=" %%F in ("%FILELIST%") do call :prepare_file "%%F"

if "!DO_SLICE!"=="0" (
    if "!CALL_PS1!"=="0" goto skip_slicing_phase
)

echo.
echo ==============================================================
echo PHASE 2: SLICING ^& DATA EXTRACTION (Unattended)
echo ==============================================================
if exist "%SLICELIST%" (
    for /f "usebackq delims=" %%F in ("%SLICELIST%") do call :slice_file "%%F"
) else (
    echo No files queued for slicing.
)

:skip_slicing_phase
del "%FILELIST%" 2>nul
del "%SLICELIST%" 2>nul

echo -------------------------------------------
echo Done.
echo Prepped: !PREP_PROCESSED! / Skipped: !PREP_SKIPPED! / Failed: !PREP_ERRORS!
if "!DO_SLICE!"=="1" echo Sliced:  !SLICE_PROCESSED! / Failed: !SLICE_ERRORS!
echo -------------------------------------------
pause
exit /b


:prepare_file
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTNAME=%~nx1"
set "INPUTBASE=%~n1"
set "TEMPOUT=%~dp1%~n1_merged_temp.3mf"

:: The ~0,-4 math gracefully handles spaces, periods, and underscores
set "NESTBASE=!INPUTBASE:~0,-4!Nest"
set "FINALBASE=!INPUTBASE:~0,-4!Final"
set "NESTNAME=!NESTBASE!.3mf"
set "FINAL_PATH=!INPUTDIR!!FINALBASE!.3mf"
set "NEST_PATH=!INPUTDIR!!NESTNAME!"

set /a _IDX=PREP_PROCESSED+PREP_ERRORS+PREP_SKIPPED+1
echo.
echo [!_IDX!/!TOTAL!] Preparing: !INPUTNAME!

:: --- NEW PHASE 1 IMAGE CHECK ---
if not "!GEN_IMAGE_SWITCH!"=="" (
    set "FOUND_PNG="
    for %%P in ("!INPUTDIR!*.png") do (
        if /I not "%%~nxP"=="!INPUTBASE!.png" set "FOUND_PNG=%%P"
    )

    if "!FOUND_PNG!"=="" (
        echo   [!] No custom image found for !INPUTNAME!
        set /p "DROPPED_IMG=      Drag and drop an image here and press Enter (or press Enter to skip): "

        if not "!DROPPED_IMG!"=="" (
            :: Strip out the quotes Windows adds during drag-and-drop
            set "DROPPED_IMG=!DROPPED_IMG:"=!"
            if exist "!DROPPED_IMG!" (
                copy /Y "!DROPPED_IMG!" "!INPUTDIR!" >nul
                echo   [+] Image accepted and copied.
            ) else (
                echo   [-] Invalid path. Will use internal thumbnail later.
            )
        ) else (
            echo   [*] Skipped. Will use internal thumbnail later.
        )
    ) else (
        echo   [+] Found custom image: %%~nxFOUND_PNG%%
    )
)
:: -------------------------------

:: --- PRE-FLIGHT REVERT CHECK ---
if exist "!NEST_PATH!" (
    echo   [!] PREVIOUS MERGE DETECTED.
    choice /C YN /M "      Do you want to REVERT this file and re-process it"
    if errorlevel 2 (
        echo   [-] Skipping Prep for !INPUTNAME!.
        set /a PREP_SKIPPED+=1

        :: Queue for Phase 2 if we need images or data
        if "!CALL_PS1!"=="1" (
            echo !INPUT!>> "%SLICELIST%"
        )
        goto :eof
    )
    echo   [+] Calling Revert Worker...
    set "WORKER_MODE=1"
    call "%~dp0RevertMerge.bat" "!NEST_PATH!"
    set "WORKER_MODE=0"
)

for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORK=%TEMP%\merge_work_%%T"
mkdir "!WORK!" 2>nul

powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUT!', '!WORK!')" >nul 2>&1
if errorlevel 1 ( echo   ERROR: Extract failed. & set /a PREP_ERRORS+=1 & goto cleanup_prep )

if "!DO_COLORS!"=="1" (
    echo   Checking/Updating Colors...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\ColorUpdateOnly_worker.ps1" -WorkDir "!WORK!" -FileName "!INPUTNAME!" -OriginalZip "!INPUT!"
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkDir "!WORK!" -InputPath "!INPUT!" -OutputPath "!TEMPOUT!" -ReportPath "nul"
if errorlevel 1 ( echo   ERROR: Merge script failed. & set /a PREP_ERRORS+=1 & del "!TEMPOUT!" 2>nul & goto cleanup_prep )

ren "!INPUT!" "!NESTNAME!"
ren "!TEMPOUT!" "!INPUTNAME!"

if exist "!FINAL_PATH!" del /f /q "!FINAL_PATH!"
set "WORK_SINGLE=%TEMP%\single_work_%RANDOM%"
mkdir "!WORK_SINGLE!" 2>nul

powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUTDIR!!NESTNAME!', '!WORK_SINGLE!')" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\isolate_final_worker.ps1" -WorkDir "!WORK_SINGLE!" -OutputPath "!FINAL_PATH!"
rmdir /s /q "!WORK_SINGLE!" 2>nul

if not exist "!FINAL_PATH!" echo   [!] WARNING: Final.3mf failed to generate.

:: Add successfully prepared files to the Slicing Queue
set /a PREP_PROCESSED+=1
echo !INPUT!>> "%SLICELIST%"

:cleanup_prep
rmdir /s /q "!WORK!" 2>nul
goto :eof


:slice_file
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTNAME=%~nx1"
set "INPUTBASE=%~n1"
set "FINALBASE=!INPUTBASE:~0,-4!Final"
set "FINAL_PATH=!INPUTDIR!!FINALBASE!.3mf"

set "SLICED_OUT=!INPUTDIR!!INPUTBASE!.gcode.3mf"
set "SLICED_FINAL_TEMP=!INPUTDIR!!FINALBASE!.gcode.3mf"

set /a _SIDX=SLICE_PROCESSED+SLICE_ERRORS+1
echo.
echo [!_SIDX!/!PREP_PROCESSED!] Processing Phase 2: !INPUTNAME!

:: 1. CONDITIONAL SLICING
if "!DO_SLICE!"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\Slice_worker.ps1" -InputPath "!INPUT!" -IsolatedPath "!FINAL_PATH!"
    if errorlevel 1 (
        set /a SLICE_ERRORS+=1
        goto :eof
    )
) else (
    if not exist "!SLICED_OUT!" (
        echo   [-] No sliced .gcode.3mf found to extract from. Skipping.
        set /a SLICE_ERRORS+=1
        goto :eof
    ) else (
        echo   [*] Slicing bypassed. Extracting from existing file...
    )
)

:: 2. EXTRACTION & IMAGE GENERATION
set "EXTRACT_FLAGS="
if "!DO_EXTRACT!"=="0" set "EXTRACT_FLAGS=-SkipExtraction"

if exist "!SLICED_FINAL_TEMP!" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\DataExtract_worker.ps1" -InputFile "!SLICED_OUT!" -SingleFile "!SLICED_FINAL_TEMP!" -MasterTsvPath "!MASTER_DATA!" -IndividualTsvPath "!INPUTDIR!!INPUTBASE!_Data.tsv" !GEN_IMAGE_SWITCH! !EXTRACT_FLAGS!
    del "!SLICED_FINAL_TEMP!" /q
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\DataExtract_worker.ps1" -InputFile "!SLICED_OUT!" -MasterTsvPath "!MASTER_DATA!" -IndividualTsvPath "!INPUTDIR!!INPUTBASE!_Data.tsv" !GEN_IMAGE_SWITCH! !EXTRACT_FLAGS!
)

echo   OK --^> Success.
set /a SLICE_PROCESSED+=1
goto :eof