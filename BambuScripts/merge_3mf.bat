@echo off
setlocal EnableDelayedExpansion

set "SCRIPT=%~dp0merge_3mf_worker.ps1"
set "ERRORS=0"
set "PROCESSED=0"
set "TOTAL=0"

if "%~1"=="" exit /b 1

:: Derive Target Name
for %%I in ("%~1") do set "TARGET_NAME=%%~nxI"
set "TARGET_NAME=!TARGET_NAME:.3mf=!"
set "TARGET_NAME=!TARGET_NAME:.3MF=!"

if exist "%~1\" ( set "REPORT_DIR=%~f1\" ) else ( set "REPORT_DIR=%~dp1" )
set "MASTER_DATA=!REPORT_DIR!!TARGET_NAME!_Design_Data.tsv"
set "FILELIST=%TEMP%\merge_3mf_list_%RANDOM%.txt"

:: Silent Cleanup
for /d /r "%~dp1" %%d in (temp_3mf_extract) do ( if exist "%%d" rmdir /s /q "%%d" 2>nul )

if exist "%FILELIST%" del "%FILELIST%"
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

set /p "CHOICE_SLICE=Automated slice and extract data? (Y/N): "
if /I "!CHOICE_SLICE!"=="Y" ( set "DO_SLICE=1" ) else ( set "DO_SLICE=0" )

set /p "CHOICE_COLORS=Scan and pick new colors? (Y/N): "
if /I "!CHOICE_COLORS!"=="Y" ( set "DO_COLORS=1" ) else ( set "DO_COLORS=0" )
echo.

for /f "usebackq delims=" %%F in ("%FILELIST%") do call :process_one "%%F"
del "%FILELIST%" 2>nul

echo -------------------------------------------
echo Done. Succeeded: !PROCESSED!   Failed: !ERRORS!
echo -------------------------------------------
pause
exit /b

:process_one
set "INPUT=%~1"
set "INPUTDIR=%~dp1"
set "INPUTNAME=%~nx1"
set "INPUTBASE=%~n1"
set "TEMPOUT=%~dp1%~n1_merged_temp.3mf"

set "NESTBASE=!INPUTBASE:~0,-4!Nest"
set "FINALBASE=!INPUTBASE:~0,-4!Final"
set "NESTNAME=!NESTBASE!.3mf"
set "FINAL_PATH=!INPUTDIR!!FINALBASE!.3mf"

for /f %%T in ('powershell -NoProfile -Command "[System.IO.Path]::GetRandomFileName()"') do set "WORK=%TEMP%\merge_work_%%T"

set /a _IDX=PROCESSED+ERRORS+1
echo [!_IDX!/!TOTAL!] Processing: !INPUTNAME!
mkdir "!WORK!" 2>nul

powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUT!', '!WORK!')" >nul 2>&1
if errorlevel 1 ( echo   ERROR: Extract failed. & set /a ERRORS+=1 & goto cleanup )

if "!DO_COLORS!"=="1" (
    echo   Checking/Updating Colors...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update_colors_worker.ps1" -WorkDir "!WORK!"
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkDir "!WORK!" -InputPath "!INPUT!" -OutputPath "!TEMPOUT!" -ReportPath "nul"
if errorlevel 1 ( echo   ERROR: Merge script failed. & set /a ERRORS+=1 & del "!TEMPOUT!" 2>nul & goto cleanup )

ren "!INPUT!" "!NESTNAME!"
ren "!TEMPOUT!" "!INPUTNAME!"

if exist "!FINAL_PATH!" del /f /q "!FINAL_PATH!"
set "WORK_SINGLE=%TEMP%\single_work_%RANDOM%"
mkdir "!WORK_SINGLE!" 2>nul

powershell -NoProfile -Command "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::ExtractToDirectory('!INPUTDIR!!NESTNAME!', '!WORK_SINGLE!')" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0isolate_final_worker.ps1" -WorkDir "!WORK_SINGLE!" -OutputPath "!FINAL_PATH!"
rmdir /s /q "!WORK_SINGLE!" 2>nul

if not exist "!FINAL_PATH!" echo   [!] WARNING: Final.3mf failed to generate.

if "!DO_SLICE!"=="0" goto skip_slicing

set "BAMBU_GUI=C:\Program Files\Bambu Studio\bambu-studio.exe"
set "SLICED_OUT=!INPUTDIR!!INPUTBASE!.gcode.3mf"
set "SLICED_FINAL_TEMP=!INPUTDIR!!FINALBASE!.gcode.3mf"

echo   Slicing Merged Plate...
timeout /t 3 /nobreak > nul

"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SLICED_OUT!" "!INPUTDIR!!INPUTNAME!" > "%TEMP%\slice_log.txt" 2>&1
if not exist "!SLICED_OUT!" (
    echo   WARNING: Slicing failed. Here is the error from Bambu Studio:
    echo   ======================================================================
    type "%TEMP%\slice_log.txt"
    echo   ======================================================================
    del "%TEMP%\slice_log.txt" 2>nul
    goto skip_slicing
)
del "%TEMP%\slice_log.txt" 2>nul

echo   Slicing Isolated Object...
"!BAMBU_GUI!" --debug 3 --no-check --slice 1 --min-save --export-3mf "!SLICED_FINAL_TEMP!" "!FINAL_PATH!" > "%TEMP%\slice_log.txt" 2>&1

if exist "!SLICED_FINAL_TEMP!" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!SLICED_OUT!" -SingleFile "!SLICED_FINAL_TEMP!" -MasterTsvPath "!MASTER_DATA!" >nul 2>&1
    del "!SLICED_FINAL_TEMP!" /q
) else (
    echo   [!] WARNING: Isolated object failed to slice.
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Extract-3MFData.ps1" -InputFile "!SLICED_OUT!" -MasterTsvPath "!MASTER_DATA!" >nul 2>&1
)
del "%TEMP%\slice_log.txt" 2>nul

:skip_slicing
echo   OK --^> Success.
set /a PROCESSED+=1

:cleanup
rmdir /s /q "!WORK!" 2>nul
goto :eof