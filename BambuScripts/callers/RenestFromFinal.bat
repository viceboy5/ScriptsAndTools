@echo off
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo Drag and drop a Final.3mf onto this script to auto-renest it.
    echo.
    echo The script will look for a sibling _Nest.3mf or _Full.3mf to
    echo read plate transforms from, then produce a new _Renest.3mf.
    pause
    exit /b
)

set "scriptdir=%~dp0"
set "finalpath=%~1"
set "transformsource=%~2"

if not "%transformsource%"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%scriptdir%..\workers\RenestFromFinal_worker.ps1" -FinalPath "%finalpath%" -TransformSourcePath "%transformsource%"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%scriptdir%..\workers\RenestFromFinal_worker.ps1" -FinalPath "%finalpath%"
)

pause
