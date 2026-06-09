@echo off
:: ============================================================
:: ResetPurgeMatrix.bat  -  Reset purge volumes to Bambu defaults
::
:: DRAG AND DROP onto this file:
::   - A .3mf file         -> resets that file
::   - Multiple .3mf files -> resets each one
::   - A folder            -> resets all project .3mf files inside it (recursive)
::
:: For each file:
::   1. Removes the custom flush_volumes_matrix from project_settings.config
::   2. Re-exports through Bambu Studio so it recomputes defaults
::
:: .gcode.3mf sliced outputs are skipped automatically.
:: ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\ResetPurgeMatrix_worker.ps1" -Paths %*

pause
