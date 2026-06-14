@echo off
:: ============================================================
:: ResetMetadata.bat  -  Strip and reset .3mf metadata
::
:: DRAG AND DROP onto this file:
::   - A .3mf file         -> resets that file
::   - Multiple .3mf files -> resets each one
::   - A folder            -> resets all project .3mf files inside it (recursive)
::
:: For each file:
::   - Removes flush_volumes_matrix (purge volumes reset to Bambu defaults)
::   - Deletes all embedded .png thumbnails
::   - Re-exports through Bambu Studio to write a clean file
::
:: .gcode.3mf sliced outputs are skipped automatically.
:: ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\ResetMetadata_worker.ps1" -Paths %*

pause
