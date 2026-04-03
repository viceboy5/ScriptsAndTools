@echo off
:: ============================================================
:: patch_3mf.bat  –  Bambu .3mf settings patcher launcher
::
:: DRAG AND DROP onto this file:
::   - A folder          -> patches all Final.3mf / Full.3mf inside it
::   - One or more files -> patches only those that match the filter
::
:: Both this .bat and patch_3mf.ps1 must stay in the same folder.
:: ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\AddChinaCoolingSettings.ps1" %*

pause