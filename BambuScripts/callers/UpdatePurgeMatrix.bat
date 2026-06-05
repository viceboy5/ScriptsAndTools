@echo off
:: ============================================================
:: UpdatePurgeMatrix.bat  -  Purge matrix updater launcher
::
:: DRAG AND DROP onto this file:
::   - A .3mf file         -> updates that file
::   - Multiple .3mf files -> updates each one
::   - A folder            -> updates all .3mf files inside it
::
:: Uses Tuned_Volume (col E) from PurgeDictionary.csv.
:: Entries with no tuned value are left unchanged.
:: ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\workers\UpdatePurgeMatrix_worker.ps1" %*

pause
