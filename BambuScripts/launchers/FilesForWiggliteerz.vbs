Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the exact folder where this VBScript lives
strScriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPs1 = objFSO.BuildPath(strScriptFolder, "..\workers\Master-Controller.ps1")

' The "0" at the end tells Windows to launch PowerShell completely hidden
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPs1 & """", 0, False
