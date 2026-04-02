Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the exact folder where this VBScript lives
strScriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPs1 = objFSO.BuildPath(strScriptFolder, "CardQueueEditorWPF.ps1")

' Added -STA (CRITICAL FOR WPF) and temporarily changed to WindowStyle Normal
strCmd = "powershell.exe -STA -ExecutionPolicy Bypass -WindowStyle Normal -NoProfile -File """ & strPs1 & """"

' Pass every dragged-and-dropped folder/file directly into the script
For Each strArg In WScript.Arguments
    strCmd = strCmd & " """ & strArg & """"
Next

' Run with window visible (1) instead of hidden (0) to catch any crashes
objShell.Run strCmd, 1, False