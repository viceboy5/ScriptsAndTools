Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the exact folder where this VBScript lives
strScriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPs1 = objFSO.BuildPath(strScriptFolder, "Card-Editor-Skeleton.ps1")

' Build the silent PowerShell command
strCmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File """ & strPs1 & """"

' Pass every dragged-and-dropped folder/file directly into the script
For Each strArg In WScript.Arguments
    strCmd = strCmd & " """ & strArg & """"
Next

' Run it completely hidden (0) and don't lock up the VBScript waiting for it to finish (False)
objShell.Run strCmd, 0, False