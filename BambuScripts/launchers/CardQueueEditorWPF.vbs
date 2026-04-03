Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the exact folder where this VBScript lives
strScriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPs1 = objFSO.BuildPath(strScriptFolder, "..\workers\CardQueueEditorWPF.ps1")

strCmd = "powershell.exe -STA -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File """ & strPs1 & """"

' Pass every dragged-and-dropped folder/file directly into the script
For Each strArg In WScript.Arguments
    strCmd = strCmd & " """ & strArg & """"
Next

objShell.Run strCmd, 0, False