Set objShell = CreateObject("WScript.Shell")
' The "0" at the end tells Windows to launch PowerShell completely hidden
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""Master-Controller.ps1""", 0, False