' Launches Task Inspector with no console window at all. Double-click this file.
Dim sh, dir
Set sh = CreateObject("WScript.Shell")
dir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & dir & "TaskInspector.ps1""", 0, False
