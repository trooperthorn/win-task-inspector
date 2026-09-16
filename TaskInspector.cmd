@echo off
rem Launches Task Inspector without a lingering console. Double-click to run.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0TaskInspector.ps1"
