@echo off
set SCRIPT_PATH="C:\Users\Miguel\OneDrive\Documents\GitHub\monsterhunterloggerdogger\scripts\log_system_monitor.ps1"

:: Check if running as admin, if not relaunch as admin
powershell.exe -Command "Start-Process -Verb RunAs -FilePath 'powershell.exe' -ArgumentList '-ExecutionPolicy Bypass -NoProfile -File %SCRIPT_PATH%'"

exit /b