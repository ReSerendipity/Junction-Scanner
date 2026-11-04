@echo off
setlocal
set "SCRIPT=%~dp0JunctionScanner.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
endlocal
