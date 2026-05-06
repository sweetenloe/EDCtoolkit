@echo off
setlocal

set "CLISCRIPT=%~dp0Scripts\EDCtoolkit\edctoolkit.ps1"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CLISCRIPT%"

endlocal
