@echo off
setlocal

set "GUISCRIPT=%~dp0Scripts\EDCtoolkit\EDCtoolkit.GUI.ps1"

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%GUISCRIPT%"
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
  echo.
  echo [INFO] PowerShell exited with code %EXITCODE%.
  pause
)

endlocal
