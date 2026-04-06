@echo off
setlocal
set "SCRIPT=%~dp0Scripts\EDCtoolkit\edctoolkit.ps1"

if not exist "%SCRIPT%" (
  echo [ERROR] Could not find: %SCRIPT%
  pause
  exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
  echo.
  echo [INFO] PowerShell exited with code %EXITCODE%.
  pause
)

endlocal
