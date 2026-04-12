@echo off
setlocal

if /I "%~1"=="/elevated" shift

set "GUISCRIPT=%~dp0Scripts\EDCtoolkit\EDCtoolkit.GUI.ps1"
set "CLISCRIPT=%~dp0Scripts\EDCtoolkit\edctoolkit.ps1"
set "SCRIPT=%GUISCRIPT%"
set "ISCLI=0"

if /I "%~1"=="/cli" set "SCRIPT=%CLISCRIPT%" & set "ISCLI=1"
if /I "%~1"=="-cli" set "SCRIPT=%CLISCRIPT%" & set "ISCLI=1"
if /I "%~1"=="--cli" set "SCRIPT=%CLISCRIPT%" & set "ISCLI=1"

if "%ISCLI%"=="0" (
  net session >nul 2>&1
  if not "%ERRORLEVEL%"=="0" (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList '/elevated' -Verb RunAs"
    exit /b
  )
)

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
