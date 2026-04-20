@echo off
setlocal

set "VBLAUNCHER=%~dp0EDCtoolkit.GUI.vbs"
if exist "%VBLAUNCHER%" (
  wscript.exe "%VBLAUNCHER%"
) else (
  set "GUISCRIPT=%~dp0Scripts\EDCtoolkit\EDCtoolkit.GUI.ps1"
  powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%GUISCRIPT%" -Theme Dark
)

endlocal
