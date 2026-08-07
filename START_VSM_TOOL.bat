@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\scripts\client_start.ps1"
if errorlevel 1 (
  echo.
  echo VSM tool did not start successfully. See the message above.
  pause
)
endlocal
