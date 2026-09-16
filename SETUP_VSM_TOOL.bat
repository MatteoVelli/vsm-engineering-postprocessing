@echo off
setlocal

set "PACKAGE_ROOT=%~dp0"
set "TOOL_ROOT=%PACKAGE_ROOT%Tool"

if exist "%TOOL_ROOT%\scripts\client_setup.ps1" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TOOL_ROOT%\scripts\client_setup.ps1" -ProjectRoot "%TOOL_ROOT%"
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PACKAGE_ROOT%scripts\client_setup.ps1" -ProjectRoot "%PACKAGE_ROOT%"
)

if errorlevel 1 (
  echo.
  echo VSM tool setup did not complete successfully. See the message above.
  pause
)

endlocal
