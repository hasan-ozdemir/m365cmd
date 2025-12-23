@echo off
setlocal

set "M365CMD_ROOT=%~dp0"
set "M365CMD_MAIN=%M365CMD_ROOT%m365cmd-main.ps1"

where pwsh >nul 2>&1
if errorlevel 1 (
  echo ERROR: PowerShell Core is required.
  exit /b 1
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%M365CMD_MAIN%" %*
exit /b
