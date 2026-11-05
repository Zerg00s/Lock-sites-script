@echo off
cd /d "%~dp0"
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Bulk-DecommissionSpoSite-PnP.ps1"
set RC=%ERRORLEVEL%
pause
exit /b %RC%
