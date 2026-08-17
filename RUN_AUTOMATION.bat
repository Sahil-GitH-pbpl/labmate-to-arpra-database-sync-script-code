@echo off
setlocal
cd /d "%~dp0"
echo Labmate backup restore and comparison automation
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-LabmateBackup.ps1" -ConfigPath "%~dp0labmate-config.json"
echo.
if errorlevel 1 (
  echo FAILED. Read the error above and check the latest report in the reports folder.
) else (
  echo SUCCESS. Check the reports folder next to your backup files.
)
pause
