@echo off
setlocal
cd /d "%~dp0"
echo Labmate database comparison only - SQL restore will be skipped
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-LabmateBackup.ps1" -ConfigPath "%~dp0labmate-config.json" -CompareOnly
echo.
if errorlevel 1 (
  echo FAILED. Read the error above.
) else (
  echo SUCCESS. Check the reports folder next to your backup files.
)
pause
