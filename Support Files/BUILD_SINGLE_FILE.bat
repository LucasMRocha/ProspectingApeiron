@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\Development\tools\build-single-file.ps1"
if errorlevel 1 (
  echo Failed to build single-file dashboard.
  pause
  exit /b 1
)
echo.
echo Build complete.
pause
