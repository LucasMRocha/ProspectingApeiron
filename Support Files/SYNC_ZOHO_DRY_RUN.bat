@echo off
setlocal
cd /d "%~dp0.."

if not exist "Development\tools\zoho-env.ps1" (
  echo Missing Development\tools\zoho-env.ps1
  echo Run:
  echo   powershell -NoProfile -ExecutionPolicy Bypass -File ".\Development\tools\get-zoho-token.ps1" -Code "1000.xxxxx.xxxxx"
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ". .\Development\tools\zoho-env.ps1; python .\Development\tools\sync-zoho-to-db.py --dry-run --skip-build"
if errorlevel 1 (
  echo.
  echo Dry-run failed.
  pause
  exit /b 1
)

echo.
echo Dry-run finished successfully.
pause
endlocal
