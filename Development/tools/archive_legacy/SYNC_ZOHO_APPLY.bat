@echo off
setlocal
cd /d "%~dp0"

if not exist "tools\zoho-env.ps1" (
  echo Missing tools\zoho-env.ps1
  echo Run:
  echo   powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\get-zoho-token.ps1" -Code "1000.xxxxx.xxxxx"
  pause
  exit /b 1
)

echo This will update:
echo   - Apeiron_BR_Gestao_Comercial.xlsx
echo   - src\dashboard\data\leads.js
echo   - dist\APEIRON_BRASIL_-_Opportunities_Management_single_file.html
echo.
set /p CONFIRM=Type YES to continue: 
if /I not "%CONFIRM%"=="YES" (
  echo Cancelled.
  pause
  exit /b 0
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ". .\tools\zoho-env.ps1; python .\tools\sync-zoho-to-db.py"
if errorlevel 1 (
  echo.
  echo Sync failed.
  pause
  exit /b 1
)

echo.
echo Sync completed successfully.
pause
endlocal
