@echo off
setlocal
set "SCRIPT=%~dp0import_carteira_pipeline.py"

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python not found.
    pause
    exit /b 1
  )
  py -3 "%SCRIPT%" dry-run
) else (
  python "%SCRIPT%" dry-run
)
echo.
pause
