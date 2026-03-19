@echo off
setlocal

set "ROOT=%~dp0"
set "HELPER=%ROOT%ai_runtime_helper.py"
set "PORT=8011"
set "PYTHON_EXE=python"
set "USE_PY_LAUNCHER=0"

if not exist "%HELPER%" exit /b 0

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python not found. AI runtime helper not started.
    exit /b 1
  )
  set "PYTHON_EXE=py"
  set "USE_PY_LAUNCHER=1"
)

set "PID="
for /f "tokens=5" %%p in ('netstat -aon ^| findstr :%PORT% ^| findstr LISTENING') do set "PID=%%p"
if defined PID exit /b 0

if "%USE_PY_LAUNCHER%"=="1" (
  start "Apeiron AI Runtime Helper" /min %PYTHON_EXE% -3 "%HELPER%"
) else (
  start "Apeiron AI Runtime Helper" /min %PYTHON_EXE% "%HELPER%"
)
exit /b 0
