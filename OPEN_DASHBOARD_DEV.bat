@echo off
setlocal
set "ROOT=%~dp0"
set "INDEX=%ROOT%Development\src\dashboard\index.html"
set "REFRESH_PY=%ROOT%Development\tools\refresh_dashboard_from_db.py"
set "AI_HELPER_START_BAT=%ROOT%Development\tools\START_AI_RUNTIME_HELPER.bat"
set "PYTHON_EXE=python"

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python not found. Install Python 3 to run dashboard refresh.
    pause
    exit /b 1
  )
  set "PYTHON_EXE=py -3"
)

if exist "%REFRESH_PY%" (
  call %PYTHON_EXE% "%REFRESH_PY%" >nul
  if errorlevel 1 (
    echo Failed to refresh dashboard payload from DB.
    pause
    exit /b 1
  )
)

if exist "%AI_HELPER_START_BAT%" (
  call "%AI_HELPER_START_BAT%" >nul
)

if not exist "%INDEX%" (
  echo Dashboard file not found:
  echo %INDEX%
  pause
  exit /b 1
)

if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
  start "" "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" --app="file:///%INDEX:\=/%"
  exit /b 0
)

if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
  start "" "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" --app="file:///%INDEX:\=/%"
  exit /b 0
)

start "" "%INDEX%"
