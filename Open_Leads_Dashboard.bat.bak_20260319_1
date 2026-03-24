@echo off
setlocal

set "PORT=8777"
set "ROOT=%~dp0"
set "APPDIR=%~dp0Development\dist"
set "PAGE=APEIRON_BRASIL_-_Opportunities_Management_single_file.html"
set "URL=http://localhost:%PORT%/%PAGE%"
set "BUILDPS1=%ROOT%Development\tools\build-single-file.ps1"
set "REFRESH_PY=%ROOT%Development\tools\refresh_dashboard_from_db.py"
set "AI_HELPER_START_BAT=%ROOT%Development\tools\START_AI_RUNTIME_HELPER.bat"
set "PYTHON_EXE=python"

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python not found. Install Python 3 to run dashboard refresh/server.
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

if exist "%BUILDPS1%" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%BUILDPS1%" >nul
)

if exist "%AI_HELPER_START_BAT%" (
  call "%AI_HELPER_START_BAT%" >nul
)

if not exist "%APPDIR%\%PAGE%" (
  echo Dashboard file not found:
  echo %APPDIR%\%PAGE%
  pause
  exit /b 1
)

set "PID="
for /f "tokens=5" %%p in ('netstat -aon ^| findstr :%PORT% ^| findstr LISTENING') do set "PID=%%p"

if not defined PID (
  start "Opportunities Management Dashboard Server" /min cmd /c "cd /d ""%APPDIR%"" && %PYTHON_EXE% -m http.server %PORT% --bind 127.0.0.1"
  timeout /t 2 >nul
)

set "EDGE1=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
set "EDGE2=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
set "CHROME1=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
set "CHROME2=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"

if exist "%EDGE1%" (
  start "" "%EDGE1%" --app="%URL%" --new-window
  goto :end
)
if exist "%EDGE2%" (
  start "" "%EDGE2%" --app="%URL%" --new-window
  goto :end
)
if exist "%CHROME1%" (
  start "" "%CHROME1%" --app="%URL%" --new-window
  goto :end
)
if exist "%CHROME2%" (
  start "" "%CHROME2%" --app="%URL%" --new-window
  goto :end
)

start "" "%URL%"

:end
endlocal
