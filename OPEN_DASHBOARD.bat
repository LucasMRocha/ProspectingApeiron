@echo off
setlocal

set "PORT=8777"
set "APPDIR=%~dp0Development\dist"
set "PAGE=Prospecting_Dashboard_single_file.html"
set "URL=http://localhost:%PORT%/%PAGE%"

if not exist "%APPDIR%\%PAGE%" (
  echo Dashboard file not found:
  echo %APPDIR%\%PAGE%
  pause
  exit /b 1
)

set "PID="
for /f "tokens=5" %%p in ('netstat -aon ^| findstr :%PORT% ^| findstr LISTENING') do set "PID=%%p"

if not defined PID (
  start "Prospecting Dashboard Server" /min cmd /c "cd /d ""%APPDIR%"" && python -m http.server %PORT% --bind 127.0.0.1"
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
