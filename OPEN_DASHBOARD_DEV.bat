@echo off
setlocal
set "ROOT=%~dp0"
set "INDEX=%ROOT%src\dashboard\index.html"

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
