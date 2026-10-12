@echo off
cd /d "%~dp0"
if not exist "dist\HipoLaw_v16_4_9.exe" (
  echo Please run BUILD_V16_4_9.bat first.
  pause
  exit /b 1
)
start "" "%~dp0dist\HipoLaw_v16_4_9.exe"
