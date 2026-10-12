@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "MOTION=%~1"
set "REVIEW=%~2"
if "%MOTION%"=="" (
  echo Drag your selected Word .docx into this window or type the path.
  set /p "MOTION=Motion .docx path: "
)
if "%REVIEW%"=="" (
  echo Drag your approved Excel .xlsx into this window or type the path.
  set /p "REVIEW=Approved .xlsx path: "
)
set "MOTION=%MOTION:"=%"
set "REVIEW=%REVIEW:"=%"
py -u generate_legal_tables_v16_4_8.py --preflight "%MOTION%" "%REVIEW%"
pause
endlocal
