@echo off
setlocal EnableExtensions
cd /d "%~dp0"
echo ============================================================
echo HipoLaw V16.4.2 - Windows build [V15 Word engine + 3 steps]
echo ============================================================
where py >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python Launcher missing. Install Python 3.10+ for Windows.
  pause
  exit /b 1
)
py -m pip install --upgrade pyinstaller pywin32 openpyxl pymupdf
if errorlevel 1 (
  echo ERROR: Dependency install failed.
  pause
  exit /b 1
)
py -m PyInstaller --noconfirm --clean --onefile --windowed --name HipoLaw_v16_4_2 --icon hipolaw_icon.ico --add-data "assets;assets" --collect-all pymupdf --collect-submodules win32com generate_legal_tables_v16_4_2.py
if errorlevel 1 (
  echo ERROR: PyInstaller failed. Send the last 30 lines of this window.
  pause
  exit /b 1
)
if not exist "dist\HipoLaw_v16_4_2.exe" (
  echo ERROR: Executable missing: dist\HipoLaw_v16_4_2.exe
  pause
  exit /b 1
)
if not exist "dist\HipoLaw_Style_Settings.json" (
  copy /Y "HipoLaw_Style_Settings.json" "dist\HipoLaw_Style_Settings.json" >nul
  if errorlevel 1 (echo ERROR copying settings & pause & exit /b 1)
) else (
  echo Existing style settings retained.
)
if not exist "dist\HipoLaw_ChatGPT_Prompt.txt" (
  copy /Y "HipoLaw_ChatGPT_Prompt.txt" "dist\HipoLaw_ChatGPT_Prompt.txt" >nul
  if errorlevel 1 (echo ERROR copying ChatGPT prompt & pause & exit /b 1)
) else (
  echo Existing editable ChatGPT prompt retained.
)
echo.
echo SUCCESS: dist\HipoLaw_v16_4_2.exe
echo External editable files in dist:
echo   HipoLaw_Style_Settings.json
echo   HipoLaw_ChatGPT_Prompt.txt
echo No rebuild needed to edit these two text files.
echo.
pause
endlocal
