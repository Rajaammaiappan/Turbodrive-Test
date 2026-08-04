@echo off
title Convert Python to EXE - ALTEN EFS

:: ════════════════════════════════════════════════════════════════
::  UNIVERSAL Python → EXE Converter
::  Usage : Put this .bat in the same folder as your .py file
::          Edit ONLY the line below with your script name
:: ════════════════════════════════════════════════════════════════

set SCRIPT=your_script_name.py

:: ── DO NOT EDIT BELOW THIS LINE ─────────────────────────────────

set PYTHON=C:\ProgramData\Anaconda3\python.exe

echo.
echo  Building EXE for: %SCRIPT%
echo  Please wait...
echo.

"%PYTHON%" -m PyInstaller --onefile --console ^
  --hidden-import=win32com.client ^
  --hidden-import=win32api ^
  --hidden-import=win32con ^
  --hidden-import=win32gui ^
  --hidden-import=pywintypes ^
  --hidden-import=openpyxl ^
  --hidden-import=flask ^
  --hidden-import=sqlite3 ^
  --hidden-import=pkg_resources.py2_warn ^
  %SCRIPT%

echo.
echo  DONE. Your EXE is in the dist\ folder.
echo.
pause
