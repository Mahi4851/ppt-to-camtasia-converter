@echo off
setlocal EnableExtensions

cd /d "%~dp0"
title PPT to Editable Camtasia Converter

echo ============================================================
echo PPT to Editable Camtasia Converter
echo ============================================================
echo.

set "PY=python"
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"

%PY% --version >nul 2>nul
if errorlevel 1 goto python_error

echo Checking Python requirements...
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo Starting converter UI...
%PY% "scripts\ui_launcher.py"
if errorlevel 1 goto fail
exit /b 0

:python_error
echo.
echo ERROR: Python was not found. Install Python 3.10+ and run this again.
pause
exit /b 1

:fail
echo.
echo ============================================================
echo FAILED
echo ============================================================
echo Check the error above.
echo Requirements:
echo - Python 3.10+
echo - Microsoft PowerPoint installed
echo - Camtasia to open the final .tscproj
pause
exit /b 1
