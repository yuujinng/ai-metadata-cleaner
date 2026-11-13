@echo off
rem Launcher for Document Metadata Cleaner (ASCII only - do not add non-ASCII text)
cd /d "%~dp0"

set PY=
python -c "import sys" >nul 2>nul
if not errorlevel 1 set PY=python
if defined PY goto ready
py -c "import sys" >nul 2>nul
if not errorlevel 1 set PY=py
if defined PY goto ready

echo [ERROR] Python not found. Please install Python 3 first.
echo Download: https://www.python.org/downloads/
echo (Remember to check "Add Python to PATH" during installation.)
pause
exit /b 1

:ready
%PY% -c "import pypdf, tkinterdnd2" >nul 2>nul
if errorlevel 1 (
    echo Installing support packages: pypdf, tkinterdnd2. Please wait...
    %PY% -m pip install --quiet --disable-pip-version-check pypdf tkinterdnd2
)

%PY% doc_meta_cleaner.py
if errorlevel 1 pause
exit /b 0
