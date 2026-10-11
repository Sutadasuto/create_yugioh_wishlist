@echo off
setlocal
cd /d "%~dp0"

set "ENV_PYTHON=%CD%\.venv\Scripts\python.exe"
set "BASE_PYTHON="

for /f "delims=" %%P in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "BASE_PYTHON=%%P"
if not defined BASE_PYTHON (
    for /f "delims=" %%P in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "BASE_PYTHON=%%P"
)

if not defined BASE_PYTHON (
    echo Python 3 was not found. Install Python and make it available through the py launcher or PATH.
    pause
    exit /b 1
)

"%BASE_PYTHON%" -c "import tkinter" >nul 2>nul
if errorlevel 1 (
    echo The installed Python does not include Tkinter. Reinstall Python with Tcl/Tk support.
    pause
    exit /b 1
)

if not exist "%ENV_PYTHON%" (
    "%BASE_PYTHON%" -m venv ".venv"
    if errorlevel 1 (
        echo Could not create the application virtual environment.
        pause
        exit /b 1
    )
)

"%ENV_PYTHON%" -c "import requests, bs4, PIL, openpyxl, tkinter" >nul 2>nul
if errorlevel 1 (
    "%ENV_PYTHON%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Could not install the required Python packages.
        pause
        exit /b 1
    )
)

"%ENV_PYTHON%" yugioh_wishlist.py %*
if errorlevel 1 pause

endlocal