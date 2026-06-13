@echo off
setlocal
cd /d "%~dp0"
title Amazing Mart Dashboard

echo.
echo ========================================
echo   Amazing Mart Dashboard Launcher
echo ========================================
echo.

set "PYTHON_EXE="

where py >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%P in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%P"
)

if not defined PYTHON_EXE (
    where python >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%P in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%P"
    )
)

if not defined PYTHON_EXE (
    echo ERROR: Python was not found.
    echo Install Python 3.10 or newer from python.org.
    echo During installation, check "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo ERROR: Python 3.10 or newer is required.
    "%PYTHON_EXE%" --version
    echo.
    pause
    exit /b 1
)

if not exist "Amazing_Mart_Dataset_1.xlsx" (
    echo ERROR: Amazing_Mart_Dataset_1.xlsx is missing.
    echo Keep the Excel file in the same folder as this launcher.
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating a private Python environment...
    "%PYTHON_EXE%" -m venv ".venv"
    if errorlevel 1 goto :failed
)

".venv\Scripts\python.exe" -c "import streamlit, pandas, plotly, openpyxl, numpy" >nul 2>&1
if errorlevel 1 (
    echo Installing dashboard packages...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    if errorlevel 1 goto :failed
    ".venv\Scripts\python.exe" -m pip install -r "requirements.txt"
    if errorlevel 1 goto :failed
)

echo Starting the dashboard...
echo A browser window should open automatically.
echo Keep this window open while using the dashboard.
echo.
".venv\Scripts\python.exe" -m streamlit run "amazing_mart_streamlit.py"
if errorlevel 1 goto :failed
exit /b 0

:failed
echo.
echo ERROR: The dashboard could not start.
echo Check the error shown above. Internet access is required during the first setup.
echo.
pause
exit /b 1
