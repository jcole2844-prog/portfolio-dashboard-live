@echo off
echo ============================================================
echo  Portfolio Dashboard — First-Time Setup
echo ============================================================
echo.

REM Check if Python is installed via winget-installed location
set PYEXE=
for %%P in (
    "C:\Users\Julie\AppData\Local\Programs\Python\Python313\python.exe"
    "C:\Users\Julie\AppData\Local\Programs\Python\Python312\python.exe"
    "C:\Users\Julie\AppData\Local\Programs\Python\Python311\python.exe"
    "C:\Users\Julie\AppData\Local\Programs\Python\Python310\python.exe"
    "C:\ProgramData\Python\Python313\python.exe"
    "C:\ProgramData\Python\Python312\python.exe"
) do (
    if exist %%P set PYEXE=%%P
)

REM Also try PATH
if "%PYEXE%"=="" (
    python --version >nul 2>&1 && set PYEXE=python
)

if "%PYEXE%"=="" (
    echo Python not found. Installing Python 3.12 via winget...
    winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    if errorlevel 1 (
        echo.
        echo ERROR: winget install failed. Please download Python from:
        echo   https://www.python.org/downloads/
        echo Then re-run this script.
        pause
        exit /b 1
    )
    REM Refresh PATH
    call refreshenv >nul 2>&1
    set PYEXE=python
    echo Python installed successfully.
)

echo Python found: %PYEXE%
echo.
echo Installing required packages...
%PYEXE% -m pip install --upgrade pip --quiet
%PYEXE% -m pip install -r "%~dp0requirements.txt" --quiet

if errorlevel 1 (
    echo.
    echo ERROR: Package installation failed. Please check your internet connection.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Setup complete! Run the dashboard with:  run_dashboard.bat
echo ============================================================
pause
