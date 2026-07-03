@echo off
echo Starting Portfolio Dashboard...
echo.
echo DATA SOURCE: Google Sheet "Portfolio Dashboard" (Holdings / Watchlist /
echo Fixed Income tabs) is the source of truth. Edit it on the dashboard or
echo directly in Google Sheets. The local Excel files are only a fallback.
echo.
echo The dashboard will open in your browser automatically.
echo Close this window to stop the dashboard.
echo.

REM Find Python
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

if "%PYEXE%"=="" (
    python --version >nul 2>&1 && set PYEXE=python
)

if "%PYEXE%"=="" (
    echo ERROR: Python not found. Please run setup.bat first.
    pause
    exit /b 1
)

cd /d "%~dp0"
REM Open the dashboard in Microsoft Edge once the server has started.
start "" cmd /c "timeout /t 5 >nul & start msedge http://localhost:8501"
%PYEXE% -m streamlit run app.py --server.headless true --browser.gatherUsageStats false
pause
