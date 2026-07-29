@echo off
REM Starts a local web server for the Frescoone finance dashboard.
REM Serves this folder on http://localhost:8765 -- reachable only from this PC.
REM Close this window to stop the server.

cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
    set "PYEXE=python"
) else (
    set "PYEXE=C:\Users\admin_fresco\AppData\Local\Programs\Python\Python312\python.exe"
)

if not exist dashboard.html (
    echo dashboard.html not found in this folder.
    echo Run finance_analyzer.py first to generate it.
    pause
    exit /b 1
)

echo Starting Frescoone Finance Dashboard server...
echo Dashboard will open at http://localhost:8765/dashboard.html
echo Close this window any time to stop the server.
echo.

start "" "http://localhost:8765/dashboard.html"
"%PYEXE%" -m http.server 8765 --bind 127.0.0.1
