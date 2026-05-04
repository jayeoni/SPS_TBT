@echo off
cd /d "%~dp0"

set PYTHON=%~dp0python\python-3.13.13-embed-amd64\python.exe

if not exist "%PYTHON%" (
    echo [ERROR] Bundled Python not found: %PYTHON%
    pause
    exit /b 1
)

for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    echo [SPS Tool] Stopping old server PID %%a ...
    taskkill /PID %%a /F > nul 2>&1
)

echo [SPS Tool] Starting server... open http://localhost:5000 in your browser
"%PYTHON%" app.py
pause
