@echo off
chcp 65001 > nul
cd /d "%~dp0"

:: Use bundled portable Python (persists across reboots)
set PYTHON=%~dp0python\python-3.13.13-embed-amd64\python.exe

if not exist "%PYTHON%" (
    echo [오류] 번들 Python을 찾을 수 없습니다: %PYTHON%
    echo sps_tool\python\python-3.13.13-embed-amd64\ 폴더가 있는지 확인하세요.
    pause
    exit /b 1
)

:: Kill any existing Flask process on port 5000 to avoid stale-server issues
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    echo [SPS Tool] 기존 서버 종료 중 (PID %%a)...
    taskkill /PID %%a /F > nul 2>&1
)

echo [SPS Tool] 서버 시작 중... (브라우저에서 http://localhost:5000 열립니다)
echo [SPS Tool] 엔진 설정은 브라우저에서 "설정" 메뉴를 이용하세요.
"%PYTHON%" app.py
pause
