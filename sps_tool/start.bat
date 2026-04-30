@echo off
chcp 65001 > nul
cd /d "%~dp0"

:: Check Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았습니다.
    echo https://www.python.org/downloads/ 에서 Python 3.12 이상을 설치하세요.
    echo 설치 시 "Add Python to PATH" 를 반드시 체크하세요.
    pause
    exit /b 1
)

:: Install dependencies if needed
if not exist ".deps_installed" (
    echo [SPS Tool] 패키지 설치 중 (최초 1회)...
    python -m pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
        pause
        exit /b 1
    )
    echo installed > .deps_installed
    echo [SPS Tool] 패키지 설치 완료.
)

echo [SPS Tool] 서버 시작 중... (브라우저에서 http://localhost:5000 열립니다)
echo [SPS Tool] 엔진 설정은 브라우저에서 "설정" 메뉴를 이용하세요.
python app.py
pause
