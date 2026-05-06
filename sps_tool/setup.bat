@echo off
cd /d "%~dp0"

set EMBED_DIR=%~dp0python\python-3.13.13-embed-amd64
set PYTHON_VER=3.13.13
set STDLIB_ZIP=%EMBED_DIR%\python313.zip
set PYTHON_EXE=%EMBED_DIR%\python.exe
set DOWNLOAD_URL=https://www.python.org/ftp/python/%PYTHON_VER%/python-%PYTHON_VER%-embed-amd64.zip
set TEMP_ZIP=%TEMP%\python-embed-temp.zip

echo ============================================
echo  SPS Tool - Environment Setup / Repair
echo ============================================
echo.

rem --- Check python.exe ---
if not exist "%PYTHON_EXE%" (
    echo [ERROR] python.exe not found at:
    echo   %PYTHON_EXE%
    echo The python\ folder may be corrupted or missing entirely.
    echo Re-clone the repository and place the python\ folder back.
    goto fail
)
echo [OK] python.exe found

rem --- Check python313.zip and repair if missing ---
if exist "%STDLIB_ZIP%" (
    echo [OK] python313.zip found
    goto check_packages
)

echo [REPAIR] python313.zip is missing - downloading from python.org...
echo   Fetching: %DOWNLOAD_URL%
powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%TEMP_ZIP%' -UseBasicParsing" 2>&1
if errorlevel 1 (
    echo [ERROR] Download failed. Check your internet connection and try again.
    goto fail
)

echo [REPAIR] Extracting python313.zip from downloaded package...
powershell -Command ^
  "Add-Type -Assembly System.IO.Compression.FileSystem;" ^
  "$z = [System.IO.Compression.ZipFile]::OpenRead('%TEMP_ZIP%');" ^
  "$e = $z.Entries | Where-Object { $_.Name -eq 'python313.zip' };" ^
  "[System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, '%STDLIB_ZIP%', $true);" ^
  "$z.Dispose()"
if errorlevel 1 (
    echo [ERROR] Extraction failed.
    del "%TEMP_ZIP%" 2>nul
    goto fail
)
del "%TEMP_ZIP%" 2>nul
echo [OK] python313.zip restored

:check_packages
rem --- Verify core packages are importable ---
echo.
echo Checking installed packages...
"%PYTHON_EXE%" -c "import flask, docx, openpyxl, pandas, dotenv, anthropic; print('[OK] All packages verified')" 2>nul
if errorlevel 1 (
    echo [REPAIR] Some packages missing - reinstalling from requirements.txt...
    "%PYTHON_EXE%" -m pip install -r requirements.txt --no-warn-script-location
    if errorlevel 1 (
        echo [ERROR] Package installation failed.
        goto fail
    )
    echo [OK] Packages reinstalled
)

echo.
echo ============================================
echo  Setup complete. Run start.bat to launch.
echo ============================================
pause
exit /b 0

:fail
echo.
echo Setup did not complete successfully. See errors above.
pause
exit /b 1
