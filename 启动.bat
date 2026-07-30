@echo off
setlocal
if not "%~1"=="__KEEP__" (
    cmd /k "%~f0" __KEEP__
    exit /b
)
:main
cd /d "%~dp0"
set "LOG=%~dp0launch.log"
echo [%date% %time%] Launcher started > "%LOG%"
echo ================================================== >> "%LOG%"
echo   Data Masking Tool - Launcher
echo ================================================== >> "%LOG%"
echo. >> "%LOG%"
echo ==================================================
echo   Data Masking Tool - Launcher
echo ==================================================
echo.
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "VENV_ST=%~dp0.venv\Scripts\streamlit.exe"

REM ---------- Primary path: if .venv is ready, launch directly ----------
if exist "%VENV_ST%" (
    echo [INFO] Using existing .venv >> "%LOG%"
    echo [INFO] Using existing .venv
    goto :launch
)

REM ---------- Fallback: build .venv with a Python 3.9+ ----------
echo [INFO] .venv not ready, locating Python 3.9+ ... >> "%LOG%"
echo [INFO] .venv not ready, locating Python 3.9+ ...
set "PY="
py -3 "%~dp0findpy.py" >nul 2>nul
if %errorlevel%==0 (
    for /f "delims=" %%L in ('py -3 "%~dp0findpy.py" 2^>nul') do set "PY=%%L"
) else (
    for /f "delims=" %%L in ('python "%~dp0findpy.py" 2^>nul') do set "PY=%%L"
)
if not defined PY (
    echo [ERROR] Python 3.9+ not found. Install Python and tick "Add to PATH". >> "%LOG%"
    echo [ERROR] Python 3.9+ not found. Install Python 3.9+ and tick "Add to PATH".
    goto :done
)
if "%PY%"=="NONE" (
    echo [ERROR] No Python 3.9+ found on this machine. >> "%LOG%"
    echo [ERROR] No Python 3.9+ found on this machine.
    goto :done
)
echo [INFO] Using Python: %PY% >> "%LOG%"
echo [INFO] Using Python: %PY%

if not exist "%VENV_PY%" (
    echo [INFO] Building .venv ... >> "%LOG%"
    echo [INFO] Building .venv ...
    "%PY%" -m venv .venv >> "%LOG%" 2>&1
    if errorlevel 1 (
        echo [ERROR] Failed to create .venv. Check Python install. >> "%LOG%"
        echo [ERROR] Failed to create .venv. Check Python install.
        goto :done
    )
)
"%VENV_PY%" -m pip install --upgrade pip >> "%LOG%" 2>&1
"%VENV_PY%" -m pip install -r requirements.txt >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] Dependency install failed. Check network. >> "%LOG%"
    echo [ERROR] Dependency install failed. Check network connection.
    goto :done
)

:launch
taskkill /F /IM streamlit.exe >nul 2>nul
echo [INFO] Starting, open http://localhost:8501 in your browser >> "%LOG%"
echo.
echo [INFO] Starting, open http://localhost:8501 in your browser
echo [INFO] To stop: close this window
echo.
"%VENV_PY%" -m streamlit run mask_tool.py --server.headless true >> "%LOG%" 2>&1
echo [INFO] App exited (code %ERRORLEVEL%). >> "%LOG%"
:done
echo.
echo ==================================================
echo If launch failed, check launch.log in this folder.
echo ==================================================
pause
