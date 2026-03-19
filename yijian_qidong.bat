@echo off
title gaokao v3.4
chcp 65001 > nul

cd /d "%~dp0"

echo.
echo ================================================
echo   Ji Lin Gao Kao Zhi Yuan  v3.4
echo ================================================
echo.

echo [0/4] Stop old server (if running)...
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    echo   Killing PID %%a on port 5000...
    taskkill /F /PID %%a > nul 2>&1
)
echo   Port 5000 clear.

echo.
echo [1/4] Check Python...

python --version > nul 2>&1
if not errorlevel 1 (
    set PYTHON=python
    goto PYTHON_OK
)
py --version > nul 2>&1
if not errorlevel 1 (
    set PYTHON=py
    goto PYTHON_OK
)
echo.
echo   [ERROR] Python not found!
echo   Install: https://www.python.org/downloads/
echo   Check box: Add Python to PATH
echo.
pause
exit /b 1

:PYTHON_OK
for /f "tokens=*" %%i in ('%PYTHON% --version 2^>^&1') do echo   %%i

echo.
echo [2/4] Check packages...

set NEED=0
%PYTHON% -c "import flask"    > nul 2>&1
if errorlevel 1 set NEED=1
%PYTHON% -c "import openpyxl" > nul 2>&1
if errorlevel 1 set NEED=1
%PYTHON% -c "import pandas"   > nul 2>&1
if errorlevel 1 set NEED=1
%PYTHON% -c "import numpy"    > nul 2>&1
if errorlevel 1 set NEED=1

if "%NEED%"=="0" (
    echo   All packages ready
    goto DEPS_OK
)

echo.
echo [3/4] Installing packages (1-3 min)...
echo.
%PYTHON% -m pip install flask openpyxl pandas numpy -q
if not errorlevel 1 goto DEPS_OK

%PYTHON% -m pip install flask openpyxl pandas numpy --user -q
if not errorlevel 1 goto DEPS_OK

echo   [ERROR] Install failed!
echo   Run: python -m pip install flask openpyxl pandas numpy
pause
exit /b 1

:DEPS_OK
echo   Packages OK

echo.
echo [4/4] Check files...

if not exist "app.py" (
    echo   [ERROR] app.py not found - run inside gaokao_local folder!
    echo   Current dir: %CD%
    pause
    exit /b 1
)

if not exist "data\*.xlsx" (
    echo   [ERROR] Data file missing - re-extract the zip package
    pause
    exit /b 1
)

echo   app.py      OK
echo   data/*.xlsx OK
if exist "data\df_cache.pkl" (
    echo   cache       OK  (fast)
) else (
    echo   cache       none (first run ~30sec)
)

echo.
echo ================================================
echo   Starting... http://localhost:5000
echo   Close this window OR press Ctrl+C to stop
echo ================================================
echo.

set PYTHONIOENCODING=utf-8
start "" cmd /c "ping 127.0.0.1 -n 4 > nul & start http://localhost:5000"

%PYTHON% app.py

echo.
echo Server stopped.
pause
