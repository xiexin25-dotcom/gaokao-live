@echo off
chcp 65001 > nul
echo.
echo  =============================================
echo    吉林省高考志愿规划系统 · 本地版 v3.4
echo  =============================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo  [错误] 未找到 Python，请先安装 Python 3.9+
    echo  下载: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo  正在安装依赖包（首次启动需要约1分钟）...
python -m pip install flask openpyxl pandas numpy
if errorlevel 1 (
    echo.
    echo  [提示] 如安装失败，请手动执行：
    echo    python -m pip install flask openpyxl pandas numpy
    echo  或尝试：
    echo    python -m pip install flask openpyxl pandas numpy --user
    pause
    exit /b 1
)

echo.
echo  正在启动服务...
echo  浏览器将自动打开 http://localhost:5000
echo  按 Ctrl+C 停止服务
echo.
start "" http://localhost:5000
set PYTHONIOENCODING=utf-8
python app.py
pause
