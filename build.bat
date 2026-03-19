@echo off
chcp 65001 > nul
title 吉林高考志愿规划 · 构建打包

cd /d "%~dp0"

echo.
echo ================================================
echo   吉林高考志愿规划系统 · 构建脚本
echo ================================================
echo.

:: ── [1/4] 检查 Python ──────────────────────────────────
echo [1/4] 检查 Python...
set PYTHON=
python --version > nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" py --version > nul 2>&1 && set PYTHON=py
if "%PYTHON%"=="" (
    echo   [错误] 未找到 Python，请先安装并加入 PATH
    pause & exit /b 1
)
for /f "tokens=*" %%i in ('%PYTHON% --version 2^>^&1') do echo   %%i

:: ── [2/4] 安装/检查 PyInstaller ───────────────────────
echo.
echo [2/4] 检查 PyInstaller...
%PYTHON% -c "import PyInstaller" > nul 2>&1
if errorlevel 1 (
    echo   安装 PyInstaller...
    %PYTHON% -m pip install pyinstaller -q
    if errorlevel 1 ( echo   [错误] PyInstaller 安装失败 & pause & exit /b 1 )
)
echo   PyInstaller OK

:: ── [3/4] PyInstaller 构建 ────────────────────────────
echo.
echo [3/4] PyInstaller 构建（约 3-5 分钟）...
echo.

if exist dist\gaokao rmdir /s /q dist\gaokao
if exist build\gaokao rmdir /s /q build\gaokao

set PYTHONIOENCODING=utf-8
%PYTHON% -m PyInstaller gaokao.spec --noconfirm

if errorlevel 1 (
    echo.
    echo   [错误] PyInstaller 构建失败，请查看上方日志
    pause & exit /b 1
)
echo.
echo   构建完成 → dist\gaokao\

:: ── [4/4] 检查 Inno Setup 并生成安装包 ───────────────
echo.
echo [4/4] 生成安装程序...

set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC=C:\Program Files\Inno Setup 6\ISCC.exe

if "%ISCC%"=="" (
    echo.
    echo   未检测到 Inno Setup 6，跳过安装包生成。
    echo   如需生成安装包：
    echo     1. 下载 Inno Setup: https://jrsoftware.org/isdl.php
    echo     2. 安装后重新运行 build.bat
    echo        或手动用 Inno Setup 打开 installer\setup.iss 编译
    echo.
    echo   当前可直接运行: dist\gaokao\gaokao.exe
    goto DONE
)

"%ISCC%" installer\setup.iss
if errorlevel 1 (
    echo   [错误] Inno Setup 编译失败
    pause & exit /b 1
)
echo.
echo   安装包已生成 → installer\吉林高考志愿规划_v3.4_安装包.exe

:DONE
echo.
echo ================================================
echo   构建完成！
echo ================================================
echo.
echo   直接运行:  dist\gaokao\gaokao.exe
if not "%ISCC%"=="" echo   安装程序:  installer\吉林高考志愿规划_v3.4_安装包.exe
echo.
pause
