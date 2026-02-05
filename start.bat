@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
title 网页表格自动抓取 - 启动器

cd /d "%~dp0"
cls

echo ============================================================
echo(
echo  ######   ######  ########  ######   ##    ##  ##    ##
echo  ##   ## ##    ## ##       ##    ##  ##    ##   ##  ##
echo  ######  ##    ## #####    ##    ##  ##    ##    ####
echo  ##   ## ##    ## ##       ##    ##  ##    ##     ##
echo  ######   ######  ########  ######    ######      ##
echo(
echo              网 页 表 格 自 动 抓 取
echo                  启 动 器
echo(
echo              Author : ADD2048
echo ============================================================
echo(

REM ===== 检测 Python =====
echo [检测] Python 运行环境
where py >nul 2>&1
if %errorlevel%==0 (
    set PY_CMD=py
    echo [通过] Python Launcher 可用
    goto CHECK_PIP
)

where python >nul 2>&1
if %errorlevel%==0 (
    set PY_CMD=python
    echo [通过] Python 可用
    goto CHECK_PIP
)

echo(
echo [错误] 未检测到 Python
echo 请安装 Python 3.10+ 并勾选 Add to PATH
pause
exit /b

:CHECK_PIP
echo(
echo [检测] pip
%PY_CMD% -m pip --version >nul 2>&1
if not %errorlevel%==0 (
    echo [修复] 正在安装 pip
    %PY_CMD% -m ensurepip --upgrade
)

echo(
echo [安装] 项目依赖
%PY_CMD% -m pip install --upgrade pip
%PY_CMD% -m pip install -r requirements.txt

echo(
echo [运行] 启动主程序
%PY_CMD% main.py

echo(
echo ============================================================
echo 程序已结束
echo Author : ADD2048
echo ============================================================
echo(

for /L %%i in (5,-1,1) do (
    echo 窗口将在 %%i 秒后自动关闭...
    timeout /t 1 >nul
)

exit
