@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
pushd "%~dp0" || (
    echo 无法进入项目目录。
    exit /b 1
)

rem Windows PowerShell / CMD entry. The .sh file is ONLY for Linux/Kylin.
rem Prefer the interpreter installed in this repository's virtual environment.
set "PYTHON="
if exist ".venv\Scripts\python.exe" set "PYTHON=%CD%\.venv\Scripts\python.exe"
if not defined PYTHON (
    where python >nul 2>&1
    if not errorlevel 1 set "PYTHON=python"
)
if not defined PYTHON (
    where py >nul 2>&1
    if not errorlevel 1 set "PYTHON=py"
)
if not defined PYTHON (
    echo 没有找到 Python。请先安装 Python 3.10+，或者创建本项目的 .venv。
    popd
    exit /b 1
)
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if "%~1"=="" goto launch_default
if /I "%~1"=="start" goto launch_explicit
if /I "%~1"=="check" goto diagnose
if /I "%~1"=="help" goto usage
if /I "%~1"=="--help" goto usage
if /I "%~1"=="-h" goto usage
echo 未识别的命令：%~1
goto usage_error

:launch_default
echo [Windows] 启动市委大事记浏览器操作台……
"%PYTHON%" -u webui.py
set "RET=%ERRORLEVEL%"
goto finish

:launch_explicit
echo [Windows] 启动市委大事记浏览器操作台……
rem Supported webui.py options: --no-browser and --port NUMBER.
"%PYTHON%" -u webui.py %2 %3 %4 %5 %6 %7 %8 %9
set "RET=%ERRORLEVEL%"
goto finish

:diagnose
echo [Windows] 检查项目 Python 和采集依赖……
"%PYTHON%" kylin_check.py
set "RET=%ERRORLEVEL%"
goto finish

:usage
echo 市委大事记：Windows 图形界面
echo.
echo   .\市委大事记.cmd               启动网页图形界面，自动打开浏览器
echo   .\市委大事记.cmd check         检查 Python 和采集依赖
echo   .\市委大事记.cmd start         启动网页，可附 --no-browser --port 8765
echo   .\市委大事记.cmd help          查看用法
echo.
echo 在 Windows 上请使用 .cmd，银河麒麟和 Linux 使用 .sh。
echo 也可以运行：python webui.py
echo 如需安装依赖：.\.venv\Scripts\python.exe -m pip install -r requirements.txt
set "RET=0"
goto finish

:usage_error
echo 请执行 .\市委大事记.cmd help 查看支持的命令。
set "RET=2"
goto finish

:finish
popd
exit /b %RET%
