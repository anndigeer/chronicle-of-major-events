@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
py -3 jrjj_v2.py
if errorlevel 1 (
 echo.
 echo 运行出现错误，请检查上方说明。
)
pause
