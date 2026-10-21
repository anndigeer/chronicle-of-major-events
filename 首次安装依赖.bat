@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 -m pip install -r requirements_v2.txt
pause
