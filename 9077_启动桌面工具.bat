@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "9077_桌面工具.py"
if errorlevel 1 pause
