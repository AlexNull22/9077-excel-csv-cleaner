@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "9077_表格清洗.py" "9077_示例订单.csv" "9077_清洗规则.json" "9077_本次演示结果"
pause
