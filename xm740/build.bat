@echo off
chcp 65001 >nul
cd /d "%~dp0"
pyinstaller --onefile --windowed --name "740工装数据解析工具" parser_740_gui.py
echo.
echo 打包完成，输出在 dist 目录
pause
