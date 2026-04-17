@echo off
cd /d "%~dp0"
h:\xm4\venv\Scripts\pyinstaller.exe --onefile --windowed --name "740parser" parser_740_gui.py
echo.
echo done, see dist folder
pause
