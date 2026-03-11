@echo off
chcp 65001 >nul
echo ========================================
echo 电压电流分析工具 - 打包程序
echo ========================================
echo.

REM 检查是否存在虚拟环境
if exist "venv\Scripts\activate.bat" (
    echo [1/4] 激活虚拟环境...
    call venv\Scripts\activate.bat
) else (
    echo 警告: 未找到虚拟环境，使用系统Python
)

echo.
echo [2/4] 检查PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 未安装PyInstaller，正在安装...
    pip install pyinstaller
)

echo.
echo [3/4] 开始打包...
pyinstaller --noconfirm "电压电流分析工具.spec"

echo.
echo [4/4] 清理临时文件...
if exist "build" rmdir /s /q build
if exist "__pycache__" rmdir /s /q __pycache__

echo.
echo ========================================
echo 打包完成！
echo 可执行文件位置: dist\电压电流分析工具.exe
echo ========================================
echo.
pause
