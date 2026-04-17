@echo off
chcp 65001 >nul
echo ========================================
echo 电压电流对比图工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.7+
    pause
    exit /b 1
)

REM 检查依赖
echo 检查依赖包...
python -c "import tkinter, pandas, plotly" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖包...
    pip install pandas plotly numpy openpyxl
)

echo.
echo 正在启动应用...
echo.
python voltage_current_plotter.py

if errorlevel 1 (
    echo.
    echo 应用运行出错！
    pause
)
