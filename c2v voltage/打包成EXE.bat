@echo off
chcp 65001 >nul
echo ========================================
echo 电压电流对比图工具 - EXE打包工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.7+
    pause
    exit /b 1
)

echo 正在安装/检查PyInstaller...
python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

if errorlevel 1 (
    echo PyInstaller安装失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 开始打包...
echo ========================================
echo.

REM 执行打包脚本
python build_exe.py

if errorlevel 1 (
    echo.
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo EXE文件已生成在 dist 文件夹中
echo 文件名: 电压电流对比图工具.exe
echo.
echo 您可以:
echo 1. 进入 dist 文件夹
echo 2. 将 exe 文件复制到任意位置
echo 3. 双击运行即可使用
echo.

pause

REM 询问是否打开dist文件夹
set /p opendir="是否打开dist文件夹? (Y/N): "
if /i "%opendir%"=="Y" (
    explorer dist
)
