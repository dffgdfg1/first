@echo off
chcp 65001
echo ====================================
echo ECharts编辑器打包工具
echo ====================================
echo.

echo [1/4] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "ECharts编辑器.exe" del /f /q "ECharts编辑器.exe"
echo 清理完成
echo.

echo [2/4] 检查依赖...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 未安装PyInstaller，正在安装...
    pip install pyinstaller
)
echo 依赖检查完成
echo.

echo [3/4] 开始打包...
pyinstaller "ECharts编辑器.spec" --clean
echo.

if exist "dist\ECharts编辑器.exe" (
    echo [4/4] 打包成功！
    echo.
    echo 可执行文件位置: dist\ECharts编辑器.exe
    echo.
    echo 是否打开输出目录？
    pause
    explorer dist
) else (
    echo [4/4] 打包失败！
    echo 请检查错误信息
    pause
)
