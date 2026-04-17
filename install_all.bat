@echo off
echo ======================================
echo  一键安装所有子项目依赖
echo ======================================
echo.

echo [1/5] 安装 c2v 依赖...
pip install -r c2v\requirements.txt
echo.

echo [2/5] 安装 echarts 依赖...
pip install -r echarts\requirements.txt
echo.

echo [3/5] 安装 xm5 依赖...
pip install -r xm5\requirements.txt
echo.

echo [4/5] 安装 xmRT 依赖...
pip install -r xmRT\requirements.txt
echo.

echo [5/5] 安装 analyze_voltage_current v2 依赖...
pip install -r "analyze_voltage_current   v2\requirements.txt"
echo.

echo ======================================
echo  全部安装完成！
echo ======================================
pause
