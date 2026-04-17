"""
打包脚本 - 将电压电流对比图工具打包成exe文���
使用PyInstaller进行打包
"""
import os
import subprocess
import sys

def build_exe():
    """执行打包命令"""
    # 检查PyInstaller是否安装
    try:
        import PyInstaller
    except ImportError:
        print("正在安装PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller安装完成！\n")

    print("=" * 60)
    print("开始打包电压电流对比图工具")
    print("=" * 60)
    print()

    # PyInstaller命令
    cmd = [
        "pyinstaller",
        "--name=电压电流对比图工具",
        "--onefile",           # 打包成单个exe文件
        "--windowed",          # 无控制台窗口
        "--icon=NONE",         # 图标（如果有ico文件可以指定）
        "--clean",             # 清理临时文件
        "--noconfirm",         # 覆盖输出目录不询问
        "voltage_current_plotter.py"
    ]

    print("执行命令:")
    print(" ".join(cmd))
    print()

    try:
        subprocess.check_call(cmd)
        print()
        print("=" * 60)
        print("打包成功！")
        print("=" * 60)
        print()
        print("exe文件位置: dist\\电压电流对比图工具.exe")
        print()
        print("使用说明:")
        print("1. 将 dist\\电压电流对比图工具.exe 复制到任意位置")
        print("2. 双击运行即可使用")
        print("3. 首次运行会自动创建工作目录和生成文件")
        print()
    except subprocess.CalledProcessError as e:
        print()
        print("=" * 60)
        print("打包失败！")
        print("=" * 60)
        print(f"错误代码: {e.returncode}")
        print()
        return False

    return True

if __name__ == "__main__":
    build_exe()
