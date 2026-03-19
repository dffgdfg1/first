# setup.py
from cx_Freeze import setup, Executable
import sys

# 基础配置
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# 要打包的文件
executables = [Executable("csv_splitter.py", base=base)]

# 依赖项
build_exe_options = {
    "packages": ["pandas", "openpyxl", "tkinter"],
    "excludes": ["matplotlib", "scipy", "numpy.random._examples"],
    "include_files": [],
    "optimize": 2
}

setup(
    name="CSV拆分工具",
    version="1.0",
    description="CSV文件拆分工具，支持超大文件拆分",
    options={"build_exe": build_exe_options},
    executables=executables
)