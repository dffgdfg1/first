# build_exe.py
import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))
script_path = os.path.join(current_dir, "voltage_current_plotter copy.py")

# PyInstaller 参数
args = [
    script_path,
    "--onefile",  # 打包成单个exe文件
    "--windowed",  # 不显示控制台窗口
    "--name=数据曲线合成工具",  # 可执行文件名
    "--clean",  # 清理临时文件
    "--noconfirm",  # 覆盖现有文件时不确认
    
    # 添加隐藏的导入
    "--hidden-import=pandas._libs.tslibs.timedeltas",
    "--hidden-import=pandas._libs.tslibs.np_datetime",
    "--hidden-import=pandas._libs.tslibs.parsing",
    "--hidden-import=scipy.special._ufuncs_cxx",
    "--hidden-import=scipy.special._ufuncs",
    "--hidden-import=matplotlib.backends.backend_tkagg",
    "--hidden-import=matplotlib.backends.backend_qt5agg",
    "--hidden-import=PIL._tkinter_finder",
    
    # 排除不必要的包
    "--exclude-module=matplotlib.tests",
    "--exclude-module=numpy.random._examples",
    "--exclude-module=scipy.spatial.transform",
    
    # 添加数据文件（如果有图标）
    # "--add-data=icon.ico;.",
]

# 如果你有图标文件，添加以下参数
# args.extend(["--icon=icon.ico"])

# 运行打包
PyInstaller.__main__.run(args)

print("打包完成！可执行文件在 dist 文件夹中。")