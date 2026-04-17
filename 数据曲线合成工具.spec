# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['h:\\xm4\\xm3\\voltage_current_plotter copy.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.parsing', 'scipy.special._ufuncs_cxx', 'scipy.special._ufuncs', 'matplotlib.backends.backend_tkagg', 'matplotlib.backends.backend_qt5agg', 'PIL._tkinter_finder'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib.tests', 'numpy.random._examples', 'scipy.spatial.transform'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='数据曲线合成工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
