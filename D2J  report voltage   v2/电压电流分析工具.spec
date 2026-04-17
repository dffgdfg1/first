# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['analyze_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.edge',
        'selenium.webdriver.edge.webdriver',
        'selenium.webdriver.edge.service',
        'selenium.webdriver.edge.options',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.service',
        'selenium.webdriver.common.options',
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.support.wait',
        'openpyxl',
        'openpyxl.styles',
        'PIL',
        'PIL.Image',
        'numpy',
        'pandas',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='电压电流分析工具',
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
