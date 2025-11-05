# -*- mode: python ; coding: utf-8 -*-

# 是否加密
block_cipher = None

py_files = [
    '..\\modules\\cm_sop_translate\\main.py'
]

add_files = [
    ('D:\\Code\\Project\\tools\\modules\\cm_sop_translate\\conf\\json', 'config'),
]

a = Analysis(
    py_files,
    pathex=["D:/Code/Config/Private"],
    binaries=[],
    datas=add_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
#    icon = 'C:\\Users\\shawn\\Downloads\\favicon.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='translator',
)
