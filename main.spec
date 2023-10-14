# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['D:\\dev\\my\\rep_insulin'],
    datas=[],
    hiddenimports=[],
    hookspath=None,
    runtime_hooks=None
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    a.zipfiles,
    name='report.exe',
    debug=False,
    strip=None,
    upx=True,
    console=False,
    icon=['D:\\dev\\my\\rep_insulin\\icon.ico'],
)
