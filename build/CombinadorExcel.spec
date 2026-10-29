# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:/Users/M2PC86/Desktop/CORPOSALUD-TOOLS-main/excel_merger.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl', 'xlrd'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['jupyter', 'jupyterlab', 'notebook', 'ipykernel', 'IPython', 'matplotlib', 'scipy', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'pytest', 'sphinx', 'docutils'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    name='CombinadorExcel',
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
