# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('Models', 'Models'), ('Functions', 'Functions'), ('Screens', 'Screens'), ('Main', 'Main')]
binaries = [('dlls/python3.dll', '.'), ('dlls/python313.dll', '.'), ('dlls/vcruntime140.dll', '.'), ('dlls/vcruntime140_1.dll', '.'), ('dlls/python3.dll', '.'), ('dlls/python313.dll', '.'), ('dlls/vcruntime140.dll', '.'), ('dlls/vcruntime140_1.dll', '.'), ('dlls/sqlite3.dll', '.'), ('dlls/tcl86t.dll', '.'), ('dlls/tk86t.dll', '.'), ('dlls/tcldde14.dll', '.'), ('dlls/tclreg13.dll', '.'), ('dlls/msvcp140-d64049c6e3865410a7dda6a7e9f0c575.dll', '.'), ('dlls/msvcp140-0f2ea95580b32bcfc81c235d5751ce78.dll', '.'), ('dlls/arrow_python.dll', '.'), ('dlls/arrow_python_flight.dll', '.'), ('dlls/arrow_python_parquet_encryption.dll', '.'), ('dlls/msvcp140-c691275f5538a516494cdd39ad3f5ca6.dll', '.'), ('dlls/libtkdnd2.9.3.dll', '.'), ('dlls/libtkdnd2.9.4.dll', '.'), ('dlls/libtkdnd2.9.4.dll', '.'), ('dlls/msvcp140-d64049c6e3865410a7dda6a7e9f0c575.dll', '.'), ('dlls/msvcp140-0f2ea95580b32bcfc81c235d5751ce78.dll', '.'), ('dlls/arrow_python.dll', '.'), ('dlls/arrow_python_flight.dll', '.'), ('dlls/arrow_python_parquet_encryption.dll', '.'), ('dlls/msvcp140-c691275f5538a516494cdd39ad3f5ca6.dll', '.'), ('dlls/libtkdnd2.9.3.dll', '.'), ('dlls/libtkdnd2.9.4.dll', '.'), ('dlls/libtkdnd2.9.4.dll', '.')]
hiddenimports = ['openpyxl', 'tkcalendar', 'tkinter']
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkcalendar')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['Main\\MainController.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    a.binaries,
    a.datas,
    [],
    name='WorkerAssignment_Full',
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
