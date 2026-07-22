# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import os
import sys

datas = [('app_icon.ico', '.')]
python_root = sys.base_prefix
tcl_root = os.path.join(python_root, 'tcl')
python_dlls = os.path.join(python_root, 'DLLs')
datas += [
    (os.path.join(tcl_root, 'tcl8.6'), '_tcl_data'),
    (os.path.join(tcl_root, 'tk8.6'), '_tk_data'),
]
binaries = [
    (os.path.join(python_dlls, '_tkinter.pyd'), '.'),
    (os.path.join(python_dlls, 'tcl86t.dll'), '.'),
    (os.path.join(python_dlls, 'tk86t.dll'), '.'),
]
hiddenimports = ['tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.ttk', 'tkinter.font']
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['Tach_Ghep_Sheet_File_V3.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=['pyinstaller_hooks'],
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
    name='Excel_Toolbox',
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
    icon=['app_icon.ico'],
)
