# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [
    ('D:\\AI\\little-helper\\res\\icon.ico', '.'),
    ('D:\\AI\\little-helper\\lib\\lhm', 'lhm'),
]
binaries = []
hiddenimports = [
    'PIL._tkinter_finder',
    'win32timezone',
    'psutil',
    'pynvml',
    'clr',
    'clr._extra',
    'Python.Runtime',
    'config',
    'clipboard_paste',
    'screenshot',
    'hotkey',
    'gpu_power',
    'system_overlay',
    'fan_control',
]
tmp_ret = collect_all('wmi')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['D:\\AI\\little-helper\\src\\main.pyw'],
    pathex=['D:\\AI\\little-helper\\src'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='LittleHelper',
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
    icon=['D:\\AI\\little-helper\\res\\icon.ico'],
)
