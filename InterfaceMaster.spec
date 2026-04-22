# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.'), ('logo.png', '.'), ('*.py', '.')],
    hiddenimports=['about', 'contextmenu', 'createnew', 'desktop', 'disk', 'explorer', 'mod', 'sendto', 'str', 'taskbar', 'PIL', 'PIL.Image', 'tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.ttk', 'winreg', 'hashlib', 'ctypes', 'ctypes.wintypes', 'importlib', 'importlib.util'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'pandas'],
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
    name='InterfaceMaster',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    uac_admin=True,
    icon=['icon.ico'],
    manifest='admin_manifest.xml',
)
