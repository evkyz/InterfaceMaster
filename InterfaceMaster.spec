# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.'), ('logo.png', '.'), ('*.py', '.')],
    hiddenimports=['PIL', 'PIL._imaging', 'PIL.Image', 'PIL.ImageFile', 'PIL.ImageOps', 'PIL.ImageTk', 'winreg', 'hashlib', 'tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.ttk', 'ctypes', 'ctypes.wintypes', 'os', 'sys', 'subprocess', 'types'],
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
    name='InterfaceMaster',
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
    version='version_info.txt',
    uac_admin=True,
    icon=['icon.ico'],
    manifest='admin_manifest.xml',
)
