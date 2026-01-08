# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['raidstat_py/main.py'],
    pathex=['raidstat_py'],
    binaries=[],
    datas=[
        ('raidstat_py/bin', 'bin'),
        ('raidstat_py/gui/assets/icon.ico', 'gui/assets'),
    ],
    hiddenimports=[
        'PIL',
        'PIL._tkinter_finder',
        'cv2',
        'numpy',
        'thefuzz',
        'pytesseract',
        'openpyxl',
        'pandas',
        'customtkinter',
        'requests',
        'win32com',
        'win32com.client',
        'pythoncom'
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
    [],
    exclude_binaries=True,
    name='raidstat',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='raidstat_py/gui/assets/icon.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='raidstat',
)
