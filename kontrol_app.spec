# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['kontrol.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Отчёт.xlsx', '.'),
        ('Сводный.xlsx', '.'),
        ('plavka.xlsx', '.'),
        ('control.xlsx', '.'),
    ],
    hiddenimports=['win32api', 'win32print', 'win32com.client', 'openpyxl', 'PySide6'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'tkinter', 'matplotlib', 'scipy', 'torch', 'sympy', 'IPython'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Электронный журнал контроля. Версия 3.1.1',
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