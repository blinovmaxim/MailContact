# -*- mode: python ; coding: utf-8 -*-

added_files = [
    ('.env', '.'),          # .env файл в корневую директорию
    ('styles.qss', '.'),     # Файл стилей в корневую директорию
    ('mailboxes.json', '.'),     # JSON-файл в корневую директорию
    ('icons/', 'icons')     # Папка с иконками в директорию icons
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'requests',
        'python-dotenv',
        'PyQt5',
        'csv',
        'json',
        'os',
        'dotenv'
    ],
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
    exclude_binaries=False,
    name='MailContact',
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
)
