# -*- mode: python ; coding: utf-8 -*-

added_files = [
    ('.env', '.'),          # Файл .env в корневую директорию
    ('styles.qss', '.'),     # Файл стилей в корневую директорию
    ('mailboxes.json', '.'),     # JSON-файл в корневую директорию
    ('icons/', 'icons')     # Папка с иконками в директорию icons
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='MailContact',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MailContact',
)
