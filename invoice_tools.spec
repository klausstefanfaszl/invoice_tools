# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

_exchangelib_imports = collect_submodules('exchangelib')
#
# PyInstaller-Spec für invoice_tools
# Kombiniert invoice_extractor und inbox_processor in eine exe
# mit gemeinsamem _internal-Verzeichnis.
#
# Build:
#   pyinstaller invoice_tools.spec
#
# Ausgabe: dist/invoice_tools/invoice_tools.exe  (Windows)
#          dist/invoice_tools/invoice_tools       (Linux)
#
# Verwendung:
#   invoice_tools.exe extractor rechnung.pdf
#   invoice_tools.exe inbox -m unread

a = Analysis(
    ['invoice_tools.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'invoice_extractor',
        'inbox_processor',
        'imaplib',
        'email',
        'email.header',
        'email.mime',
        'email.mime.multipart',
        'email.mime.base',
        'urllib3',
        'google.genai',
        'google.genai.types',
    ] + _exchangelib_imports,
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
    name='invoice_tools',
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
    name='invoice_tools',
)
