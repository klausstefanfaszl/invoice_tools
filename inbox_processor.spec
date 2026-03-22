# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller-Spec für inbox_processor
#
# Build:
#   pyinstaller inbox_processor.spec
#
# Ausgabe: dist/inbox_processor/inbox_processor  (Linux)
#          dist/inbox_processor/inbox_processor.exe  (Windows)
#
# Hinweis: Die Konfigurationsdateien (invoice_inbox_config.xml,
#          invoice_extractor_config_RE.xml) werden NICHT eingebettet —
#          sie müssen neben der exe liegen und sind dort editierbar.

a = Analysis(
    ['inbox_processor.py'],
    pathex=['.'],           # invoice_extractor.py liegt im selben Verzeichnis
    binaries=[],
    datas=[],
    hiddenimports=[
        'invoice_extractor',    # wird per sys.path + import geladen
        'exchangelib',
        'exchangelib.autodiscover',
        'exchangelib.protocol',
        'exchangelib.transport',
        'imaplib',
        'email',
        'email.header',
        'email.mime',
        'email.mime.multipart',
        'email.mime.base',
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
    [],
    exclude_binaries=True,
    name='inbox_processor',
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
    name='inbox_processor',
)
