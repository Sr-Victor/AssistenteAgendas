# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['MAIN PATH\\Monitoring.py'],
    pathex=[],
    binaries=[],
    datas=[('FEATURES/Mon Features/credentials.json', 'FEATURES/Mon Features'), ('FEATURES/Mon Features/token.json', 'FEATURES/Mon Features'), ('FEATURES/CalendarF/CALENDAR.json', 'FEATURES/CalendarF'), ('FEATURES/CalendarF/token.json', 'FEATURES/CalendarF')],
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
    a.binaries,
    a.datas,
    [],
    name='Monitoring',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
