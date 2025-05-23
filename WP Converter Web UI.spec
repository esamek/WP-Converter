# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/Users/evansamek/Code/WP Converter/web_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('ui', 'ui')],
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
    name='WP Converter Web UI',
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
    icon=['/Users/evansamek/Code/WP Converter/icon.icns'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WP Converter Web UI',
)
app = BUNDLE(
    coll,
    name='WP Converter Web UI.app',
    icon='/Users/evansamek/Code/WP Converter/icon.icns',
    bundle_identifier=None,
)
