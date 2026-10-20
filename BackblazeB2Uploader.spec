# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:/Automatisation/Sauvegarde Backblaze B2 CM/app.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/PMD/AppData/Local/Programs/Python/Python312/Lib/site-packages/customtkinter', 'customtkinter'), ('D:/Automatisation/Sauvegarde Backblaze B2 CM/assets', 'assets')],
    hiddenimports=['openpyxl', 'pandas', 'b2sdk', 'PIL.ImageTk'],
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
    name='BackblazeB2Uploader',
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
    icon=['D:/Automatisation/Sauvegarde Backblaze B2 CM/assets/icone_application_b2.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BackblazeB2Uploader',
)
