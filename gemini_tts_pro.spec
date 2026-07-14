# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['Gemini_TTS_Pro_v3_8_13.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['google.genai', 'docx', 'pypdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,   # <-- this is what makes it one-dir instead of one-file
    name='GeminiTTSPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # windowed app (PyQt6 GUI), no terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app_icon.ico',   # uncomment and add an .ico file if you want a custom icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GeminiTTSPro',
)
