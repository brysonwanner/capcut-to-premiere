# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — works on Windows and Mac
# Build: pyinstaller capcut_premiere.spec

import sys

block_cipher = None

a = Analysis(
    ['capcut_premiere_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CapCut XML Export Tool',
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon=None,
)

if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='CapCut to Premiere.app',
        icon=None,  # replace with 'icon.icns' when ready
        bundle_identifier='com.capcut.premiere.converter',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0',
        },
    )
