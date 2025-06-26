# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

# Get the current directory (where the spec file is located)
current_dir = os.path.dirname(os.path.abspath(SPEC))

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[current_dir],
    binaries=[],
    datas=[
        # Include template files
        ('templates', 'templates'),
        # Include static files (company logo)
        ('static', 'static'),
        # Include report generator
        ('report_generator.py', '.'),
        ('app.py', '.'),
        # Include sample database file if it exists
        ('logo_database/ArtDBSample.xlsx', 'logo_database') if os.path.exists('logo_database/ArtDBSample.xlsx') else None,
        # Include sample logo images if they exist
        ('logo_images', 'logo_images') if os.path.exists('logo_images') else None,
    ],
    hiddenimports=[
        'flask',
        'pandas',
        'fpdf',
        'openpyxl',
        'PIL',
        'werkzeug',
        'xlrd',
        'report_generator'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Filter out None values from datas
a.datas = [x for x in a.datas if x is not None]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Art_Instructions_Generator',
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Art_Instructions_Generator',
)