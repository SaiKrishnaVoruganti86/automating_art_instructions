# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Define data files to include
added_files = [
    ('templates', 'templates'),
    ('static', 'static'),
    ('report_generator.py', '.'),
]

# Hidden imports - modules that PyInstaller might miss
hidden_imports = [
    'flask',
    'pandas',
    'fpdf2',
    'openpyxl', 
    'PIL',
    'PIL.Image',
    'werkzeug',
    'werkzeug.utils',
    'xlrd',
    'uuid',
    'threading',
    'time',
    'collections',
    'datetime',
    'os',
    'zipfile',
    'pathlib',
    'webbrowser'
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Art_Instructions_Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to False to hide console window (but useful for debugging)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)