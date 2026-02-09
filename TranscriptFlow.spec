# -*- mode: python ; coding: utf-8 -*-
"""
TranscriptFlow PyInstaller Specification File

This spec file configures PyInstaller to bundle TranscriptFlow with all dependencies,
including PyEnchant spell checking libraries and data files.

Build command: pyinstaller TranscriptFlow.spec
"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

block_cipher = None

# Get the application directory
app_dir = os.path.abspath(SPECPATH)

# Collect PyEnchant data and binaries
enchant_datas = collect_data_files('enchant')
enchant_binaries = collect_dynamic_libs('enchant')

# Collect all hidden imports
hiddenimports = [
    'enchant',
    'enchant.checker',
    'enchant.tokenize',
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.QtWidgets',
    'PyQt6.QtPrintSupport',
    'vlc',
    'ffmpeg',
    'numpy',
    'pynput',
    'pynput.keyboard',
    'pynput.mouse',
]

# Add all enchant submodules
hiddenimports.extend(collect_submodules('enchant'))

# Data files to include
datas = [
    # Application icons
    ('app_icon.ico', '.'),
    ('app_icon.png', '.'),
    
    # MPV library for media playback
    ('mpv-1.dll', '.'),
    
    # Dictionary files (bundled with app)
    ('dicts', 'dicts'),
    
    # Custom dictionaries
    ('custom_dicts', 'custom_dicts'),
    
    # Default configuration
    ('config.json', '.'),
    
    # Documentation
    ('HELP.html', '.'),
    ('splash.png', '.'),
    ('Jameel Noori Nastaleeq.ttf', '.'),
]

# Add enchant data files
datas.extend(enchant_datas)

# Binaries to include
binaries = []
binaries.extend(enchant_binaries)

a = Analysis(
    ['main.py'],
    pathex=[app_dir],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'pandas',
        'PIL',
        'tkinter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TranscriptFlow',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to True for debugging, False for release
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
    version='version_info.txt',  # Add version information to executable
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TranscriptFlow',
)
