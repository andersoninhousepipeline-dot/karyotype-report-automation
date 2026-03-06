# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Karyotype Report Generator
# Build:  pyinstaller KaryotypeReport.spec

import os
from pathlib import Path

block_cipher = None

# Collect all font TTF files
fonts_dir = os.path.join(SPECPATH, 'fonts')
font_files = [(os.path.join(fonts_dir, f), 'fonts')
              for f in os.listdir(fonts_dir) if f.endswith('.ttf')]

a = Analysis(
    ['karyotype_report_generator.py'],
    pathex=[SPECPATH],
    binaries=[],
    datas=font_files + [
        ('assets', 'assets'),
    ],
    hiddenimports=[
        'karyotype_template',
        'karyotype_assets',
        'PIL',
        'PIL.Image',
        'PIL.ImageFile',
        'pandas',
        'openpyxl',
        'reportlab',
        'reportlab.pdfbase.pdfmetrics',
        'reportlab.pdfbase.ttfonts',
        'pypdfium2',
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

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='KaryotypeReportGen',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/karyotype_icon.png',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='KaryotypeReportGen',
)
