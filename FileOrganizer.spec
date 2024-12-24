# -*- mode: python ; coding: utf-8 -*-

import os
import sys

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('config.json', '.')],
    hiddenimports=['xml.parsers.expat', 'pkg_resources.py2_warn'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0
)

# Add Python DLL from your conda environment
python_path = os.path.dirname(sys.executable)
a.binaries += TOC([
    (os.path.basename(sys.executable), sys.executable, 'BINARY'),
    ('python312.dll', os.path.join(python_path, 'python312.dll'), 'BINARY'),
])

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FileOrganizer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
