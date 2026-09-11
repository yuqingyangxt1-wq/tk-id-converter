# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the TK Indonesia table converter (macOS local build)."""

from pathlib import Path
import sys

block_cipher = None
APP_NAME = "TK-ID-Converter"
PROJECT_DIR = Path(SPECPATH).resolve()

sys.path.insert(0, str(PROJECT_DIR))

a = Analysis(
    [str(PROJECT_DIR / "app" / "__main__.py")],
    pathex=[str(PROJECT_DIR)],
    binaries=[],
    datas=[(str(PROJECT_DIR / "assets"), "assets")],
    hiddenimports=[
        "app",
        "app.config",
        "app.source_reader",
        "app.converter",
        "app.product_pool",
        "app.tiktok_writer",
        "tkinterdnd2",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "scipy", "pandas", "numpy.testing",
        "pytest", "setuptools",
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
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
    icon=None,
)
