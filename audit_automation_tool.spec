# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_submodules


project_dir = Path(__file__).resolve().parent

# Collect Streamlit's package data (frontend assets) and submodules
streamlit_all = collect_all("streamlit")

# Collect common scientific stack metadata that can be missed
pandas_submodules = collect_submodules("pandas")
openpyxl_submodules = collect_submodules("openpyxl")
ant_submodules = collect_submodules("anthropic")

a = Analysis(
    [str(project_dir / "main.py")],
    pathex=[str(project_dir)],
    binaries=streamlit_all[1],
    datas=streamlit_all[0] + [(str(project_dir / "ui" / "app.py"), "ui")],
    hiddenimports=streamlit_all[2] + pandas_submodules + openpyxl_submodules + ant_submodules,
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
    a.binaries,
    a.datas,
    [],
    name="AuditAutomationTool",
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
