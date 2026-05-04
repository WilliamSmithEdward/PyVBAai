# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for PyVBAai.

Build with:
    pyinstaller PyVBAai.spec
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── Collect data files ────────────────────────────────────────────────────────
datas = []
datas += collect_data_files("markdown")
datas += collect_data_files("openai")

# ── Hidden imports ────────────────────────────────────────────────────────────
hidden_imports = [
    # win32com
    "win32com",
    "win32com.client",
    "win32com.client.gencache",
    "pywintypes",
    "win32api",
    "win32con",
    "pythoncom",
    # openai
    "openai",
    "openai.types",
    "openai.types.chat",
    "httpx",
    "anyio",
    "sniffio",
    # markdown
    "markdown",
    "markdown.extensions.fenced_code",
    "markdown.extensions.tables",
    "markdown.extensions.nl2br",
    # PySide6 essentials
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Trim unused Qt modules to reduce size
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets",
        "PySide6.Qt3DCore",
        "PySide6.Qt3DRender",
        "PySide6.QtQuick",
        "PySide6.QtQml",
        "PySide6.QtMultimedia",
        "PySide6.QtLocation",
        "PySide6.QtBluetooth",
        "PySide6.QtSensors",
        "PySide6.QtCharts",
        "PySide6.QtDataVisualization",
        "unittest",
        "pytest",
        "tkinter",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="PyVBAai",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Uncomment and add your .ico file path to set a custom icon:
    # icon="resources/icon.ico",
    version=None,
)
