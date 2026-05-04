# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Application configuration helpers.

All persistent settings are stored in:
    %APPDATA%\\PyVBAai\\config.ini   (Windows)
    ~/.config/PyVBAai/config.ini      (fallback on non-Windows)
"""
from __future__ import annotations

import os
import pathlib

from PySide6.QtCore import QSettings


def config_dir() -> pathlib.Path:
    """Return the directory that holds config.ini, creating it if necessary."""
    appdata = os.environ.get("APPDATA") or str(pathlib.Path.home())
    d = pathlib.Path(appdata) / "PyVBAai"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_settings() -> QSettings:
    """Return a QSettings instance backed by %APPDATA%\\PyVBAai\\config.ini."""
    path = str(config_dir() / "config.ini")
    return QSettings(path, QSettings.Format.IniFormat)
