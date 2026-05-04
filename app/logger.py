# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Debug logging for PyVBAai.

When DEBUG=true in .env.local (project root), a rotating log file is written to
debug/pyvbaai.log.  All other code calls get_logger() to obtain the logger; if
debug mode is off the logger has a NullHandler and is effectively silent.
"""
from __future__ import annotations

import logging
import os
import pathlib
from logging.handlers import RotatingFileHandler

_LOG_NAME = "pyvbaai"
_ENV_FILE = pathlib.Path(__file__).parent.parent / ".env.local"
_DEBUG_DIR = pathlib.Path(__file__).parent.parent / "debug"
_LOG_FILE = _DEBUG_DIR / "pyvbaai.log"

_initialised = False


def _is_debug() -> bool:
    """Return True if DEBUG=true in .env.local or in the process environment."""
    if os.environ.get("PYVBAAI_DEBUG", "").lower() == "true":
        return True
    try:
        text = _ENV_FILE.read_text(encoding="utf-8")
        for line in text.splitlines():
            line = line.strip()
            if line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            if key.strip().upper() == "DEBUG" and val.strip().lower() == "true":
                return True
    except OSError:
        pass
    return False


def init_logging() -> None:
    """Call once at app startup.  Sets up file handler when debug is enabled."""
    global _initialised  # noqa: PLW0603
    if _initialised:
        return
    _initialised = True

    logger = logging.getLogger(_LOG_NAME)
    logger.setLevel(logging.DEBUG)

    if not _is_debug():
        logger.addHandler(logging.NullHandler())
        return

    _DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s:%(module)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = RotatingFileHandler(
        _LOG_FILE,
        maxBytes=5 * 1024 * 1024,   # 5 MB
        backupCount=3,
        encoding="utf-8",
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info("=== PyVBAai debug logging started (log: %s) ===", _LOG_FILE)


def get_logger(name: str = "") -> logging.Logger:
    """Return a child logger under 'pyvbaai'.  Pass __name__ from the caller."""
    child = f"{_LOG_NAME}.{name}" if name else _LOG_NAME
    return logging.getLogger(child)
