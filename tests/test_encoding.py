# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Guard against mojibake re-entering the codebase.

Allowed non-ASCII code points (intentional Unicode):
  U+2500  ─  BOX DRAWINGS LIGHT HORIZONTAL  (section-header comments)
  U+2014  —  EM DASH                         (prose comments / docstrings)
  U+2192  →  RIGHTWARDS ARROW               (inline comments)
  U+2248  ≈  ALMOST EQUAL TO               (token-estimate comment)
  U+00D7  ×  MULTIPLICATION SIGN            (row×col dimension strings)
  U+FEFF     BOM                             (UTF-8-BOM file header)

Any other non-ASCII character is presumed to be mojibake and will fail.
"""
from __future__ import annotations

import pathlib

import pytest

# Characters that are explicitly permitted in source files.
_ALLOWED = frozenset([
    0x2500,  # ─  BOX DRAWINGS LIGHT HORIZONTAL
    0x2014,  # —  EM DASH
    0x2192,  # →  RIGHTWARDS ARROW
    0x2248,  # ≈  ALMOST EQUAL TO
    0x00D7,  # ×  MULTIPLICATION SIGN
    0xFEFF,  # BOM (UTF-8-BOM encoding marker)
])

_SKIP_DIRS = {".venv", "__pycache__", "build", "dist"}

ROOT = pathlib.Path(__file__).parent.parent


def _py_files():
    for f in sorted(ROOT.rglob("*.py")):
        if not any(part in _SKIP_DIRS for part in f.parts):
            yield f


@pytest.mark.parametrize("src_file", _py_files(), ids=lambda f: str(f.relative_to(ROOT)))
def test_no_mojibake(src_file: pathlib.Path) -> None:
    """Every .py file must contain only ASCII or explicitly allowed Unicode."""
    text = src_file.read_bytes().decode("utf-8")
    violations: list[str] = []
    for lineno, line in enumerate(text.splitlines(), 1):
        for col, ch in enumerate(line, 1):
            cp = ord(ch)
            if cp > 127 and cp not in _ALLOWED:
                violations.append(
                    f"  line {lineno}, col {col}: U+{cp:04X} {ch!r} in: {line.strip()[:70]}"
                )
    assert not violations, (
        f"{src_file.relative_to(ROOT)} contains disallowed non-ASCII characters:\n"
        + "\n".join(violations[:20])
    )
