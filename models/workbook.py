# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Workbook data models."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CellData:
    row: int
    col: int
    address: str
    value: Any
    formula: str = ""


@dataclass
class SheetData:
    index: int
    name: str
    used_range_address: str
    row_count: int
    col_count: int
    cells: dict[str, CellData] = field(default_factory=dict)
    visible: bool = True


@dataclass
class VBAModule:
    name: str
    module_type: int   # 1=Standard, 2=Class, 3=Form, 100=Document
    type_name: str
    code: str


@dataclass
class NamedRange:
    name: str
    refers_to: str
    scope: str = "Workbook"


@dataclass
class WorkbookData:
    file_path: str
    name: str
    sheets: list[SheetData] = field(default_factory=list)
    vba_modules: list[VBAModule] = field(default_factory=list)
    named_ranges: list[NamedRange] = field(default_factory=list)
    has_vba: bool = False
    extraction_error: str | None = None


@dataclass
class ContextConfig:
    """What to include when building the GPT context."""
    include_formulas: bool = True
    include_vba: bool = True
    include_named_ranges: bool = True
    included_sheets: list[str] | None = None       # None = all sheets
    included_vba_modules: list[str] | None = None  # None = all modules
    excluded_sheets: list[str] = field(default_factory=list)
    excluded_vba_modules: list[str] = field(default_factory=list)
    excluded_areas: dict[str, list[str]] = field(default_factory=dict)
