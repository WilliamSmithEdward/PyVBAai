# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Workbook data models."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CellFormat:
    """Non-default formatting for a cell.  Only populated fields are serialised."""
    number_format: str | None = None   # e.g. "#,##0.00", "dd/mm/yyyy"
    bold: bool | None = None
    italic: bool | None = None
    font_color: str | None = None      # 6-digit hex, e.g. "FF0000"
    bg_color: str | None = None        # 6-digit hex; None when no fill
    h_align: str | None = None         # "left"|"center"|"right"|"general"
    wrap_text: bool | None = None


@dataclass
class CellData:
    row: int
    col: int
    address: str
    value: Any
    formula: str = ""
    fmt: CellFormat | None = None      # None when format reading is disabled


@dataclass
class SheetData:
    index: int
    name: str
    used_range_address: str
    row_count: int
    col_count: int
    cells: dict[str, CellData] = field(default_factory=dict)
    visible: bool = True
    area_addresses: list[str] = field(default_factory=list)


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
    loaded_mtime: float = 0.0  # os.path.getmtime() at the time the file was read


@dataclass
class ContextConfig:
    """What to include when building the GPT context."""
    include_formulas: bool = True
    include_vba: bool = True
    include_named_ranges: bool = True
    # Format aspects
    include_number_format: bool = False
    include_font_style: bool = False    # bold + italic
    include_font_color: bool = False
    include_bg_color: bool = False
    include_alignment: bool = False
    # Row limit per contiguous area (0 = no limit)
    limit_rows_per_area: bool = False
    max_rows_per_area: int = 100
    included_sheets: list[str] | None = None       # None = all sheets
    included_vba_modules: list[str] | None = None  # None = all modules
    excluded_sheets: list[str] = field(default_factory=list)
    excluded_vba_modules: list[str] = field(default_factory=list)
    excluded_areas: dict[str, list[str]] = field(default_factory=dict)

    @property
    def include_any_format(self) -> bool:
        return any([
            self.include_number_format, self.include_font_style,
            self.include_font_color, self.include_bg_color, self.include_alignment,
        ])
