# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Workbook data models."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CellFormat:
    """Formatting attributes for a single cell."""
    number_format: str = ""          # e.g. "0.00", "dd/mm/yyyy", "$#,##0"
    bold: bool = False
    italic: bool = False
    strikethrough: bool = False
    underline: bool = False
    font_name: str = ""              # e.g. "Calibri"
    font_size: float = 0.0
    font_color: str = ""             # hex RGB, e.g. "FF0000"
    bg_color: str = ""               # hex RGB fill colour
    h_align: str = ""                # "left", "center", "right", "fill", "justify"
    v_align: str = ""                # "top", "center", "bottom"
    wrap_text: bool = False
    locked: bool = True
    # Border sides: "<style>" or "<style>:<RRGGBB>"  (thin|medium|thick|dashed|dotted|double|hair)
    border_top: str = ""
    border_bottom: str = ""
    border_left: str = ""
    border_right: str = ""


@dataclass
class CellData:
    row: int
    col: int
    address: str
    value: Any
    formula: str = ""
    fmt: CellFormat | None = None    # None means default/unset formatting


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


# All format field keys that can be included or excluded from context output.
ALL_FMT_FIELDS: frozenset[str] = frozenset({
    "number_format",
    "bold", "italic", "strikethrough", "underline",
    "font_name", "font_size", "font_color",
    "bg_color",
    "h_align", "v_align", "wrap_text",
    "border_top", "border_bottom", "border_left", "border_right",
})


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
    # Limit rows shown per contiguous row-run (None = unlimited)
    max_rows_per_area: int | None = None
    # Which formatting fields to emit in cell notation (default = all)
    fmt_include: set[str] = field(default_factory=lambda: set(ALL_FMT_FIELDS))
