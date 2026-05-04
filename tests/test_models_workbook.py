# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for models/workbook.py dataclasses."""
from __future__ import annotations

import pytest

from models.workbook import (
    CellData,
    ContextConfig,
    NamedRange,
    SheetData,
    VBAModule,
    WorkbookData,
)

# ── CellData ──────────────────────────────────────────────────────────────────

class TestCellData:
    def test_basic_fields(self):
        c = CellData(row=3, col=5, address="E3", value="test")
        assert c.row == 3
        assert c.col == 5
        assert c.address == "E3"
        assert c.value == "test"
        assert c.formula == ""

    def test_formula_field(self):
        c = CellData(row=1, col=1, address="A1", value=0, formula="=SUM(B1:B10)")
        assert c.formula == "=SUM(B1:B10)"

    def test_numeric_value(self):
        c = CellData(row=1, col=2, address="B1", value=3.14)
        assert c.value == pytest.approx(3.14)

    def test_none_value(self):
        c = CellData(row=1, col=1, address="A1", value=None)
        assert c.value is None


# ── SheetData ─────────────────────────────────────────────────────────────────

class TestSheetData:
    def test_defaults(self):
        s = SheetData(index=1, name="Sheet1", used_range_address="A1:B2",
                      row_count=2, col_count=2)
        assert s.visible is True
        assert s.cells == {}

    def test_cell_assignment(self, simple_cell):
        s = SheetData(index=1, name="Sheet1", used_range_address="A1",
                      row_count=1, col_count=1)
        s.cells["A1"] = simple_cell
        assert "A1" in s.cells
        assert s.cells["A1"].value == "Hello"

    def test_hidden_sheet(self):
        s = SheetData(index=2, name="Hidden", used_range_address="",
                      row_count=0, col_count=0, visible=False)
        assert s.visible is False


# ── VBAModule ─────────────────────────────────────────────────────────────────

class TestVBAModule:
    def test_fields(self, vba_module):
        assert vba_module.name == "Module1"
        assert vba_module.module_type == 1
        assert vba_module.type_name == "Standard"
        assert "Sub Hello" in vba_module.code

    def test_empty_code(self):
        m = VBAModule(name="Empty", module_type=1, type_name="Standard", code="")
        assert m.code == ""

    def test_class_module_type(self):
        m = VBAModule(name="MyClass", module_type=2, type_name="Class", code="")
        assert m.module_type == 2


# ── NamedRange ────────────────────────────────────────────────────────────────

class TestNamedRange:
    def test_fields(self, named_range):
        assert named_range.name == "MyRange"
        assert named_range.refers_to == "=Sheet1!$A$1:$D$10"
        assert named_range.scope == "Workbook"

    def test_default_scope(self):
        nr = NamedRange(name="X", refers_to="=Sheet1!$A$1")
        assert nr.scope == "Workbook"

    def test_sheet_scoped(self):
        nr = NamedRange(name="Local", refers_to="=Sheet1!$A$1", scope="Sheet1")
        assert nr.scope == "Sheet1"


# ── WorkbookData ──────────────────────────────────────────────────────────────

class TestWorkbookData:
    def test_basic_fields(self, workbook):
        assert workbook.file_path == "C:/test/book.xlsx"
        assert workbook.name == "book.xlsx"
        assert workbook.has_vba is True
        assert workbook.extraction_error is None

    def test_default_collections(self):
        wb = WorkbookData(file_path="f.xlsx", name="f.xlsx")
        assert wb.sheets == []
        assert wb.vba_modules == []
        assert wb.named_ranges == []
        assert wb.has_vba is False

    def test_extraction_error(self):
        wb = WorkbookData(file_path="f.xlsx", name="f.xlsx",
                          extraction_error="VBA access denied")
        assert wb.extraction_error == "VBA access denied"

    def test_sheet_count(self, workbook):
        assert len(workbook.sheets) == 1

    def test_vba_module_count(self, workbook):
        assert len(workbook.vba_modules) == 1

    def test_named_range_count(self, workbook):
        assert len(workbook.named_ranges) == 1


# ── ContextConfig ─────────────────────────────────────────────────────────────

class TestContextConfig:
    def test_defaults(self):
        c = ContextConfig()
        assert c.include_formulas is True
        assert c.include_vba is True
        assert c.include_named_ranges is True
        assert c.included_sheets is None
        assert c.included_vba_modules is None
        assert c.excluded_areas == {}

    def test_custom_values(self):
        c = ContextConfig(
            include_formulas=False,
            include_vba=False,
            included_sheets=["Sheet1", "Sheet2"],
            included_vba_modules=["Module1"],
            excluded_areas={"Sheet1": ["$A$1:$C$10"]},
        )
        assert c.include_formulas is False
        assert c.included_sheets == ["Sheet1", "Sheet2"]
        assert c.included_vba_modules == ["Module1"]
        assert c.excluded_areas == {"Sheet1": ["$A$1:$C$10"]}
