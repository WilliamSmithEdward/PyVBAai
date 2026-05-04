# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for core/excel_writer.py.

Non-VBA handlers use real openpyxl workbooks created in-memory.
VBA-COM handlers use MagicMock to avoid requiring Excel.
"""
from __future__ import annotations

from unittest.mock import MagicMock

import openpyxl
import pytest

from core.excel_writer import (
    ApplyError,
    _add_named_range,
    _add_sheet,
    _clear_range,
    _com_add_vba_module,
    _com_delete_vba_module,
    _com_get_vba_component,
    _com_set_vba,
    _copy_sheet,
    _delete_named_range,
    _delete_sheet,
    _dispatch_openpyxl,
    _expand_flags,
    _get_openpyxl_sheet,
    _merge_cells,
    _move_sheet,
    _rename_sheet,
    _set_cell,
    _set_format,
    _set_range,
    _to_argb,
    _unmerge_cells,
)
from models.conversation import Change

# ── Fixtures ──────────────────────────────────────────────────────────────────

@pytest.fixture()
def owb():
    """A real openpyxl Workbook with two sheets for testing."""
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.create_sheet("Sheet2")
    return wb


@pytest.fixture()
def com_wb():
    """A MagicMock representing an Excel Workbook COM object (VBA tests only)."""
    return MagicMock(name="ComWorkbook")


# ── _get_openpyxl_sheet ───────────────────────────────────────────────────────

class TestGetOpenpyxlSheet:
    def test_returns_sheet(self, owb):
        ws = _get_openpyxl_sheet(owb, "Sheet1")
        assert ws.title == "Sheet1"

    def test_raises_apply_error_on_missing(self, owb):
        with pytest.raises(ApplyError, match="Sheet not found"):
            _get_openpyxl_sheet(owb, "Ghost")


# ── _com_get_vba_component ────────────────────────────────────────────────────

class TestComGetVbaComponent:
    def test_returns_component(self, com_wb):
        fake_comp = MagicMock()
        com_wb.VBProject.VBComponents.return_value = fake_comp
        result = _com_get_vba_component(com_wb, "Module1")
        assert result is fake_comp

    def test_raises_apply_error_on_failure(self, com_wb):
        com_wb.VBProject.VBComponents.side_effect = Exception("access denied")
        with pytest.raises(ApplyError, match="VBA module 'Missing' not found"):
            _com_get_vba_component(com_wb, "Missing")

    def test_error_includes_trust_hint(self, com_wb):
        com_wb.VBProject.VBComponents.side_effect = Exception("x")
        with pytest.raises(ApplyError, match="Trust access"):
            _com_get_vba_component(com_wb, "X")


# ── _set_cell ─────────────────────────────────────────────────────────────────

class TestSetCell:
    def test_sets_value(self, owb):
        _set_cell(owb, {"sheet": "Sheet1", "cell": "A1", "value": 42})
        assert owb["Sheet1"]["A1"].value == 42

    def test_sets_formula(self, owb):
        _set_cell(owb, {"sheet": "Sheet1", "cell": "B2", "formula": "=SUM(A:A)"})
        assert owb["Sheet1"]["B2"].value == "=SUM(A:A)"

    def test_formula_takes_priority(self, owb):
        _set_cell(owb, {"sheet": "Sheet1", "cell": "A1", "formula": "=1+1", "value": 99})
        assert owb["Sheet1"]["A1"].value == "=1+1"

    def test_raises_for_unknown_sheet(self, owb):
        with pytest.raises(ApplyError):
            _set_cell(owb, {"sheet": "Ghost", "cell": "A1", "value": 1})


# ── _set_range ────────────────────────────────────────────────────────────────

class TestSetRange:
    def test_sets_values(self, owb):
        _set_range(owb, {"sheet": "Sheet1", "range": "A1:C2",
                         "values": [[1, 2, 3], [4, 5, 6]]})
        ws = owb["Sheet1"]
        assert ws["A1"].value == 1
        assert ws["B2"].value == 5
        assert ws["C2"].value == 6

    def test_single_row(self, owb):
        _set_range(owb, {"sheet": "Sheet1", "range": "A1:C1",
                         "values": [["a", "b", "c"]]})
        ws = owb["Sheet1"]
        assert ws["A1"].value == "a"
        assert ws["C1"].value == "c"


# ── _clear_range ──────────────────────────────────────────────────────────────

class TestClearRange:
    def test_clears_values(self, owb):
        ws = owb["Sheet1"]
        ws["A1"].value = 100
        ws["B2"].value = "hello"
        _clear_range(owb, {"sheet": "Sheet1", "range": "A1:B2"})
        assert ws["A1"].value is None
        assert ws["B2"].value is None


# ── _add_sheet ────────────────────────────────────────────────────────────────

class TestAddSheet:
    def test_add_new_sheet(self, owb):
        _add_sheet(owb, {"name": "NewSheet"})
        assert "NewSheet" in owb.sheetnames

    def test_add_at_position_1(self, owb):
        _add_sheet(owb, {"name": "First", "position": 1})
        assert owb.sheetnames[0] == "First"

    def test_default_appends(self, owb):
        count_before = len(owb.sheetnames)
        _add_sheet(owb, {"name": "Last"})
        assert owb.sheetnames[-1] == "Last"
        assert len(owb.sheetnames) == count_before + 1


# ── _delete_sheet ─────────────────────────────────────────────────────────────

class TestDeleteSheet:
    def test_deletes_sheet(self, owb):
        _delete_sheet(owb, {"name": "Sheet2"})
        assert "Sheet2" not in owb.sheetnames

    def test_raises_for_missing_sheet(self, owb):
        with pytest.raises(ApplyError):
            _delete_sheet(owb, {"name": "Ghost"})


# ── _rename_sheet ─────────────────────────────────────────────────────────────

class TestRenameSheet:
    def test_renames(self, owb):
        _rename_sheet(owb, {"old_name": "Sheet1", "new_name": "Renamed"})
        assert "Renamed" in owb.sheetnames
        assert "Sheet1" not in owb.sheetnames


# ── _move_sheet ───────────────────────────────────────────────────────────────

class TestMoveSheet:
    def test_move_to_position_1(self, owb):
        # owb starts: [Sheet1, Sheet2]
        _move_sheet(owb, {"name": "Sheet2", "position": 1})
        assert owb.sheetnames[0] == "Sheet2"

    def test_move_to_end(self, owb):
        owb.create_sheet("Sheet3")
        _move_sheet(owb, {"name": "Sheet1", "position": 3})
        assert owb.sheetnames[-1] == "Sheet1"


# ── _copy_sheet ───────────────────────────────────────────────────────────────

class TestCopySheet:
    def test_copies_sheet(self, owb):
        owb["Sheet1"]["A1"].value = "orig"
        _copy_sheet(owb, {"source": "Sheet1", "dest": "CopyOfSheet1", "position": 3})
        assert "CopyOfSheet1" in owb.sheetnames

    def test_raises_for_missing_source(self, owb):
        with pytest.raises(ApplyError):
            _copy_sheet(owb, {"source": "Ghost", "dest": "Copy"})


# ── _merge_cells / _unmerge_cells ─────────────────────────────────────────────

class TestMergeCells:
    def test_merge(self, owb):
        _merge_cells(owb, {"sheet": "Sheet1", "range": "A1:B2"})
        assert "A1:B2" in [str(m) for m in owb["Sheet1"].merged_cells.ranges]

    def test_unmerge(self, owb):
        owb["Sheet1"].merge_cells("C1:D2")
        _unmerge_cells(owb, {"sheet": "Sheet1", "range": "C1:D2"})
        assert "C1:D2" not in [str(m) for m in owb["Sheet1"].merged_cells.ranges]


# ── _add_named_range / _delete_named_range ────────────────────────────────────

class TestNamedRange:
    def test_add_named_range(self, owb):
        _add_named_range(owb, {"name": "MyRange", "refers_to": "Sheet1!$A$1:$D$10"})
        assert "MyRange" in owb.defined_names

    def test_delete_named_range(self, owb):
        from openpyxl.workbook.defined_name import DefinedName
        owb.defined_names["ToDelete"] = DefinedName("ToDelete", attr_text="Sheet1!$A$1")
        _delete_named_range(owb, {"name": "ToDelete"})
        assert "ToDelete" not in owb.defined_names

    def test_delete_missing_raises(self, owb):
        with pytest.raises(ApplyError, match="Named range 'Ghost' not found"):
            _delete_named_range(owb, {"name": "Ghost"})


# ── _com_set_vba ──────────────────────────────────────────────────────────────

class TestComSetVba:
    def test_replaces_existing_code(self, com_wb):
        comp = MagicMock()
        code_module = MagicMock()
        code_module.CountOfLines = 5
        comp.CodeModule = code_module
        com_wb.VBProject.VBComponents.return_value = comp

        _com_set_vba(com_wb, {"module": "Module1", "code": "Sub New()\nEnd Sub"})
        code_module.DeleteLines.assert_called_once_with(1, 5)
        code_module.InsertLines.assert_called_once_with(1, "Sub New()\nEnd Sub")

    def test_empty_module_no_delete(self, com_wb):
        comp = MagicMock()
        code_module = MagicMock()
        code_module.CountOfLines = 0
        comp.CodeModule = code_module
        com_wb.VBProject.VBComponents.return_value = comp

        _com_set_vba(com_wb, {"module": "Module1", "code": "Sub X()\nEnd Sub"})
        code_module.DeleteLines.assert_not_called()
        code_module.InsertLines.assert_called_once()

    def test_empty_code_no_insert(self, com_wb):
        comp = MagicMock()
        code_module = MagicMock()
        code_module.CountOfLines = 3
        comp.CodeModule = code_module
        com_wb.VBProject.VBComponents.return_value = comp

        _com_set_vba(com_wb, {"module": "Module1", "code": ""})
        code_module.DeleteLines.assert_called_once()
        code_module.InsertLines.assert_not_called()


# ── _com_add_vba_module ───────────────────────────────────────────────────────

class TestComAddVbaModule:
    def test_adds_and_names_module(self, com_wb):
        new_comp = MagicMock()
        com_wb.VBProject.VBComponents.Add.return_value = new_comp
        _com_add_vba_module(com_wb, {"name": "NewMod", "code": "Sub Test()\nEnd Sub"})
        com_wb.VBProject.VBComponents.Add.assert_called_once_with(1)
        assert new_comp.Name == "NewMod"
        new_comp.CodeModule.InsertLines.assert_called_once_with(1, "Sub Test()\nEnd Sub")

    def test_no_insert_when_no_code(self, com_wb):
        new_comp = MagicMock()
        com_wb.VBProject.VBComponents.Add.return_value = new_comp
        _com_add_vba_module(com_wb, {"name": "Empty", "code": ""})
        new_comp.CodeModule.InsertLines.assert_not_called()

    def test_raises_apply_error_on_com_failure(self, com_wb):
        com_wb.VBProject.VBComponents.Add.side_effect = Exception("Trust Center blocked")
        with pytest.raises(ApplyError, match="Cannot add VBA module"):
            _com_add_vba_module(com_wb, {"name": "X", "code": ""})


# ── _com_delete_vba_module ────────────────────────────────────────────────────

class TestComDeleteVbaModule:
    def test_removes_component(self, com_wb):
        comp = MagicMock()
        com_wb.VBProject.VBComponents.return_value = comp
        _com_delete_vba_module(com_wb, {"name": "Module1"})
        com_wb.VBProject.VBComponents.Remove.assert_called_once_with(comp)

    def test_raises_when_not_found(self, com_wb):
        com_wb.VBProject.VBComponents.side_effect = Exception("not found")
        with pytest.raises(ApplyError):
            _com_delete_vba_module(com_wb, {"name": "Ghost"})


# ── _dispatch_openpyxl ────────────────────────────────────────────────────────

class TestDispatchOpenpyxl:
    def test_unknown_type_raises(self, owb):
        change = Change(type="do_magic", params={})
        with pytest.raises(ApplyError, match="Unknown change type: 'do_magic'"):
            _dispatch_openpyxl(owb, change)

    @pytest.mark.parametrize("op_type,params", [
        ("set_cell",          {"sheet": "Sheet1", "cell": "A1", "value": 1}),
        ("set_range",         {"sheet": "Sheet1", "range": "A1:B2", "values": [[1, 2], [3, 4]]}),
        ("clear_range",       {"sheet": "Sheet1", "range": "A1:Z100"}),
        ("delete_sheet",      {"name": "Sheet2"}),
        ("rename_sheet",      {"old_name": "Sheet1", "new_name": "Renamed"}),
    ])
    def test_known_types_dispatched(self, owb, op_type, params):
        """Known types should dispatch (may mutate owb) without 'Unknown change type'."""
        change = Change(type=op_type, params=params)
        try:
            _dispatch_openpyxl(owb, change)
        except ApplyError as exc:
            assert "Unknown change type" not in str(exc)


class TestToArgb:
    def test_6_char_prefixed(self):
        assert _to_argb("FF0000") == "FFFF0000"

    def test_8_char_unchanged(self):
        assert _to_argb("FFFF0000") == "FFFF0000"

    def test_lowercase_normalised(self):
        assert _to_argb("ff0000") == "FFFF0000"

    def test_hash_prefix_stripped(self):
        assert _to_argb("#92D050") == "FF92D050"

    def test_invalid_raises(self):
        with pytest.raises(ValueError):
            _to_argb("XYZ")

    def test_too_short_raises(self):
        with pytest.raises(ValueError):
            _to_argb("F00")


class TestExpandFlagsColors:
    def test_tilde_flag_normalised_to_argb(self):
        result = _expand_flags({"flags": "~92D050"})
        assert result["bg_color"] == "FF92D050"

    def test_caret_flag_normalised_to_argb(self):
        result = _expand_flags({"flags": "^FF0000"})
        assert result["font_color"] == "FFFF0000"

    def test_explicit_6char_color_normalised(self):
        result = _expand_flags({"bg_color": "D9E1F2"})
        assert result["bg_color"] == "FFD9E1F2"

    def test_explicit_8char_color_unchanged(self):
        result = _expand_flags({"font_color": "FF123456"})
        assert result["font_color"] == "FF123456"


class TestSetFormatColors:
    def test_bg_color_applied(self, owb):
        ws = owb.active
        _set_format(owb, {"sheet": ws.title, "range": "A1", "bg_color": "92D050"})
        assert ws["A1"].fill.fgColor.rgb[-6:] == "92D050"

    def test_font_color_applied(self, owb):
        ws = owb.active
        _set_format(owb, {"sheet": ws.title, "range": "A1", "font_color": "FF0000"})
        assert ws["A1"].font.color.rgb[-6:] == "FF0000"

    def test_set_cell_with_color_flag(self, owb):
        ws = owb.active
        _set_cell(owb, {"sheet": ws.title, "cell": "B1", "value": "x", "flags": "~92D050"})
        assert ws["B1"].fill.fgColor.rgb[-6:] == "92D050"
