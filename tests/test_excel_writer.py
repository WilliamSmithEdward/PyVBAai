# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for core/excel_writer.py — COM calls mocked with MagicMock."""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from core.excel_writer import (
    ApplyError,
    _add_named_range,
    _add_sheet,
    _add_vba_module,
    _clear_range,
    _copy_sheet,
    _delete_named_range,
    _delete_sheet,
    _delete_vba_module,
    _dispatch,
    _get_sheet,
    _get_vba_component,
    _move_sheet,
    _rename_sheet,
    _set_cell,
    _set_range,
    _set_vba,
)
from models.conversation import Change

# ── Fixtures ──────────────────────────────────────────────────────────────────

@pytest.fixture()
def wb():
    """A MagicMock representing an Excel Workbook COM object."""
    return MagicMock(name="Workbook")


# ── _get_sheet ────────────────────────────────────────────────────────────────

class TestGetSheet:
    def test_returns_sheet(self, wb):
        fake_sheet = MagicMock()
        wb.Worksheets.return_value = fake_sheet
        result = _get_sheet(wb, "Sheet1")
        wb.Worksheets.assert_called_once_with("Sheet1")
        assert result is fake_sheet

    def test_raises_apply_error_on_failure(self, wb):
        wb.Worksheets.side_effect = Exception("not found")
        with pytest.raises(ApplyError, match="Sheet not found"):
            _get_sheet(wb, "Missing")


# ── _get_vba_component ────────────────────────────────────────────────────────

class TestGetVbaComponent:
    def test_returns_component(self, wb):
        fake_comp = MagicMock()
        wb.VBProject.VBComponents.return_value = fake_comp
        result = _get_vba_component(wb, "Module1")
        assert result is fake_comp

    def test_raises_apply_error_on_failure(self, wb):
        wb.VBProject.VBComponents.side_effect = Exception("access denied")
        with pytest.raises(ApplyError, match="VBA module 'Missing' not found"):
            _get_vba_component(wb, "Missing")

    def test_error_includes_trust_hint(self, wb):
        wb.VBProject.VBComponents.side_effect = Exception("x")
        with pytest.raises(ApplyError, match="Trust access"):
            _get_vba_component(wb, "X")


# ── _set_cell ─────────────────────────────────────────────────────────────────

class TestSetCell:
    def test_sets_value(self, wb):
        cell = MagicMock()
        wb.Worksheets.return_value.Range.return_value = cell
        _set_cell(wb, {"sheet": "Sheet1", "cell": "A1", "value": 42})
        assert cell.Value == 42

    def test_sets_formula(self, wb):
        cell = MagicMock()
        wb.Worksheets.return_value.Range.return_value = cell
        _set_cell(wb, {"sheet": "Sheet1", "cell": "B2", "formula": "=SUM(A:A)"})
        assert cell.Formula == "=SUM(A:A)"

    def test_formula_takes_priority_over_value(self, wb):
        cell = MagicMock()
        wb.Worksheets.return_value.Range.return_value = cell
        _set_cell(wb, {"sheet": "Sheet1", "cell": "A1", "formula": "=1+1", "value": 99})
        assert cell.Formula == "=1+1"
        # Value should not be set when formula is present
        assert not hasattr(cell, '_value_set') or cell.Value != 99

    def test_raises_for_unknown_sheet(self, wb):
        wb.Worksheets.side_effect = Exception("no such sheet")
        with pytest.raises(ApplyError):
            _set_cell(wb, {"sheet": "Ghost", "cell": "A1", "value": 1})


# ── _set_range ────────────────────────────────────────────────────────────────

class TestSetRange:
    def test_sets_values_as_tuple_of_tuples(self, wb):
        rng = MagicMock()
        wb.Worksheets.return_value.Range.return_value = rng
        values = [[1, 2, 3], [4, 5, 6]]
        _set_range(wb, {"sheet": "Sheet1", "range": "A1:C2", "values": values})
        assert rng.Value == ((1, 2, 3), (4, 5, 6))

    def test_single_row(self, wb):
        rng = MagicMock()
        wb.Worksheets.return_value.Range.return_value = rng
        _set_range(wb, {"sheet": "Sheet1", "range": "A1:C1", "values": [["a", "b", "c"]]})
        assert rng.Value == (("a", "b", "c"),)


# ── _clear_range ──────────────────────────────────────────────────────────────

class TestClearRange:
    def test_calls_clear_contents(self, wb):
        rng = MagicMock()
        wb.Worksheets.return_value.Range.return_value = rng
        _clear_range(wb, {"sheet": "Sheet1", "range": "A1:Z100"})
        rng.ClearContents.assert_called_once()


# ── _add_sheet ────────────────────────────────────────────────────────────────

class TestAddSheet:
    def test_add_at_position_1(self, wb):
        wb.Worksheets.Count = 2
        before_sheet = MagicMock()
        new_sheet = MagicMock()
        wb.Worksheets.Add.return_value = new_sheet
        wb.Worksheets.side_effect = lambda x: before_sheet if x == 1 else MagicMock()
        _add_sheet(wb, {"name": "NewSheet", "position": 1})
        wb.Worksheets.Add.assert_called_once_with(Before=before_sheet)
        assert new_sheet.Name == "NewSheet"

    def test_add_at_end(self, wb):
        wb.Worksheets.Count = 2
        last_sheet = MagicMock()
        wb.Worksheets.side_effect = lambda x: last_sheet
        _add_sheet(wb, {"name": "Last", "position": 99})
        wb.Worksheets.Add.assert_called_once_with(After=last_sheet)

    def test_default_position_appends(self, wb):
        wb.Worksheets.Count = 1
        last_sheet = MagicMock()
        wb.Worksheets.side_effect = lambda x: last_sheet
        _add_sheet(wb, {"name": "Appended"})
        # position defaults to Count+1 = 2, which is > Count=1 → After
        wb.Worksheets.Add.assert_called_once_with(After=last_sheet)


# ── _delete_sheet ─────────────────────────────────────────────────────────────

class TestDeleteSheet:
    def test_calls_delete(self, wb):
        sheet = MagicMock()
        wb.Worksheets.return_value = sheet
        _delete_sheet(wb, {"name": "Sheet1"})
        sheet.Delete.assert_called_once()

    def test_raises_for_missing_sheet(self, wb):
        wb.Worksheets.side_effect = Exception("not found")
        with pytest.raises(ApplyError):
            _delete_sheet(wb, {"name": "Ghost"})


# ── _rename_sheet ─────────────────────────────────────────────────────────────

class TestRenameSheet:
    def test_sets_name(self, wb):
        sheet = MagicMock()
        wb.Worksheets.return_value = sheet
        _rename_sheet(wb, {"old_name": "OldName", "new_name": "NewName"})
        assert sheet.Name == "NewName"


# ── _move_sheet ───────────────────────────────────────────────────────────────

class TestMoveSheet:
    def test_move_before(self, wb):
        sheet = MagicMock()
        target = MagicMock()
        # First call returns the sheet to move; subsequent calls return target sheets
        wb.Worksheets.Count = 3
        call_count = [0]
        def side(x):
            call_count[0] += 1
            if call_count[0] == 1:
                return sheet
            return target
        wb.Worksheets.side_effect = side
        _move_sheet(wb, {"name": "Sheet1", "position": 2})
        sheet.Move.assert_called_once_with(Before=target)

    def test_clamps_position_to_1(self, wb):
        sheet = MagicMock()
        target = MagicMock()
        wb.Worksheets.Count = 3
        call_count = [0]
        def side(x):
            call_count[0] += 1
            return sheet if call_count[0] == 1 else target
        wb.Worksheets.side_effect = side
        _move_sheet(wb, {"name": "Sheet1", "position": -99})
        sheet.Move.assert_called_once()


# ── _copy_sheet ───────────────────────────────────────────────────────────────

class TestCopySheet:
    def test_copies_and_renames(self, wb):
        source = MagicMock()
        wb.Worksheets.Count = 2
        call_count = [0]
        def side(x):
            call_count[0] += 1
            return source
        wb.Worksheets.side_effect = side
        _copy_sheet(wb, {"source": "Sheet1", "dest": "Sheet1_Copy", "position": 1})
        source.Copy.assert_called_once()
        assert wb.ActiveSheet.Name == "Sheet1_Copy"


# ── _set_vba ──────────────────────────────────────────────────────────────────

class TestSetVba:
    def test_replaces_existing_code(self, wb):
        comp = MagicMock()
        code_module = MagicMock()
        code_module.CountOfLines = 5
        comp.CodeModule = code_module
        wb.VBProject.VBComponents.return_value = comp

        _set_vba(wb, {"module": "Module1", "code": "Sub New()\nEnd Sub"})
        code_module.DeleteLines.assert_called_once_with(1, 5)
        code_module.InsertLines.assert_called_once_with(1, "Sub New()\nEnd Sub")

    def test_empty_module_no_delete(self, wb):
        comp = MagicMock()
        code_module = MagicMock()
        code_module.CountOfLines = 0
        comp.CodeModule = code_module
        wb.VBProject.VBComponents.return_value = comp

        _set_vba(wb, {"module": "Module1", "code": "Sub X()\nEnd Sub"})
        code_module.DeleteLines.assert_not_called()
        code_module.InsertLines.assert_called_once()

    def test_empty_code_no_insert(self, wb):
        comp = MagicMock()
        code_module = MagicMock()
        code_module.CountOfLines = 3
        comp.CodeModule = code_module
        wb.VBProject.VBComponents.return_value = comp

        _set_vba(wb, {"module": "Module1", "code": ""})
        code_module.DeleteLines.assert_called_once()
        code_module.InsertLines.assert_not_called()


# ── _add_vba_module ───────────────────────────────────────────────────────────

class TestAddVbaModule:
    def test_adds_and_names_module(self, wb):
        new_comp = MagicMock()
        wb.VBProject.VBComponents.Add.return_value = new_comp
        _add_vba_module(wb, {"name": "NewMod", "code": "Sub Test()\nEnd Sub"})
        wb.VBProject.VBComponents.Add.assert_called_once_with(1)
        assert new_comp.Name == "NewMod"
        new_comp.CodeModule.InsertLines.assert_called_once_with(1, "Sub Test()\nEnd Sub")

    def test_no_insert_when_no_code(self, wb):
        new_comp = MagicMock()
        wb.VBProject.VBComponents.Add.return_value = new_comp
        _add_vba_module(wb, {"name": "Empty", "code": ""})
        new_comp.CodeModule.InsertLines.assert_not_called()

    def test_raises_apply_error_on_com_failure(self, wb):
        wb.VBProject.VBComponents.Add.side_effect = Exception("Trust Center blocked")
        with pytest.raises(ApplyError, match="Cannot add VBA module"):
            _add_vba_module(wb, {"name": "X", "code": ""})


# ── _delete_vba_module ────────────────────────────────────────────────────────

class TestDeleteVbaModule:
    def test_removes_component(self, wb):
        comp = MagicMock()
        wb.VBProject.VBComponents.return_value = comp
        _delete_vba_module(wb, {"name": "Module1"})
        wb.VBProject.VBComponents.Remove.assert_called_once_with(comp)

    def test_raises_when_not_found(self, wb):
        wb.VBProject.VBComponents.side_effect = Exception("not found")
        with pytest.raises(ApplyError):
            _delete_vba_module(wb, {"name": "Ghost"})


# ── _add_named_range ──────────────────────────────────────────────────────────

class TestAddNamedRange:
    def test_calls_names_add(self, wb):
        _add_named_range(wb, {"name": "MyRange", "refers_to": "=Sheet1!$A$1:$D$10"})
        wb.Names.Add.assert_called_once_with(Name="MyRange", RefersTo="=Sheet1!$A$1:$D$10")


# ── _delete_named_range ───────────────────────────────────────────────────────

class TestDeleteNamedRange:
    def test_calls_delete(self, wb):
        nr = MagicMock()
        wb.Names.return_value = nr
        _delete_named_range(wb, {"name": "MyRange"})
        nr.Delete.assert_called_once()

    def test_raises_apply_error_when_not_found(self, wb):
        wb.Names.side_effect = Exception("not found")
        with pytest.raises(ApplyError, match="Named range 'Ghost' not found"):
            _delete_named_range(wb, {"name": "Ghost"})


# ── _dispatch ─────────────────────────────────────────────────────────────────

class TestDispatch:
    def test_unknown_type_raises(self, wb):
        change = Change(type="do_magic", params={})
        with pytest.raises(ApplyError, match="Unknown change type: 'do_magic'"):
            _dispatch(wb, change)

    @pytest.mark.parametrize("op_type,params", [
        ("set_cell",          {"sheet": "S", "cell": "A1", "value": 1}),
        ("set_range",         {"sheet": "S", "range": "A1:B2", "values": [[1, 2]]}),
        ("clear_range",       {"sheet": "S", "range": "A1:Z100"}),
        ("delete_sheet",      {"name": "S"}),
        ("rename_sheet",      {"old_name": "Old", "new_name": "New"}),
        ("add_named_range",   {"name": "NR", "refers_to": "=Sheet1!$A$1"}),
    ])
    def test_known_types_dispatched(self, wb, op_type, params):
        """Known operation types should not raise ApplyError for the dispatch itself."""
        change = Change(type=op_type, params=params)
        # The underlying handler will fail (mock wb), but it should not be
        # an "Unknown change type" error — it may raise ApplyError for other reasons.
        try:
            _dispatch(wb, change)
        except ApplyError as exc:
            assert "Unknown change type" not in str(exc)
        except Exception:
            pass  # COM errors from the mock are acceptable
