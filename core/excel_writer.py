# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Apply AI-generated changes to an Excel workbook.

Non-VBA operations (cell writes, sheet structure, named ranges, formatting) use
openpyxl directly -- no Excel process needed.

VBA operations (set_vba, add_vba_module, delete_vba_module) require COM because
VBA is stored as a binary OLE blob that openpyxl cannot modify.
"""
from __future__ import annotations

import os
import shutil
import tempfile
from typing import Any

from models.conversation import Change

_VBA_OPS = {"set_vba", "add_vba_module", "delete_vba_module"}


class ApplyError(Exception):
    """Raised when a change operation fails."""


def apply_changes(file_path: str, changes: list[Change]) -> str:
    """
    Apply every change to *file_path* and save.

    Returns the actual saved path (may differ from file_path if .xlsx was
    promoted to .xlsm due to VBA changes).

    Raises ApplyError on the first failure.
    Callers should create a backup BEFORE calling this function.
    """
    # Fail fast with a clear message if the file is locked.
    try:
        with open(file_path, "r+b"):
            pass
    except PermissionError as exc:
        fname = os.path.basename(file_path)
        raise ApplyError(
            f"'{fname}' is open in another program. Close it and try again."
        ) from exc
    except OSError as exc:
        raise ApplyError(f"Cannot access '{os.path.basename(file_path)}': {exc}") from exc

    vba_changes = [c for c in changes if c.type in _VBA_OPS]
    other_changes = [c for c in changes if c.type not in _VBA_OPS]
    has_vba_changes = bool(vba_changes)
    is_xlsx = file_path.lower().endswith(".xlsx")

    # Determine final saved path (promote .xlsx -> .xlsm when VBA is added)
    if has_vba_changes and is_xlsx:
        saved_path = os.path.abspath(file_path[:-5] + ".xlsm")
    else:
        saved_path = os.path.abspath(file_path)

    # ── Step 1: openpyxl for all non-VBA changes ─────────────────────────────
    if other_changes or has_vba_changes:
        import openpyxl
        owb = openpyxl.load_workbook(file_path, keep_vba=True, keep_links=False)
        for change in other_changes:
            _dispatch_openpyxl(owb, change)

        if has_vba_changes and is_xlsx:
            # Save to temp dir first (OneDrive / network share safe)
            tmp_dir = tempfile.mkdtemp()
            try:
                tmp_path = os.path.join(tmp_dir, os.path.basename(saved_path))
                owb.save(tmp_path)
                shutil.move(tmp_path, saved_path)
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            try:
                os.remove(file_path)
            except OSError:
                pass
        else:
            tmp_dir = tempfile.mkdtemp()
            try:
                tmp_path = os.path.join(tmp_dir, os.path.basename(saved_path))
                owb.save(tmp_path)
                shutil.move(tmp_path, saved_path)
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)

    # ── Step 2: COM for VBA-only changes ─────────────────────────────────────
    if vba_changes:
        _apply_vba_via_com(saved_path, vba_changes)

    return saved_path


# ── openpyxl dispatcher ───────────────────────────────────────────────────────

_OPENPYXL_HANDLERS: dict = {}


def _dispatch_openpyxl(owb: Any, change: Change) -> None:
    op = change.type
    handler = _OPENPYXL_HANDLERS.get(op)
    if handler is None:
        raise ApplyError(f"Unknown change type: '{op}'")
    handler(owb, change.params)


# ── cell helpers ──────────────────────────────────────────────────────────────

def _get_openpyxl_sheet(owb: Any, name: str) -> Any:
    try:
        return owb[name]
    except KeyError as exc:
        raise ApplyError(f"Sheet not found: '{name}'") from exc


def _set_cell(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    cell = ws[p["cell"]]
    if "formula" in p:
        cell.value = p["formula"]
    else:
        cell.value = p["value"]


def _set_range(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    import openpyxl.utils
    rng = p["range"]
    min_col_letter, min_row, max_col_letter, max_row = openpyxl.utils.range_boundaries(rng)
    assert min_col_letter is not None and min_row is not None
    for r_idx, row_data in enumerate(p["values"]):
        for c_idx, val in enumerate(row_data):
            ws.cell(row=min_row + r_idx, column=min_col_letter + c_idx, value=val)


def _clear_range(owb: Any, p: dict) -> None:
    import openpyxl.utils
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(p["range"])
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = None


# ── sheet structure ───────────────────────────────────────────────────────────

def _add_sheet(owb: Any, p: dict) -> None:
    name: str = p["name"]
    position: int = p.get("position", len(owb.sheetnames) + 1)
    position = max(1, min(position, len(owb.sheetnames) + 1))
    owb.create_sheet(title=name, index=position - 1)


def _delete_sheet(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["name"])
    del owb[ws.title]


def _rename_sheet(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["old_name"])
    ws.title = p["new_name"]


def _move_sheet(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["name"])
    target = max(0, p["position"] - 1)
    owb.move_sheet(ws, offset=target - owb.index(ws))


def _copy_sheet(owb: Any, p: dict) -> None:
    source = _get_openpyxl_sheet(owb, p["source"])
    idx = p.get("position", len(owb.sheetnames) + 1) - 1
    idx = max(0, min(idx, len(owb.sheetnames)))
    new_ws = owb.copy_worksheet(source)
    new_ws.title = p["dest"]
    owb.move_sheet(new_ws, offset=idx - owb.index(new_ws))


def _hide_sheet(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["name"])
    ws.sheet_state = "hidden"


def _unhide_sheet(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["name"])
    ws.sheet_state = "visible"


# ── named range operations ────────────────────────────────────────────────────

def _add_named_range(owb: Any, p: dict) -> None:
    from openpyxl.workbook.defined_name import DefinedName
    defn = DefinedName(name=p["name"], attr_text=p["refers_to"])
    owb.defined_names[p["name"]] = defn


def _delete_named_range(owb: Any, p: dict) -> None:
    name = p["name"]
    if name not in owb.defined_names:
        raise ApplyError(f"Named range '{name}' not found.")
    del owb.defined_names[name]


# ── merge operations ──────────────────────────────────────────────────────────

def _merge_cells(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.merge_cells(p["range"])


def _unmerge_cells(owb: Any, p: dict) -> None:
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.unmerge_cells(p["range"])


# ── format operation ──────────────────────────────────────────────────────────

def _set_format(owb: Any, p: dict) -> None:
    """Apply formatting to a cell or range.

    params: sheet, range, and any subset of: number_format, bold, italic,
    underline, font_name, font_size, font_color, bg_color, h_align, v_align,
    wrap_text.
    """
    import openpyxl.utils
    from openpyxl.styles import Alignment, Font, PatternFill
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    rng = p["range"]
    # Expand to all cells in range
    try:
        min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(rng)
        assert min_col is not None and min_row is not None and max_col is not None and max_row is not None
        cells = [ws.cell(row=r, column=c)
                 for r in range(min_row, max_row + 1)
                 for c in range(min_col, max_col + 1)]
    except Exception:  # noqa: BLE001
        cells = [ws[rng]]

    for cell in cells:
        if "number_format" in p:
            cell.number_format = p["number_format"]
        if any(k in p for k in ("bold", "italic", "underline", "font_name", "font_size", "font_color")):
            existing = cell.font or Font()
            cell.font = Font(
                bold=p.get("bold", existing.bold),
                italic=p.get("italic", existing.italic),
                underline="single" if p.get("underline", existing.underline) else None,
                name=p.get("font_name", existing.name),
                size=p.get("font_size", existing.size),
                color=p["font_color"] if "font_color" in p else (existing.color or None),
            )
        if "bg_color" in p:
            cell.fill = PatternFill(fill_type="solid", fgColor=p["bg_color"])
        if any(k in p for k in ("h_align", "v_align", "wrap_text")):
            existing_a = cell.alignment or Alignment()
            cell.alignment = Alignment(
                horizontal=p.get("h_align", existing_a.horizontal),
                vertical=p.get("v_align", existing_a.vertical),
                wrap_text=p.get("wrap_text", existing_a.wrap_text),
            )


# ── VBA via COM ───────────────────────────────────────────────────────────────

def _apply_vba_via_com(file_path: str, changes: list[Change]) -> None:
    """Apply VBA-only changes via COM. file_path must already be .xlsm."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise ApplyError("pywin32 is required for VBA operations.") from exc

    pythoncom.CoInitialize()
    excel = None
    com_wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True   # must be visible during Open -- headless blocks on OneDrive sync
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        try:
            excel.Calculation = -4135
        except Exception:  # noqa: BLE001
            pass

        com_wb = excel.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False)
        excel.Visible = False

        for change in changes:
            if change.type == "set_vba":
                _com_set_vba(com_wb, change.params)
            elif change.type == "add_vba_module":
                _com_add_vba_module(com_wb, change.params)
            elif change.type == "delete_vba_module":
                _com_delete_vba_module(com_wb, change.params)

        com_wb.Save()

    except ApplyError:
        raise
    except Exception as exc:
        raise ApplyError(f"COM VBA error: {exc}") from exc
    finally:
        if excel is not None:
            try:
                excel.EnableEvents = True
            except Exception:  # noqa: BLE001
                pass
            try:
                excel.Calculation = -4105
            except Exception:  # noqa: BLE001
                pass
            try:
                excel.DisplayAlerts = True
                excel.ScreenUpdating = True
            except Exception:  # noqa: BLE001
                pass
        try:
            if com_wb is not None:
                com_wb.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:  # noqa: BLE001
            pass
        pythoncom.CoUninitialize()


def _com_get_vba_component(com_wb: Any, name: str) -> Any:
    try:
        return com_wb.VBProject.VBComponents(name)
    except Exception as exc:
        raise ApplyError(
            f"VBA module '{name}' not found. "
            "Enable 'Trust access to the VBA project object model'."
        ) from exc


def _com_set_vba(com_wb: Any, p: dict) -> None:
    comp = _com_get_vba_component(com_wb, p["module"])
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    code: str = p["code"]
    if code:
        cm.InsertLines(1, code)


def _com_add_vba_module(com_wb: Any, p: dict) -> None:
    try:
        comp = com_wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = p["name"]
        code: str = p.get("code", "")
        if code:
            comp.CodeModule.InsertLines(1, code)
    except Exception as exc:
        raise ApplyError(
            f"Cannot add VBA module: {exc}. "
            "Enable 'Trust access to the VBA project object model'."
        ) from exc


def _com_delete_vba_module(com_wb: Any, p: dict) -> None:
    comp = _com_get_vba_component(com_wb, p["name"])
    com_wb.VBProject.VBComponents.Remove(comp)


# ── handler registry (module-level for performance) ───────────────────────────

_OPENPYXL_HANDLERS.update({
    "set_cell":          _set_cell,
    "set_range":         _set_range,
    "clear_range":       _clear_range,
    "add_sheet":         _add_sheet,
    "delete_sheet":      _delete_sheet,
    "rename_sheet":      _rename_sheet,
    "move_sheet":        _move_sheet,
    "copy_sheet":        _copy_sheet,
    "hide_sheet":        _hide_sheet,
    "unhide_sheet":      _unhide_sheet,
    "set_named_range":   _add_named_range,
    "delete_named_range": _delete_named_range,
    "merge_cells":       _merge_cells,
    "unmerge_cells":     _unmerge_cells,
    "set_format":        _set_format,
})
