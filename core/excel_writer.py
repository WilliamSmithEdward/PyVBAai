# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Apply AI-generated changes to an Excel workbook via COM automation.

All public functions must be called from a thread where
pythoncom.CoInitialize() has already been called.
"""
from __future__ import annotations

import os
from typing import Any

from models.conversation import Change

# VBA module type constants
_VBA_MODULE = 1
_VBA_CLASS = 2
_VBA_FORM = 3

# Built once at import time — avoids a dict allocation on every _dispatch call
_HANDLERS: dict = {}  # populated after function definitions below


class ApplyError(Exception):
    """Raised when a change operation fails."""


def apply_changes(file_path: str, changes: list[Change]) -> None:
    """
    Open *file_path* in Excel (writable), apply every change, save and close.

    Raises ApplyError with a descriptive message on the first failure.
    Callers should create a backup BEFORE calling this function.
    """
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    wb = None
    norm_path = os.path.normcase(os.path.abspath(file_path))

    try:
        # Attach to a running Excel if possible
        try:
            excel = win32.GetActiveObject("Excel.Application")
            for w in excel.Workbooks:
                try:
                    if os.path.normcase(w.FullName) == norm_path:
                        wb = w
                        break
                except Exception:
                    pass
        except Exception:
            excel = None

        if excel is None:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

        if wb is None:
            wb = excel.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False)

        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.Calculation = -4135  # xlCalculationManual — no recalc between writes
        except Exception:  # noqa: BLE001
            pass

        for change in changes:
            _dispatch(wb, change)

        # Save in original format
        wb.Save()

    except ApplyError:
        raise
    except Exception as exc:
        raise ApplyError(f"COM error: {exc}") from exc
    finally:
        try:
            excel.EnableEvents = True
        except Exception:  # noqa: BLE001
            pass
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:  # noqa: BLE001
            pass
        excel.DisplayAlerts = True
        excel.ScreenUpdating = True
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None and excel.Workbooks.Count == 0:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# ── dispatcher ───────────────────────────────────────────────────────────────

def _dispatch(wb: Any, change: Change) -> None:
    op = change.type
    handler = _HANDLERS.get(op)
    if handler is None:
        raise ApplyError(f"Unknown change type: '{op}'")
    handler(wb, change.params)


# ── sheet helpers ─────────────────────────────────────────────────────────────

def _get_sheet(wb: Any, name: str) -> Any:
    try:
        return wb.Worksheets(name)
    except Exception as exc:
        raise ApplyError(f"Sheet not found: '{name}'") from exc


# ── cell / range operations ──────────────────────────────────────────────────

def _set_cell(wb: Any, p: dict) -> None:
    sheet = _get_sheet(wb, p["sheet"])
    cell = sheet.Range(p["cell"])
    if "formula" in p:
        cell.Formula = p["formula"]
    else:
        cell.Value = p["value"]


def _set_range(wb: Any, p: dict) -> None:
    sheet = _get_sheet(wb, p["sheet"])
    values: list[list] = p["values"]
    rng_addr: str = p["range"]
    rng = sheet.Range(rng_addr)
    # COM expects a tuple-of-tuples
    rng.Value = tuple(tuple(row) for row in values)


def _clear_range(wb: Any, p: dict) -> None:
    sheet = _get_sheet(wb, p["sheet"])
    sheet.Range(p["range"]).ClearContents()


# ── sheet structure operations ───────────────────────────────────────────────

def _add_sheet(wb: Any, p: dict) -> None:
    name: str = p["name"]
    position: int = p.get("position", wb.Worksheets.Count + 1)

    # Clamp position to valid range
    position = max(1, min(position, wb.Worksheets.Count + 1))

    if position <= wb.Worksheets.Count:
        new_sheet = wb.Worksheets.Add(Before=wb.Worksheets(position))
    else:
        new_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))

    new_sheet.Name = name


def _delete_sheet(wb: Any, p: dict) -> None:
    sheet = _get_sheet(wb, p["name"])
    sheet.Delete()


def _rename_sheet(wb: Any, p: dict) -> None:
    sheet = _get_sheet(wb, p["old_name"])
    sheet.Name = p["new_name"]


def _move_sheet(wb: Any, p: dict) -> None:
    sheet = _get_sheet(wb, p["name"])
    position: int = p["position"]
    position = max(1, min(position, wb.Worksheets.Count))
    if position <= wb.Worksheets.Count:
        sheet.Move(Before=wb.Worksheets(position))
    else:
        sheet.Move(After=wb.Worksheets(wb.Worksheets.Count))


def _copy_sheet(wb: Any, p: dict) -> None:
    source = _get_sheet(wb, p["source"])
    position: int = p.get("position", wb.Worksheets.Count + 1)
    position = max(1, min(position, wb.Worksheets.Count + 1))
    if position <= wb.Worksheets.Count:
        source.Copy(Before=wb.Worksheets(position))
    else:
        source.Copy(After=wb.Worksheets(wb.Worksheets.Count))
    # Copy() doesn't return the new sheet in win32com; ActiveSheet is the copy
    wb.ActiveSheet.Name = p["dest"]


# ── VBA operations ────────────────────────────────────────────────────────────

def _get_vba_component(wb: Any, name: str) -> Any:
    try:
        return wb.VBProject.VBComponents(name)
    except Exception as exc:
        raise ApplyError(
            f"VBA module '{name}' not found. "
            "Make sure 'Trust access to the VBA project object model' is enabled."
        ) from exc

def _set_vba(wb: Any, p: dict) -> None:
    comp = _get_vba_component(wb, p["module"])
    code_module = comp.CodeModule
    if code_module.CountOfLines > 0:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code: str = p["code"]
    if code:
        code_module.InsertLines(1, code)


def _add_vba_module(wb: Any, p: dict) -> None:
    try:
        # vbext_ct_StdModule = 1
        comp = wb.VBProject.VBComponents.Add(1)
        comp.Name = p["name"]
        code: str = p.get("code", "")
        if code:
            comp.CodeModule.InsertLines(1, code)
    except Exception as exc:
        raise ApplyError(
            f"Cannot add VBA module: {exc}. "
            "Enable 'Trust access to the VBA project object model'."
        ) from exc


def _delete_vba_module(wb: Any, p: dict) -> None:
    comp = _get_vba_component(wb, p["name"])
    wb.VBProject.VBComponents.Remove(comp)


# ── named range operations ────────────────────────────────────────────────────

def _add_named_range(wb: Any, p: dict) -> None:
    wb.Names.Add(Name=p["name"], RefersTo=p["refers_to"])


def _delete_named_range(wb: Any, p: dict) -> None:
    try:
        wb.Names(p["name"]).Delete()
    except Exception as exc:
        raise ApplyError(f"Named range '{p['name']}' not found.") from exc


# Populated here so all handler functions are defined before assignment
_HANDLERS.update({
    "set_cell":           _set_cell,
    "set_range":          _set_range,
    "clear_range":        _clear_range,
    "add_sheet":          _add_sheet,
    "delete_sheet":       _delete_sheet,
    "rename_sheet":       _rename_sheet,
    "move_sheet":         _move_sheet,
    "copy_sheet":         _copy_sheet,
    "set_vba":            _set_vba,
    "add_vba_module":     _add_vba_module,
    "delete_vba_module":  _delete_vba_module,
    "add_named_range":    _add_named_range,
    "delete_named_range": _delete_named_range,
})
