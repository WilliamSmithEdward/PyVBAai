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


_VBA_OPS = {"set_vba", "add_vba_module", "delete_vba_module"}


def apply_changes(file_path: str, changes: list[Change]) -> str:
    """
    Open *file_path* in Excel (writable), apply every change, save and close.

    Returns the actual saved path.  If *file_path* is ``.xlsx`` and any VBA
    changes are present the workbook is promoted to ``.xlsm`` (Excel silently
    strips VBA when saving in the macro-free xlsx format).

    Raises ApplyError with a descriptive message on the first failure.
    Callers should create a backup BEFORE calling this function.
    """
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    wb = None
    saved_path = file_path

    # Fail fast with a clear message if the file is locked by another process.
    try:
        with open(file_path, "r+b"):
            pass
    except PermissionError as exc:
        fname = os.path.basename(file_path)
        raise ApplyError(
            f"\u2018{fname}\u2019 is open in another program. "
            "Close it and try again."
        ) from exc
    except OSError as exc:
        raise ApplyError(f"Cannot access \u2018{os.path.basename(file_path)}\u2019: {exc}") from exc

    try:
        # DispatchEx always spawns a new Excel process; Dispatch may reuse a
        # lingering/zombie instance from the ROT and hang indefinitely.
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True   # must be visible during Open — headless blocks on OneDrive sync
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable — suppresses macro dialogs
        try:
            excel.Calculation = -4135  # xlCalculationManual — no recalc between writes
        except Exception:  # noqa: BLE001
            pass

        wb = excel.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False)
        excel.Visible = False  # hide after open succeeds

        for change in changes:
            _dispatch(wb, change)

        # Save — upgrade to .xlsm when VBA changes are applied to an .xlsx file;
        # Excel silently strips all VBA when saving in the macro-free xlsx format.
        if any(c.type in _VBA_OPS for c in changes) and file_path.lower().endswith(".xlsx"):
            # SaveAs directly to a OneDrive/network folder fails because Excel can't
            # create its temp file there (-2147352567).  Work around by saving to
            # %TEMP% first, then moving the result to the final destination.
            import shutil
            import tempfile
            final_path = os.path.abspath(file_path[:-5] + ".xlsm")
            tmp_dir = tempfile.mkdtemp()
            try:
                tmp_path = os.path.join(tmp_dir, os.path.basename(final_path))
                wb.SaveAs(Filename=tmp_path, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
                # Close before moving so Excel releases the file handle; prevent
                # double-close in finally (the file will be gone from tmp_path).
                wb.Close(SaveChanges=False)
                wb = None
                shutil.move(tmp_path, final_path)
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            try:
                os.remove(file_path)
            except OSError:
                pass
            saved_path = final_path
        else:
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
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    return saved_path


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


# ── format operations ─────────────────────────────────────────────────────────

def _hex_to_bgr(hex_color: str) -> int:
    """Convert 6-char RGB hex string to Excel BGR int."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


# xl HorizontalAlignment constants
_HA_CONST: dict[str, int] = {
    "general": 1, "left": -4131, "center": -4108,
    "right": -4152, "justify": -4130,
}


def _format_cell(wb: Any, p: dict) -> None:
    """Apply formatting to a cell or range.

    Supported params (all optional):
      sheet, cell/range   — target (cell overrides range)
      number_format       — e.g. "#,##0.00"
      bold, italic        — bool
      font_color          — 6-char RGB hex, e.g. "FF0000"
      bg_color            — 6-char RGB hex
      h_align             — "left"|"center"|"right"|"general"|"justify"
      wrap_text           — bool
    """
    sheet = _get_sheet(wb, p["sheet"])
    ref = p.get("cell") or p.get("range")
    if not ref:
        raise ApplyError("format_cell requires 'cell' or 'range' param")
    rng = sheet.Range(ref)

    if "number_format" in p:
        rng.NumberFormat = p["number_format"]
    if "bold" in p:
        rng.Font.Bold = p["bold"]
    if "italic" in p:
        rng.Font.Italic = p["italic"]
    if "font_color" in p:
        rng.Font.Color = _hex_to_bgr(p["font_color"])
    if "bg_color" in p:
        rng.Interior.Color = _hex_to_bgr(p["bg_color"])
    if "h_align" in p:
        const = _HA_CONST.get(str(p["h_align"]).lower())
        if const is not None:
            rng.HorizontalAlignment = const
    if "wrap_text" in p:
        rng.WrapText = p["wrap_text"]


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
    "format_cell":        _format_cell,
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
