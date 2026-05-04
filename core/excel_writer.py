# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Apply AI-generated changes to an Excel workbook.

All operations use openpyxl directly -- no Excel process needed.
VBA operations (set_vba, add_vba_module, delete_vba_module) are handled
by the UI as copy/paste and are never passed to this module.
"""
from __future__ import annotations

import os
import shutil
import tempfile
from typing import Any

from app.logger import get_logger
from models.conversation import Change

_log = get_logger(__name__)


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

    # ── Apply all changes via openpyxl ────────────────────────────────────────
    import openpyxl
    saved_path = os.path.abspath(file_path)
    owb = openpyxl.load_workbook(file_path, keep_vba=True, keep_links=False)
    try:
        for change in changes:
            _dispatch_openpyxl(owb, change)

        tmp_dir = tempfile.mkdtemp()
        try:
            tmp_path = os.path.join(tmp_dir, os.path.basename(saved_path))
            owb.save(tmp_path)
            shutil.move(tmp_path, saved_path)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
    finally:
        owb.close()

    return saved_path


# ── openpyxl dispatcher ───────────────────────────────────────────────────────

_OPENPYXL_HANDLERS: dict = {}


def _dispatch_openpyxl(owb: Any, change: Change) -> None:
    op = change.type
    handler = _OPENPYXL_HANDLERS.get(op)
    if handler is None:
        raise ApplyError(f"Unknown change type: '{op}'")
    _log.debug("dispatch: %s  params=%s", op, change.params)
    handler(owb, change.params)


# ── compact-flag helpers (AI write-back syntax) ──────────────────────────────

_FLAG_BOOL: dict[str, str] = {
    "B": "bold", "I": "italic", "S": "strikethrough",
    "U": "underline", "W": "wrap_text",
}
_FLAG_H_ALIGN: dict[str, str] = {
    "l": "left", "c": "center", "r": "right", "f": "fill", "j": "justify",
}
_FLAG_V_ALIGN: dict[str, str] = {"t": "top", "c": "center", "b": "bottom"}

_ALL_FMT_KEYS = frozenset({
    "bold", "italic", "strikethrough", "underline", "wrap_text",
    "number_format", "font_name", "font_size", "font_color", "bg_color",
    "h_align", "v_align",
    "border_top", "border_bottom", "border_left", "border_right",
})

_COLOR_KEYS = frozenset({"font_color", "bg_color"})


def _to_argb(color: str) -> str:
    """Ensure a color string is 8-char ARGB (openpyxl requirement).

    Accepts 6-char RGB ('FF0000') or 8-char ARGB ('FFFF0000').
    Raises ValueError for anything else so the problem surfaces immediately.
    """
    c = color.strip().upper().lstrip("#")
    if len(c) == 6:
        return f"FF{c}"
    if len(c) == 8:
        return c
    raise ValueError(f"Invalid color value {color!r}: expected 6 or 8 hex chars")


def _normalize_colors(p: dict) -> dict:
    """Return a copy of p with all color keys normalised to 8-char ARGB."""
    out = dict(p)
    for key in _COLOR_KEYS:
        if key in out:
            out[key] = _to_argb(out[key])
    return out


def _expand_flags(p: dict) -> dict:
    """Parse compact 'flags' string into format kwargs and merge with p.

    The AI may send: {"type":"set_cell",...,"flags":"B,#$#,##0,bt:thin,~D9E1F2"}
    which is equivalent to explicit keys bold=True, number_format="$#,##0", etc.
    Colors are normalised to 8-char ARGB at parse time.
    """
    flags_str = p.get("flags")
    if not flags_str:
        return _normalize_colors(p)
    extra: dict = {}
    for flag in str(flags_str).split(","):
        flag = flag.strip()
        if not flag:
            continue
        if flag in _FLAG_BOOL:
            extra[_FLAG_BOOL[flag]] = True
        elif flag.startswith("#"):
            extra["number_format"] = flag[1:]
        elif flag.startswith("^"):
            extra["font_color"] = _to_argb(flag[1:])
        elif flag.startswith("~"):
            extra["bg_color"] = _to_argb(flag[1:])
        elif flag.startswith("fn:"):
            extra["font_name"] = flag[3:]
        elif flag.startswith("fs:"):
            try:
                extra["font_size"] = float(flag[3:])
            except ValueError:
                pass
        elif flag.startswith("ha:"):
            extra["h_align"] = _FLAG_H_ALIGN.get(flag[3:], flag[3:])
        elif flag.startswith("va:"):
            extra["v_align"] = _FLAG_V_ALIGN.get(flag[3:], flag[3:])
        elif flag.startswith("bt:"):
            extra["border_top"] = flag[3:]
        elif flag.startswith("bb:"):
            extra["border_bottom"] = flag[3:]
        elif flag.startswith("bl:"):
            extra["border_left"] = flag[3:]
        elif flag.startswith("br:"):
            extra["border_right"] = flag[3:]
    result = {**_normalize_colors(p), **extra}
    result.pop("flags", None)
    return result


# ── cell helpers ──────────────────────────────────────────────────────────────

def _get_openpyxl_sheet(owb: Any, name: str) -> Any:
    try:
        return owb[name]
    except KeyError as exc:
        raise ApplyError(f"Sheet not found: '{name}'") from exc


def _set_cell(owb: Any, p: dict) -> None:
    p = _expand_flags(p)
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    cell = ws[p["cell"]]
    if "formula" in p:
        cell.value = p["formula"]
    else:
        cell.value = p["value"]
    # Apply any inline format keys
    if _ALL_FMT_KEYS & p.keys():
        _set_format(owb, {"sheet": p["sheet"], "range": p["cell"],
                          **{k: p[k] for k in _ALL_FMT_KEYS if k in p}})


def _set_range(owb: Any, p: dict) -> None:
    p = _expand_flags(p)
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    import openpyxl.utils
    rng = p["range"]
    min_col_letter, min_row, max_col_letter, max_row = openpyxl.utils.range_boundaries(rng)
    assert min_col_letter is not None and min_row is not None
    for r_idx, row_data in enumerate(p["values"]):
        for c_idx, val in enumerate(row_data):
            ws.cell(row=min_row + r_idx, column=min_col_letter + c_idx, value=val)
    # Apply any inline format to the entire range
    if _ALL_FMT_KEYS & p.keys():
        _set_format(owb, {"sheet": p["sheet"], "range": rng,
                          **{k: p[k] for k in _ALL_FMT_KEYS if k in p}})


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


# ── table operations ─────────────────────────────────────────────────────────

def _create_table(owb: Any, p: dict) -> None:
    """Create an Excel Table (ListObject) on a worksheet.

    params: sheet, range, name (table display name), style (optional,
    default 'TableStyleMedium9').
    """
    from openpyxl.worksheet.table import Table, TableStyleInfo
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    table_name = p["name"]
    # Remove existing table with same name if present (idempotent)
    if table_name in {t.displayName for t in ws.tables.values()}:
        del ws.tables[table_name]
    style_name = p.get("style", "TableStyleMedium9")
    tbl = Table(displayName=table_name, ref=p["range"])
    tbl.tableStyleInfo = TableStyleInfo(
        name=style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tbl)


def _delete_table(owb: Any, p: dict) -> None:
    """Remove an Excel Table by name from its sheet."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    table_name = p["name"]
    if table_name not in ws.tables:
        raise ApplyError(f"Table '{table_name}' not found on sheet '{p['sheet']}'.")
    del ws.tables[table_name]


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
        if any(k in p for k in ("bold", "italic", "strikethrough", "underline", "font_name", "font_size", "font_color")):
            existing = cell.font or Font()
            fc = p.get("font_color")  # already ARGB from _expand_flags / _normalize_colors
            cell.font = Font(
                bold=p.get("bold", existing.bold),
                italic=p.get("italic", existing.italic),
                strike=p.get("strikethrough", existing.strike),
                underline="single" if p.get("underline", existing.underline) else None,
                name=p.get("font_name", existing.name),
                size=p.get("font_size", existing.size),
                color=fc if fc is not None else (existing.color or None),
            )
        if "bg_color" in p:
            cell.fill = PatternFill(fill_type="solid", fgColor=p["bg_color"])  # already ARGB
        if any(k in p for k in ("h_align", "v_align", "wrap_text")):
            existing_a = cell.alignment or Alignment()
            cell.alignment = Alignment(
                horizontal=p.get("h_align", existing_a.horizontal),
                vertical=p.get("v_align", existing_a.vertical),
                wrap_text=p.get("wrap_text", existing_a.wrap_text),
            )
        if any(k in p for k in ("border_top", "border_bottom", "border_left", "border_right")):
            from openpyxl.styles import Border, Side

            def _make_side(spec: str, existing_side: Any) -> Any:
                if not spec:
                    return existing_side
                parts = spec.split(":", 1)
                style = parts[0]
                color = parts[1] if len(parts) > 1 else "FF000000"
                if len(color) == 6:
                    color = "FF" + color
                return Side(border_style=style, color=color)

            existing_b = cell.border or Border()
            cell.border = Border(
                top=_make_side(p.get("border_top", ""), existing_b.top),
                bottom=_make_side(p.get("border_bottom", ""), existing_b.bottom),
                left=_make_side(p.get("border_left", ""), existing_b.left),
                right=_make_side(p.get("border_right", ""), existing_b.right),
            )


# ── sheet view / dimension operations ────────────────────────────────────────

def _freeze_panes(owb: Any, p: dict) -> None:
    """Freeze rows/columns at a cell.  cell='' or absent to unfreeze."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    cell = (p.get("cell") or "").strip() or None
    ws.freeze_panes = cell


def _set_col_width(owb: Any, p: dict) -> None:
    """Set column width.  columns='A' or 'A:D', width in Excel character units."""
    from openpyxl.utils import column_index_from_string, get_column_letter
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    width = float(p["width"])
    col_spec = str(p["columns"]).strip().upper()
    if ":" in col_spec:
        start_col, end_col = col_spec.split(":", 1)
        for idx in range(column_index_from_string(start_col), column_index_from_string(end_col) + 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
    else:
        ws.column_dimensions[col_spec].width = width


def _set_row_height(owb: Any, p: dict) -> None:
    """Set row height in points."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.row_dimensions[int(p["row"])].height = float(p["height"])


def _insert_rows(owb: Any, p: dict) -> None:
    """Insert blank rows above 'row' (1-based)."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.insert_rows(int(p["row"]), int(p.get("count", 1)))


def _delete_rows(owb: Any, p: dict) -> None:
    """Delete rows starting at 'row' (1-based)."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.delete_rows(int(p["row"]), int(p.get("count", 1)))


def _insert_cols(owb: Any, p: dict) -> None:
    """Insert blank columns before 'col' (letter, e.g. 'B')."""
    from openpyxl.utils import column_index_from_string
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.insert_cols(column_index_from_string(str(p["col"]).strip().upper()), int(p.get("count", 1)))


def _delete_cols(owb: Any, p: dict) -> None:
    """Delete columns starting at 'col' (letter)."""
    from openpyxl.utils import column_index_from_string
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.delete_cols(column_index_from_string(str(p["col"]).strip().upper()), int(p.get("count", 1)))


def _set_tab_color(owb: Any, p: dict) -> None:
    """Set the sheet tab color.  color is RRGGBB or AARRGGBB.  '' to clear."""
    from openpyxl.styles.colors import Color
    ws = _get_openpyxl_sheet(owb, p["name"])
    color_str = (p.get("color") or "").strip()
    ws.sheet_properties.tabColor = Color(rgb=_to_argb(color_str)) if color_str else None


def _set_auto_filter(owb: Any, p: dict) -> None:
    """Set or clear the auto-filter on a sheet.  range='' to clear."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    rng = (p.get("range") or "").strip()
    ws.auto_filter.ref = rng or None


def _hide_rows(owb: Any, p: dict) -> None:
    """Hide rows.  row = 1-based start row, count = number of rows (default 1)."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    for r in range(int(p["row"]), int(p["row"]) + int(p.get("count", 1))):
        ws.row_dimensions[r].hidden = True


def _unhide_rows(owb: Any, p: dict) -> None:
    """Unhide rows."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    for r in range(int(p["row"]), int(p["row"]) + int(p.get("count", 1))):
        ws.row_dimensions[r].hidden = False


def _hide_cols(owb: Any, p: dict) -> None:
    """Hide columns.  col = letter (e.g. 'B'), count = number of cols (default 1)."""
    from openpyxl.utils import column_index_from_string, get_column_letter
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    start = column_index_from_string(str(p["col"]).strip().upper())
    for idx in range(start, start + int(p.get("count", 1))):
        ws.column_dimensions[get_column_letter(idx)].hidden = True


def _unhide_cols(owb: Any, p: dict) -> None:
    """Unhide columns."""
    from openpyxl.utils import column_index_from_string, get_column_letter
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    start = column_index_from_string(str(p["col"]).strip().upper())
    for idx in range(start, start + int(p.get("count", 1))):
        ws.column_dimensions[get_column_letter(idx)].hidden = False


def _clear_format(owb: Any, p: dict) -> None:
    """Reset all formatting on a cell range to defaults (values are untouched)."""
    import openpyxl.utils
    from openpyxl.styles import Alignment, Border, Font, PatternFill
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    try:
        min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(p["range"])
        cells = [ws.cell(row=r, column=c)
                 for r in range(min_row, max_row + 1)
                 for c in range(min_col, max_col + 1)]
    except Exception:  # noqa: BLE001
        cells = [ws[p["range"]]]
    for cell in cells:
        cell.font = Font()
        cell.fill = PatternFill()
        cell.border = Border()
        cell.alignment = Alignment()
        cell.number_format = "General"


def _protect_sheet(owb: Any, p: dict) -> None:
    """Password-protect a sheet.  password='' for no password."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    password = (p.get("password") or "").strip()
    ws.protection.sheet = True
    if password:
        ws.protection.set_password(password)


def _unprotect_sheet(owb: Any, p: dict) -> None:
    """Remove sheet protection."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.protection.sheet = False


def _set_print_area(owb: Any, p: dict) -> None:
    """Set the print area for a sheet.  range='' to clear."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    rng = (p.get("range") or "").strip()
    ws.print_area = rng or None


def _set_zoom(owb: Any, p: dict) -> None:
    """Set the sheet zoom level (10-400)."""
    ws = _get_openpyxl_sheet(owb, p["sheet"])
    ws.sheet_view.zoomScale = max(10, min(400, int(p["zoom"])))


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
    "create_table":       _create_table,
    "delete_table":       _delete_table,
    "freeze_panes":       _freeze_panes,
    "set_col_width":      _set_col_width,
    "set_row_height":     _set_row_height,
    "insert_rows":        _insert_rows,
    "delete_rows":        _delete_rows,
    "insert_cols":        _insert_cols,
    "delete_cols":        _delete_cols,
    "set_tab_color":      _set_tab_color,
    "auto_filter":        _set_auto_filter,
    "hide_rows":          _hide_rows,
    "unhide_rows":        _unhide_rows,
    "hide_cols":          _hide_cols,
    "unhide_cols":        _unhide_cols,
    "clear_format":       _clear_format,
    "protect_sheet":      _protect_sheet,
    "unprotect_sheet":    _unprotect_sheet,
    "set_print_area":     _set_print_area,
    "set_zoom":           _set_zoom,
    "set_named_range":   _add_named_range,
    "delete_named_range": _delete_named_range,
    "merge_cells":       _merge_cells,
    "unmerge_cells":     _unmerge_cells,
    "set_format":        _set_format,
})
