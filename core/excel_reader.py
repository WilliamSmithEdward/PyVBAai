# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Excel workbook extraction via COM automation.

Runs on any thread that calls pythoncom.CoInitialize() first.
All COM work is done inside read_workbook(); callers on a QThread
should CoInitialize in their run() method.
"""
from __future__ import annotations

import os
from functools import lru_cache
from typing import Any

from models.workbook import (
    CellData,
    CellFormat,
    ContextConfig,
    NamedRange,
    SheetData,
    VBAModule,
    WorkbookData,
)

# Maximum columns we ever read per sheet (sanity limit)
_MAX_COLS_HARD = 200


@lru_cache(maxsize=512)
def col_letter(col: int) -> str:
    """1-based column index -> letter(s), e.g. 1->'A', 27->'AA'. Cached."""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_address(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def _abs_addr(row: int, col: int) -> str:
    return f"${col_letter(col)}${row}"


def _connected_area_addresses(cells: dict) -> list[str]:
    """Return Excel-style absolute addresses for each connected block of non-empty cells.

    Uses 4-connected flood-fill on the (row, col) coordinates present in `cells`.
    Returns addresses sorted top-to-bottom, left-to-right by top-left corner.
    """
    occupied: set[tuple[int, int]] = {(cd.row, cd.col) for cd in cells.values()}
    visited: set[tuple[int, int]] = set()
    areas: list[str] = []

    for seed in sorted(occupied):
        if seed in visited:
            continue
        # BFS flood-fill
        component: list[tuple[int, int]] = []
        queue = [seed]
        visited.add(seed)
        while queue:
            r, c = queue.pop()
            component.append((r, c))
            for nr, nc in ((r - 1, c), (r + 1, c), (r, c - 1), (r, c + 1)):
                if (nr, nc) in occupied and (nr, nc) not in visited:
                    visited.add((nr, nc))
                    queue.append((nr, nc))
        min_r = min(p[0] for p in component)
        max_r = max(p[0] for p in component)
        min_c = min(p[1] for p in component)
        max_c = max(p[1] for p in component)
        areas.append(f"{_abs_addr(min_r, min_c)}:{_abs_addr(max_r, max_c)}")

    return areas


def read_workbook(file_path: str, config: ContextConfig | None = None) -> WorkbookData:
    """
    Extract workbook data using Excel COM automation.

    Launches a hidden Excel Application, opens the file read-only, reads all
    data (including non-contiguous used-range areas), then closes the workbook
    and quits Excel.
    """
    import pythoncom
    import win32com.client as win32

    if config is None:
        config = ContextConfig()

    pythoncom.CoInitialize()

    name = os.path.basename(file_path)
    wb_data = WorkbookData(file_path=file_path, name=name)
    try:
        wb_data.loaded_mtime = os.path.getmtime(file_path)
    except OSError:
        pass

    excel = None
    wb = None

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
            excel.Calculation = -4135  # xlCalculationManual — skip recalc on open
        except Exception:  # noqa: BLE001
            pass  # Some Excel configurations don't allow changing this

        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        excel.Visible = False  # hide after open succeeds

        # ── Sheets ──────────────────────────────────────────────────────────
        for i in range(1, wb.Worksheets.Count + 1):
            try:
                sheet = wb.Worksheets(i)

                # Visibility: xlSheetVisible = -1
                visible = (sheet.Visible == -1)

                # Skip if this sheet is excluded by config
                if (
                    config.included_sheets is not None
                    and sheet.Name not in config.included_sheets
                ):
                    wb_data.sheets.append(SheetData(
                        index=i - 1, name=sheet.Name,
                        used_range_address="", row_count=0, col_count=0,
                        visible=visible,
                    ))
                    continue

                used = sheet.UsedRange
                row_count = used.Rows.Count
                col_count = min(used.Columns.Count, _MAX_COLS_HARD)
                start_row = used.Row
                start_col = used.Column

                sheet_data = SheetData(
                    index=i - 1,
                    name=sheet.Name,
                    used_range_address=used.Address,
                    row_count=row_count,
                    col_count=col_count,
                    visible=visible,
                )

                # Read the whole used range in two COM calls (Value + Formula).
                # UsedRange is always a single bounding-box rectangle, so
                # iterating .Areas is unnecessary overhead.
                raw_values = used.Value
                raw_formulas = used.Formula if config.include_formulas else None

                # Format bulk-reads — each property costs one COM round-trip for
                # the entire range, far cheaper than per-cell access.
                raw_nf = used.NumberFormat    if config.include_number_format else None
                raw_bold = used.Font.Bold     if config.include_font_style    else None
                raw_italic = used.Font.Italic if config.include_font_style    else None
                raw_fc = used.Font.Color      if config.include_font_color    else None
                raw_bg = used.Interior.Color  if config.include_bg_color      else None
                raw_ha = used.HorizontalAlignment if config.include_alignment else None
                raw_wt = used.WrapText            if config.include_alignment else None

                values_2d = _normalise_range(raw_values, row_count)
                formulas_2d = (
                    _normalise_range(raw_formulas, row_count)
                    if raw_formulas is not None
                    else None
                )
                nf_2d     = _normalise_range(raw_nf,     row_count) if raw_nf     is not None else None
                bold_2d   = _normalise_range(raw_bold,   row_count) if raw_bold   is not None else None
                italic_2d = _normalise_range(raw_italic, row_count) if raw_italic is not None else None
                fc_2d     = _normalise_range(raw_fc,     row_count) if raw_fc     is not None else None
                bg_2d     = _normalise_range(raw_bg,     row_count) if raw_bg     is not None else None
                ha_2d     = _normalise_range(raw_ha,     row_count) if raw_ha     is not None else None
                wt_2d     = _normalise_range(raw_wt,     row_count) if raw_wt     is not None else None

                for r_idx, row in enumerate(values_2d):
                    for c_idx, val in enumerate(row[:col_count]):
                        if val is None or val == "":
                            continue
                        r = start_row + r_idx
                        c = start_col + c_idx
                        addr = cell_address(r, c)
                        formula = ""
                        if formulas_2d:
                            f = formulas_2d[r_idx][c_idx]
                            if isinstance(f, str) and f.startswith("="):
                                formula = f
                        fmt = _read_fmt(
                            r_idx, c_idx,
                            nf_2d, bold_2d, italic_2d, fc_2d, bg_2d, ha_2d, wt_2d,
                        ) if config.include_any_format else None
                        sheet_data.cells[addr] = CellData(
                            row=r, col=c, address=addr, value=val,
                            formula=formula, fmt=fmt,
                        )

                wb_data.sheets.append(sheet_data)
                sheet_data.area_addresses = _connected_area_addresses(sheet_data.cells)

            except Exception:  # noqa: BLE001
                # Still add a placeholder so the sheet appears in the tree
                wb_data.sheets.append(SheetData(
                    index=i - 1, name=f"Sheet{i}",
                    used_range_address="", row_count=0, col_count=0,
                ))

        # ── Named Ranges ─────────────────────────────────────────────────────
        if config.include_named_ranges:
            try:
                for nm in wb.Names:
                    try:
                        wb_data.named_ranges.append(
                            NamedRange(name=nm.Name, refers_to=nm.RefersTo)
                        )
                    except Exception:
                        pass
            except Exception:
                pass

        # ── VBA ───────────────────────────────────────────────────────────────
        if config.include_vba:
            try:
                vba_project = wb.VBProject
                type_map = {1: "Module", 2: "Class", 3: "Form", 100: "Document"}
                for comp in vba_project.VBComponents:
                    try:
                        mname = comp.Name
                        if (
                            config.included_vba_modules is not None
                            and mname not in config.included_vba_modules
                        ):
                            continue
                        lines = comp.CodeModule.CountOfLines
                        code = (
                            comp.CodeModule.Lines(1, lines)
                            if lines > 0
                            else ""
                        )
                        wb_data.vba_modules.append(
                            VBAModule(
                                name=mname,
                                module_type=comp.Type,
                                type_name=type_map.get(comp.Type, "Unknown"),
                                code=code,
                            )
                        )
                        wb_data.has_vba = True
                    except Exception:
                        pass
            except Exception:
                # VBA access blocked (Trust Center setting)
                wb_data.extraction_error = (
                    "VBA extraction blocked. Enable 'Trust access to the VBA "
                    "project object model' in Excel > Trust Center > Macro Settings."
                )

        return wb_data

    except Exception as exc:
        wb_data.extraction_error = str(exc)
        return wb_data

    finally:
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


# ── helpers ──────────────────────────────────────────────────────────────────

def _normalise_range(raw: Any, rows: int) -> list[list]:
    """
    COM returns different types depending on range shape:
      single cell  -> scalar
      single row   -> tuple of scalars
      single col   -> tuple of scalars
      multi-cell   -> tuple of tuples
    Normalise to list[list] using only row count to disambiguate.
    """
    if rows == 1:
        if not isinstance(raw, (list, tuple)):
            return [[raw]]  # 1x1 scalar
        return [list(raw)]  # 1xN row
    # rows > 1
    if isinstance(raw, (list, tuple)) and raw and isinstance(raw[0], (list, tuple)):
        return [list(r) for r in raw]  # MxN
    return [[v] for v in raw]  # Mx1 single column


# xl HorizontalAlignment constants → compact label
_HA_MAP: dict[int, str] = {
    -4108: "center", -4131: "left", -4152: "right",
    -4130: "justify", 1: "general",
}

# BGR int (win32com) → 6-char hex RGB string; None/xlNone (4294967295) → None
def _bgr_to_hex(val: Any) -> str | None:
    if val is None:
        return None
    iv = int(val)
    if iv < 0 or iv == 4294967295:  # xlColorIndexNone
        return None
    b = (iv >> 16) & 0xFF
    g = (iv >> 8)  & 0xFF
    r =  iv        & 0xFF
    return f"{r:02X}{g:02X}{b:02X}"


def _read_fmt(
    r: int, c: int,
    nf_2d: list[list] | None,
    bold_2d: list[list] | None,
    italic_2d: list[list] | None,
    fc_2d: list[list] | None,
    bg_2d: list[list] | None,
    ha_2d: list[list] | None,
    wt_2d: list[list] | None,
) -> CellFormat | None:
    """Extract a CellFormat from pre-read 2-D arrays; returns None if all defaults."""
    nf     = nf_2d[r][c]     if nf_2d     else None
    bold   = bold_2d[r][c]   if bold_2d   else None
    italic = italic_2d[r][c] if italic_2d else None
    fc     = _bgr_to_hex(fc_2d[r][c])  if fc_2d else None
    bg     = _bgr_to_hex(bg_2d[r][c])  if bg_2d else None
    ha_raw = ha_2d[r][c]     if ha_2d     else None
    ha     = _HA_MAP.get(int(ha_raw), None) if ha_raw is not None else None
    wt     = wt_2d[r][c]     if wt_2d     else None

    # Skip "General" number format — it's the default
    if isinstance(nf, str) and nf.strip().lower() == "general":
        nf = None
    # Skip black font (0) — default
    if fc == "000000":
        fc = None
    # Skip white/no background — default (no fill)
    if bg in ("FFFFFF", "000000", None):
        bg = None
    # Skip "general" alignment
    if ha == "general":
        ha = None

    fmt = CellFormat(
        number_format=nf or None,
        bold=True if bold else None,
        italic=True if italic else None,
        font_color=fc,
        bg_color=bg,
        h_align=ha,
        wrap_text=True if wt else None,
    )
    # Return None if nothing non-default was found
    if not any([
        fmt.number_format, fmt.bold, fmt.italic,
        fmt.font_color, fmt.bg_color, fmt.h_align, fmt.wrap_text,
    ]):
        return None
    return fmt
