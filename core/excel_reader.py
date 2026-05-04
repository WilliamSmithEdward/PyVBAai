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

                values_2d = _normalise_range(raw_values, row_count)
                formulas_2d = (
                    _normalise_range(raw_formulas, row_count)
                    if raw_formulas is not None
                    else None
                )

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
                        sheet_data.cells[addr] = CellData(
                            row=r, col=c, address=addr, value=val, formula=formula,
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
