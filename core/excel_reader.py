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


def col_letter(col: int) -> str:
    """1-based column index -> letter(s), e.g. 1->'A', 27->'AA'."""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_address(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


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

    excel = None
    wb = None

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )

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

                sheet_data = SheetData(
                    index=i - 1,
                    name=sheet.Name,
                    used_range_address=used.Address,
                    row_count=row_count,
                    col_count=col_count,
                    visible=visible,
                )

                # Iterate each area so non-contiguous ranges are fully read
                # and cell addresses are correct regardless of where data starts.
                for area in used.Areas:
                    area_rows = area.Rows.Count
                    area_cols = min(area.Columns.Count, _MAX_COLS_HARD)
                    start_row = area.Row
                    start_col = area.Column

                    if area_rows == 0 or area_cols == 0:
                        continue

                    rng = sheet.Range(
                        sheet.Cells(start_row, start_col),
                        sheet.Cells(start_row + area_rows - 1, start_col + area_cols - 1),
                    )
                    raw_values = rng.Value
                    raw_formulas = rng.Formula if config.include_formulas else None

                    values_2d = _normalise_range(raw_values, area_rows, area_cols)
                    formulas_2d = (
                        _normalise_range(raw_formulas, area_rows, area_cols)
                        if raw_formulas is not None
                        else None
                    )

                    for r_idx, row in enumerate(values_2d):
                        for c_idx, val in enumerate(row):
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

def _normalise_range(raw: object, rows: int, cols: int) -> list[list]:
    """
    COM returns different types depending on range shape:
      single cell  -> scalar
      single row   -> tuple of scalars
      single col   -> tuple of scalars
      multi-cell   -> tuple of tuples
    Normalise to list[list].
    """
    if rows == 1 and cols == 1:
        return [[raw]]
    if rows == 1:
        # raw is a flat tuple
        return [list(raw)]
    if cols == 1:
        # raw is a tuple of single-element tuples
        if isinstance(raw, (list, tuple)) and raw and isinstance(raw[0], (list, tuple)):
            return [list(r) for r in raw]
        return [[v] for v in raw]
    return [list(r) for r in raw]
