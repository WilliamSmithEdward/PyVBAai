# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Build a token-efficient context string from workbook data.

Cell notation: R<row>: <col>=<val>[flags]
  Flags: B=bold  I=italic  U=underline  W=wrap
         #<spec>=num_fmt   ^<hex>=font_color   ~<hex>=bg_color
         ha:<l|c|r>=h_align   va:<t|c|b>=v_align
  Formulas use braces: {SUM(A:A)}   Strings are double-quoted.

This format is designed to be compact yet unambiguous so the AI can
reliably reference sheet names, cell addresses, and VBA modules.
"""
from __future__ import annotations

import re

from models.workbook import CellFormat, ContextConfig, WorkbookData

# Rough token estimate: 1 token ~= 4 characters
_CHARS_PER_TOKEN = 4
# Fallback hard cap; overridden at runtime by the value from QSettings
_DEFAULT_MAX_CONTEXT_CHARS = 150_000

# Matches absolute Excel area addresses like $A$1:$C$10
_AREA_ADDR_RE = re.compile(r'\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)')

# Legend emitted once in the workbook header
_CELL_LEGEND = (
    "CELL NOTATION: R<row>: <COL>=<val>[flags]  "
    "val: \"str\" | number | {formula} | None\n"
    "  flags: B=bold I=italic S=strikethrough U=underline W=wrap_text\n"
    "         #<spec>=number_format  fn:<name>=font_name  fs:<pts>=font_size\n"
    "         ^<RRGGBB>=font_color  ~<RRGGBB>=bg_color\n"
    "         ha:<l|c|r>=h_align  va:<t|c|b>=v_align\n"
    "         bt/bb/bl/br:<style>[:<RRGGBB>]=border (thin|medium|thick|dashed|dotted|double|hair)"
)


# -- helpers ------------------------------------------------------------------

def _col_letter(col: int) -> str:
    """1-based column index -> letter(s), e.g. 1->'A', 27->'AA'."""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _col_letters_to_num(s: str) -> int:
    """Convert column letter(s) to 1-based column index (A->1, Z->26, AA->27)."""
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def _parse_area_addr(addr: str) -> tuple[int, int, int, int] | None:
    """Parse '$A$1:$C$10' into (r1, c1, r2, c2), 1-based. Returns None if unparseable."""
    m = _AREA_ADDR_RE.fullmatch(addr.strip())
    if not m:
        return None
    return (
        int(m.group(2)), _col_letters_to_num(m.group(1)),
        int(m.group(4)), _col_letters_to_num(m.group(3)),
    )


def _group_into_runs(row_nums: list[int]) -> list[list[int]]:
    """Group a sorted list of row numbers into contiguous runs."""
    if not row_nums:
        return []
    runs: list[list[int]] = []
    current: list[int] = [row_nums[0]]
    for n in row_nums[1:]:
        if n <= current[-1] + 1:
            current.append(n)
        else:
            runs.append(current)
            current = [n]
    runs.append(current)
    return runs


def _fmt_hints(fmt: CellFormat, include: set[str]) -> list[str]:
    """Return compact format flag tokens for the given CellFormat."""
    hints: list[str] = []
    if "number_format" in include and fmt.number_format:
        hints.append(f"#{fmt.number_format}")
    if "bold" in include and fmt.bold:
        hints.append("B")
    if "italic" in include and fmt.italic:
        hints.append("I")
    if "strikethrough" in include and fmt.strikethrough:
        hints.append("S")
    if "underline" in include and fmt.underline:
        hints.append("U")
    if "wrap_text" in include and fmt.wrap_text:
        hints.append("W")
    if "font_name" in include and fmt.font_name:
        hints.append(f"fn:{fmt.font_name}")
    if "font_size" in include and fmt.font_size:
        hints.append(f"fs:{fmt.font_size:g}")
    if "font_color" in include and fmt.font_color:
        hints.append(f"^{fmt.font_color}")
    if "bg_color" in include and fmt.bg_color:
        hints.append(f"~{fmt.bg_color}")
    if "h_align" in include and fmt.h_align:
        hints.append(f"ha:{fmt.h_align[0]}")
    if "v_align" in include and fmt.v_align:
        hints.append(f"va:{fmt.v_align[0]}")
    if "border_top" in include and fmt.border_top:
        hints.append(f"bt:{fmt.border_top}")
    if "border_bottom" in include and fmt.border_bottom:
        hints.append(f"bb:{fmt.border_bottom}")
    if "border_left" in include and fmt.border_left:
        hints.append(f"bl:{fmt.border_left}")
    if "border_right" in include and fmt.border_right:
        hints.append(f"br:{fmt.border_right}")
    return hints


# -- main entry point ---------------------------------------------------------

def build_context(wb: WorkbookData, config: ContextConfig | None = None) -> str:
    if config is None:
        config = ContextConfig()

    parts: list[str] = []

    # -- Header ---------------------------------------------------------------
    vba_names = [m.name for m in wb.vba_modules]
    nr_count = len(wb.named_ranges)

    parts.append(f"=== WORKBOOK: {wb.name} ===")
    def _sheet_label(s) -> str:  # type: ignore[no-untyped-def]
        suffix = " [hidden]" if not s.visible else ""
        return f"{s.name}({s.row_count}r\xd7{s.col_count}c){suffix}"

    parts.append(
        f"SHEETS ({len(wb.sheets)}): "
        + ", ".join(_sheet_label(s) for s in wb.sheets)
    )
    if vba_names:
        parts.append(f"VBA MODULES ({len(vba_names)}): " + ", ".join(vba_names))
    if nr_count and config.include_named_ranges:
        parts.append(f"NAMED RANGES ({nr_count}):")
        for nr in wb.named_ranges:
            parts.append(f"  {nr.name} = {nr.refers_to}")

    if wb.extraction_error:
        parts.append(f"[NOTE] {wb.extraction_error}")

    # Emit cell notation legend only when formatting may be present
    if config.fmt_include:
        parts.append(_CELL_LEGEND)

    parts.append("")

    # -- Cell Data ------------------------------------------------------------
    for sheet in wb.sheets:
        if (
            (config.included_sheets is not None
             and sheet.name not in config.included_sheets)
            or sheet.name in config.excluded_sheets
        ):
            parts.append(f"--- CELLS: {sheet.name} [excluded by settings] ---")
            continue

        used = f" [{sheet.used_range_address}]" if sheet.used_range_address else ""
        parts.append(
            f"--- CELLS: {sheet.name}{used} ({sheet.row_count}r \xd7 {sheet.col_count}c) ---"
        )

        # Excluded area rectangles (r1, c1, r2, c2) for this sheet
        excl_rects: list[tuple[int, int, int, int]] = []
        for area_addr in config.excluded_areas.get(sheet.name, []):
            parsed = _parse_area_addr(area_addr)
            if parsed is not None:
                excl_rects.append(parsed)

        if not sheet.cells:
            parts.append("  [empty]")
        else:
            # Collect rows, skipping excluded cells
            rows: dict[int, list[tuple[int, int]]] = {}  # row -> [(col, col_num)]
            for cell in sheet.cells.values():
                if excl_rects and any(
                    r1 <= cell.row <= r2 and c1 <= cell.col <= c2
                    for r1, c1, r2, c2 in excl_rects
                ):
                    continue
                rows.setdefault(cell.row, []).append((cell.col, cell.col))

            # Group consecutive rows into runs; optionally truncate each run
            sorted_row_nums = sorted(rows)
            runs = _group_into_runs(sorted_row_nums)

            for run in runs:
                limit = config.max_rows_per_area
                display = run[:limit] if limit is not None else run
                omitted = len(run) - len(display)

                for row_num in display:
                    row_parts: list[str] = []
                    for col_num, _ in sorted(rows[row_num]):
                        col_str = _col_letter(col_num)
                        addr = f"{col_str}{row_num}"
                        cell = sheet.cells[addr]
                        if cell.formula:
                            token = f"{col_str}={{{cell.formula[1:]}}}"
                        else:
                            val = cell.value
                            if isinstance(val, str):
                                truncated = val[:80] + "..." if len(val) > 80 else val
                                token = f'{col_str}="{truncated}"'
                            else:
                                token = f"{col_str}={val}"
                        if cell.fmt and config.fmt_include:
                            hints = _fmt_hints(cell.fmt, config.fmt_include)
                            if hints:
                                token += f"[{','.join(hints)}]"
                        row_parts.append(token)
                    parts.append(f"  R{row_num}: " + " ".join(row_parts))

                if omitted > 0:
                    parts.append(
                        f"  [{omitted} row{'s' if omitted != 1 else ''} omitted"
                        f" — increase 'Max rows per area' in Settings > Context]"
                    )

        # Charts on this sheet
        if sheet.charts:
            chart_strs = []
            for ch in sheet.charts:
                label = f'"{ch.title}"' if ch.title else "Untitled"
                loc = f" @ {ch.anchor}" if ch.anchor else ""
                chart_strs.append(f"{ch.chart_type} {label}{loc}")
            parts.append("  CHARTS: " + ", ".join(chart_strs))

        # Pivot tables on this sheet
        if sheet.pivot_tables:
            parts.append("  PIVOT TABLES: " + ", ".join(sheet.pivot_tables))

        parts.append("")

    # -- VBA ------------------------------------------------------------------
    if config.include_vba:
        for mod in wb.vba_modules:
            if (
                (config.included_vba_modules is not None
                 and mod.name not in config.included_vba_modules)
                or mod.name in config.excluded_vba_modules
            ):
                continue
            parts.append(f"--- VBA: {mod.name} ({mod.type_name}) ---")
            parts.append(mod.code if mod.code else "  [empty module]")
            parts.append("")

    context = "\n".join(parts)

    # Hard-trim if way over budget (should not normally happen)
    try:
        from app.config import get_settings
        max_chars = int(str(get_settings().value("context/max_chars", _DEFAULT_MAX_CONTEXT_CHARS)))
    except Exception:  # noqa: BLE001
        max_chars = _DEFAULT_MAX_CONTEXT_CHARS
    if len(context) > max_chars:
        context = (
            context[:max_chars]
            + "\n\n[CONTEXT TRUNCATED: workbook is too large to show fully. "
            "Apply changes to the sheets and ranges you can see above. "
            "Do NOT ask the user for more context — work with what is provided.]"
        )

    return context


def estimate_tokens(text: str) -> int:
    return max(1, len(text) // _CHARS_PER_TOKEN)
