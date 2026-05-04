# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Build a token-efficient context string from workbook data.

The format is designed to be compact yet unambiguous so GPT can
reliably reference sheet names, cell addresses, and VBA modules.
"""
from __future__ import annotations

import re

from models.workbook import ContextConfig, WorkbookData

# Rough token estimate: 1 token ≈ 4 characters
_CHARS_PER_TOKEN = 4
# Leave this many characters free for conversation history + response
_MAX_CONTEXT_CHARS = 60_000

# Matches absolute Excel area addresses like $A$1:$C$10
_AREA_ADDR_RE = re.compile(r'\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)')


def _col_letters_to_num(s: str) -> int:
    """Convert column letter(s) to 1-based column index (A→1, Z→26, AA→27)."""
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def _parse_area_addr(addr: str) -> tuple[int, int, int, int] | None:
    """Parse '$A$1:$C$10' into (r1, c1, r2, c2), 1-based.  Returns None if unparseable."""
    m = _AREA_ADDR_RE.fullmatch(addr.strip())
    if not m:
        return None
    return (
        int(m.group(2)), _col_letters_to_num(m.group(1)),
        int(m.group(4)), _col_letters_to_num(m.group(3)),
    )


def build_context(wb: WorkbookData, config: ContextConfig | None = None) -> str:
    if config is None:
        config = ContextConfig()

    parts: list[str] = []

    # ── Header ───────────────────────────────────────────────────────────────
    vba_names = [m.name for m in wb.vba_modules]
    nr_count = len(wb.named_ranges)

    parts.append(f"=== WORKBOOK: {wb.name} ===")
    parts.append(
        f"SHEETS ({len(wb.sheets)}): "
        + ", ".join(f"{s.name}({s.row_count}r×{s.col_count}c)" for s in wb.sheets)
    )
    if vba_names:
        parts.append(f"VBA MODULES ({len(vba_names)}): " + ", ".join(vba_names))
    if nr_count and config.include_named_ranges:
        parts.append(f"NAMED RANGES ({nr_count}):")
        for nr in wb.named_ranges:
            parts.append(f"  {nr.name} = {nr.refers_to}")

    if wb.extraction_error:
        parts.append(f"[NOTE] {wb.extraction_error}")

    parts.append("")

    # ── Cell Data ─────────────────────────────────────────────────────────────
    for sheet in wb.sheets:
        if (
            (config.included_sheets is not None
             and sheet.name not in config.included_sheets)
            or sheet.name in config.excluded_sheets
        ):
            parts.append(f"--- CELLS: {sheet.name} [excluded by settings] ---")
            continue

        addr = f" [{sheet.used_range_address}]" if sheet.used_range_address else ""
        header = f"--- CELLS: {sheet.name}{addr} ({sheet.row_count}r × {sheet.col_count}c) ---"
        parts.append(header)

        # Excluded area rectangles (r1, c1, r2, c2) for this sheet
        excl_area_rects: list[tuple[int, int, int, int]] = []
        for area_addr in config.excluded_areas.get(sheet.name, []):
            parsed = _parse_area_addr(area_addr)
            if parsed is not None:
                excl_area_rects.append(parsed)

        if not sheet.cells:
            parts.append("  [empty]")
        else:
            # Group cells by row for compact display
            rows: dict[int, list[tuple[int, str]]] = {}
            for cell in sheet.cells.values():
                if excl_area_rects and any(
                    r1 <= cell.row <= r2 and c1 <= cell.col <= c2
                    for r1, c1, r2, c2 in excl_area_rects
                ):
                    continue
                rows.setdefault(cell.row, []).append((cell.col, cell.address))
            for row_num in sorted(rows):
                row_parts: list[str] = []
                for _col_num, addr in sorted(rows[row_num]):
                    cell = sheet.cells[addr]
                    if cell.formula:
                        row_parts.append(f"{addr}={{={cell.formula[1:]}}}")
                    else:
                        val = cell.value
                        if isinstance(val, str):
                            truncated = val[:80] + "..." if len(val) > 80 else val
                            row_parts.append(f'{addr}="{truncated}"')
                        else:
                            row_parts.append(f"{addr}={val}")
                    # Append compact format hint when non-default
                    if cell.fmt:
                        hints: list[str] = []
                        if cell.fmt.number_format:
                            hints.append(f"fmt:{cell.fmt.number_format}")
                        if cell.fmt.bold:
                            hints.append("bold")
                        if cell.fmt.italic:
                            hints.append("italic")
                        if cell.fmt.font_color:
                            hints.append(f"color:{cell.fmt.font_color}")
                        if cell.fmt.bg_color:
                            hints.append(f"bg:{cell.fmt.bg_color}")
                        if hints:
                            row_parts[-1] += f"[{','.join(hints)}]"
                parts.append("  " + ", ".join(row_parts))

        parts.append("")

    # ── VBA ───────────────────────────────────────────────────────────────────
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

    # Hard-trim if way over budget (shouldn't normally happen)
    if len(context) > _MAX_CONTEXT_CHARS:
        context = context[:_MAX_CONTEXT_CHARS] + "\n\n[CONTEXT TRUNCATED - too large]"

    return context


def estimate_tokens(text: str) -> int:
    return max(1, len(text) // _CHARS_PER_TOKEN)
