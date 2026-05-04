# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for core/context_builder.py."""
from __future__ import annotations

from core.context_builder import (
    _DEFAULT_MAX_CONTEXT_CHARS,
    _infer_spill_dims,
    _parse_area_addr,
    _spill_range,
    build_context,
    estimate_tokens,
)
from models.workbook import CellData, ContextConfig, SheetData, VBAModule, WorkbookData

# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_wb(**kwargs) -> WorkbookData:
    defaults = dict(file_path="C:/test/wb.xlsx", name="wb.xlsx")
    defaults.update(kwargs)
    return WorkbookData(**defaults)


def _make_sheet(name: str, cells: dict | None = None) -> SheetData:
    s = SheetData(index=1, name=name, used_range_address="A1", row_count=1, col_count=1)
    if cells:
        s.cells = cells
        s.row_count = max(c.row for c in cells.values())
        s.col_count = max(c.col for c in cells.values())
    return s


# ── estimate_tokens ───────────────────────────────────────────────────────────

class TestEstimateTokens:
    def test_empty_string(self):
        assert estimate_tokens("") == 1  # max(1, 0//4)

    def test_four_chars_is_one_token(self):
        assert estimate_tokens("abcd") == 1

    def test_exact_multiple(self):
        assert estimate_tokens("a" * 400) == 100

    def test_non_multiple(self):
        assert estimate_tokens("a" * 9) == 2  # 9//4 = 2

    def test_long_text(self):
        text = "x" * 4000
        assert estimate_tokens(text) == 1000


# ── build_context: header ─────────────────────────────────────────────────────

class TestBuildContextHeader:
    def test_workbook_name_in_header(self, workbook):
        ctx = build_context(workbook)
        assert "=== WORKBOOK: book.xlsx ===" in ctx

    def test_sheet_count_in_header(self, workbook):
        ctx = build_context(workbook)
        assert "SHEETS (1):" in ctx

    def test_vba_modules_listed(self, workbook):
        ctx = build_context(workbook)
        assert "VBA MODULES (1):" in ctx
        assert "Module1" in ctx

    def test_named_ranges_listed(self, workbook):
        ctx = build_context(workbook)
        assert "NAMED RANGES (1):" in ctx
        assert "MyRange" in ctx
        assert "=Sheet1!$A$1:$D$10" in ctx

    def test_extraction_error_shown(self):
        wb = _make_wb(extraction_error="VBA blocked")
        ctx = build_context(wb)
        assert "[NOTE] VBA blocked" in ctx

    def test_no_error_when_none(self, workbook):
        ctx = build_context(workbook)
        assert "[NOTE]" not in ctx

    def test_no_vba_section_when_empty(self):
        wb = _make_wb()
        ctx = build_context(wb)
        assert "VBA MODULES" not in ctx

    def test_no_named_ranges_when_empty(self):
        wb = _make_wb()
        ctx = build_context(wb)
        assert "NAMED RANGES" not in ctx


# ── build_context: cell data ──────────────────────────────────────────────────

class TestBuildContextCells:
    def test_empty_sheet_shows_empty_marker(self):
        wb = _make_wb(sheets=[_make_sheet("Sheet1")])
        ctx = build_context(wb)
        assert "[empty]" in ctx

    def test_string_value_quoted(self):
        cell = CellData(row=1, col=1, address="A1", value="hello")
        s = _make_sheet("Sheet1", {"A1": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert 'R1: A="hello"' in ctx

    def test_numeric_value_unquoted(self):
        cell = CellData(row=1, col=1, address="A1", value=42)
        s = _make_sheet("Sheet1", {"A1": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "R1: A=42" in ctx

    def test_formula_shown_with_braces(self):
        cell = CellData(row=1, col=1, address="A1", value=0, formula="=SUM(B:B)")
        s = _make_sheet("Sheet1", {"A1": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        # leading '=' stripped inside braces in new compact format
        assert "A={SUM(B:B)}" in ctx

    def test_long_string_truncated(self):
        long_val = "x" * 100
        cell = CellData(row=1, col=1, address="A1", value=long_val)
        s = _make_sheet("Sheet1", {"A1": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        # Should be truncated to 80 chars + ellipsis; address is now just column letter
        assert 'A="' + "x" * 80 in ctx

    def test_none_value_rendered(self):
        cell = CellData(row=1, col=1, address="A1", value=None)
        s = _make_sheet("Sheet1", {"A1": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "R1: A=None" in ctx

    def test_sheet_header_present(self):
        wb = _make_wb(sheets=[_make_sheet("MySheet")])
        ctx = build_context(wb)
        assert "--- CELLS: MySheet" in ctx

    def test_multiple_cells_same_row_grouped(self):
        cells = {
            "A1": CellData(row=1, col=1, address="A1", value=1),
            "B1": CellData(row=1, col=2, address="B1", value=2),
        }
        s = _make_sheet("Sheet1", cells)
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        # Both columns should appear on the same R1: line
        lines = ctx.splitlines()
        data_lines = [ln for ln in lines if "R1:" in ln and "A=1" in ln and "B=2" in ln]
        assert len(data_lines) == 1

    def test_truncated_sheet_shows_row_counts(self):
        # used_range_address is shown in the context header
        s = SheetData(
            index=1, name="Big", used_range_address="$A$1:$A$500",
            row_count=500, col_count=1,
        )
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "--- CELLS: Big [$A$1:$A$500] (500r" in ctx


# ── build_context: config filtering ──────────────────────────────────────────

class TestBuildContextFiltering:
    def test_excluded_sheet_shown_as_excluded(self, workbook):
        config = ContextConfig(included_sheets=["OtherSheet"])
        ctx = build_context(workbook, config)
        assert "[excluded by settings]" in ctx

    def test_included_sheet_shown_normally(self, workbook):
        config = ContextConfig(included_sheets=["Sheet1"])
        ctx = build_context(workbook, config)
        assert "[excluded by settings]" not in ctx

    def test_vba_excluded_when_config_off(self, workbook):
        config = ContextConfig(include_vba=False)
        ctx = build_context(workbook, config)
        assert "--- VBA:" not in ctx

    def test_vba_included_when_config_on(self, workbook):
        config = ContextConfig(include_vba=True)
        ctx = build_context(workbook, config)
        assert "--- VBA: Module1" in ctx

    def test_named_ranges_excluded_when_config_off(self, workbook):
        config = ContextConfig(include_named_ranges=False)
        ctx = build_context(workbook, config)
        assert "NAMED RANGES" not in ctx

    def test_specific_vba_module_filter(self, workbook):
        # Add a second module
        workbook.vba_modules.append(
            VBAModule(name="Module2", module_type=1, type_name="Standard", code="Sub X()\nEnd Sub")
        )
        config = ContextConfig(included_vba_modules=["Module1"])
        ctx = build_context(workbook, config)
        assert "--- VBA: Module1" in ctx
        assert "--- VBA: Module2" not in ctx

    def test_none_config_uses_defaults(self, workbook):
        ctx = build_context(workbook, None)
        assert ctx  # Just verify it doesn't crash


# ── _infer_spill_dims / _spill_range ──────────────────────────────────────────

class TestInferSpillDims:
    def test_sequence_rows_cols(self):
        assert _infer_spill_dims("=SEQUENCE(5,3,1,1)") == (5, 3)

    def test_sequence_rows_only(self):
        assert _infer_spill_dims("=SEQUENCE(7)") == (7, 1)

    def test_sequence_with_xlfn_prefix(self):
        assert _infer_spill_dims("=_xlfn.SEQUENCE(5,4)") == (5, 4)

    def test_sequence_lowercase(self):
        assert _infer_spill_dims("=sequence(2, 3)") == (2, 3)

    def test_sequence_non_numeric_skipped(self):
        # Non-literal first arg cannot be statically evaluated.
        assert _infer_spill_dims("=SEQUENCE(A1, 3)") is None

    def test_randarray_dims(self):
        assert _infer_spill_dims("=RANDARRAY(4, 5)") == (4, 5)

    def test_randarray_no_args(self):
        assert _infer_spill_dims("=RANDARRAY()") == (1, 1)

    def test_unique_from_range(self):
        assert _infer_spill_dims("=UNIQUE(A2:A17)") == (16, 1)

    def test_sort_from_range(self):
        assert _infer_spill_dims("=SORT(B2:D10)") == (9, 3)

    def test_sortby_from_first_range(self):
        assert _infer_spill_dims("=SORTBY(A2:C10, C2:C10, -1)") == (9, 3)

    def test_filter_from_range(self):
        assert _infer_spill_dims('=FILTER(A2:D8, A2:A8="Fruit", "no match")') == (7, 4)

    def test_xlookup_not_inferred(self):
        # XLOOKUP's spill shape depends on whether lookup_value is a scalar
        # or a range, which we cannot determine statically, so we deliberately
        # do NOT infer a spill rectangle for it.
        assert _infer_spill_dims('=XLOOKUP("a", B2:B10, D2:F10)') is None
        assert _infer_spill_dims('=XLOOKUP("a", B2:B10)') is None

    def test_plain_sum_not_spill(self):
        assert _infer_spill_dims("=SUM(A1:A10)") is None

    def test_non_formula(self):
        assert _infer_spill_dims("hello") is None
        assert _infer_spill_dims("") is None

    def test_lambda_helpers_skipped(self):
        # MAP/BYROW/etc. depend on lambda evaluation, so we don't infer them.
        assert _infer_spill_dims("=MAP(A1:A10, LAMBDA(x, x*2))") is None
        assert _infer_spill_dims("=BYROW(A1:C10, LAMBDA(r, SUM(r)))") is None


class TestSpillRange:
    def test_multi_cell_range(self):
        # Q2 + (5 rows, 3 cols) -> Q2:S6
        # col Q = 17, +3 cols = 17, 18, 19 -> Q, R, S
        assert _spill_range(17, 2, (5, 3)) == "Q2:S6"

    def test_single_cell_returns_address(self):
        assert _spill_range(1, 1, (1, 1)) == "A1"

    def test_single_column(self):
        assert _spill_range(1, 1, (10, 1)) == "A1:A10"

    def test_single_row(self):
        assert _spill_range(1, 1, (1, 5)) == "A1:E1"


class TestSpillAnnotationInContext:
    def test_sequence_anchor_emits_spill_flag(self):
        cell = CellData(
            row=2, col=17, address="Q2", value=0,
            formula="=_xlfn.SEQUENCE(5,3,1,1)",
        )
        s = _make_sheet("Sheet1", {"Q2": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "[spill:Q2:S6]" in ctx

    def test_unique_anchor_emits_spill_flag(self):
        cell = CellData(
            row=2, col=13, address="M2", value=0,
            formula="=_xlfn.UNIQUE(A2:A17)",
        )
        s = _make_sheet("Sheet1", {"M2": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "[spill:M2:M17]" in ctx

    def test_plain_formula_has_no_spill_flag(self):
        cell = CellData(row=1, col=1, address="A1", value=0, formula="=SUM(B:B)")
        s = _make_sheet("Sheet1", {"A1": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "[spill:" not in ctx

    def test_xlookup_no_spill_flag(self):
        # XLOOKUP is deliberately not annotated -- its spill shape depends on
        # whether lookup_value is a scalar or a range.
        cell = CellData(
            row=2, col=11, address="K2", value=0,
            formula='=_xlfn.XLOOKUP("Apple", B2:B7, D2:D7, "Not found")',
        )
        s = _make_sheet("Sheet1", {"K2": cell})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert "[spill:" not in ctx


    def test_excluded_area_hides_cells_in_range(self):
        # Sheet with cells at A1 (r1,c1), C3 (r3,c3), and E5 (r5,c5).
        # Exclude the area $A$1:$C$3 — A1 and C3 should disappear; E5 remains.
        cells = {
            "A1": CellData(row=1, col=1, address="A1", value="in-area"),
            "C3": CellData(row=3, col=3, address="C3", value="in-area"),
            "E5": CellData(row=5, col=5, address="E5", value="outside"),
        }
        s = SheetData(index=1, name="Sheet1",
                      used_range_address="$A$1:$C$3,$E$5:$E$5",
                      row_count=5, col_count=5, cells=cells)
        config = ContextConfig(excluded_areas={"Sheet1": ["$A$1:$C$3"]})
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb, config)
        assert "A1=" not in ctx
        assert "C3=" not in ctx
        assert 'R5: E="outside"' in ctx

    def test_parse_area_addr_valid(self):
        assert _parse_area_addr("$A$1:$C$10") == (1, 1, 10, 3)
        assert _parse_area_addr("$E$2:$G$5") == (2, 5, 5, 7)
        assert _parse_area_addr("$AA$1:$AB$10") == (1, 27, 10, 28)

    def test_parse_area_addr_invalid(self):
        assert _parse_area_addr("not-an-address") is None
        assert _parse_area_addr("") is None


# ── build_context: hard cap ───────────────────────────────────────────────────

class TestBuildContextHardCap:
    def test_truncation_marker_when_over_limit(self):
        # Create a sheet with enough cells to exceed 60k chars.
        # Each output line is ~92 chars; need >652 rows to exceed 60k.
        cells = {}
        for r in range(1, 5000):
            addr = f"A{r}"
            cells[addr] = CellData(row=r, col=1, address=addr, value="x" * 200)
        s = SheetData(index=1, name="Big", used_range_address="A1:A4999",
                      row_count=4999, col_count=1, cells=cells)
        wb = _make_wb(sheets=[s])
        ctx = build_context(wb)
        assert len(ctx) <= _DEFAULT_MAX_CONTEXT_CHARS + 200  # small slack for the truncation message
        assert "CONTEXT TRUNCATED" in ctx

    def test_no_truncation_for_small_workbook(self, workbook):
        ctx = build_context(workbook)
        assert "CONTEXT TRUNCATED" not in ctx
