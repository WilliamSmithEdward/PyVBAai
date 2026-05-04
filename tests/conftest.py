# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Shared pytest fixtures and helpers."""
from __future__ import annotations

import pytest

from models.conversation import AIResponse, Change, Conversation
from models.workbook import CellData, ContextConfig, NamedRange, SheetData, VBAModule, WorkbookData

# ── Workbook fixtures ─────────────────────────────────────────────────────────

@pytest.fixture()
def simple_cell() -> CellData:
    return CellData(row=1, col=1, address="A1", value="Hello", formula="")


@pytest.fixture()
def formula_cell() -> CellData:
    return CellData(row=2, col=3, address="C2", value=42, formula="=SUM(A1:A10)")


@pytest.fixture()
def simple_sheet(simple_cell: CellData, formula_cell: CellData) -> SheetData:
    s = SheetData(
        index=1,
        name="Sheet1",
        used_range_address="A1:C2",
        row_count=2,
        col_count=3,
    )
    s.cells["A1"] = simple_cell
    s.cells["C2"] = formula_cell
    return s


@pytest.fixture()
def vba_module() -> VBAModule:
    return VBAModule(
        name="Module1",
        module_type=1,
        type_name="Standard",
        code="Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub",
    )


@pytest.fixture()
def named_range() -> NamedRange:
    return NamedRange(name="MyRange", refers_to="=Sheet1!$A$1:$D$10", scope="Workbook")


@pytest.fixture()
def workbook(simple_sheet: SheetData, vba_module: VBAModule, named_range: NamedRange) -> WorkbookData:
    return WorkbookData(
        file_path="C:/test/book.xlsx",
        name="book.xlsx",
        sheets=[simple_sheet],
        vba_modules=[vba_module],
        named_ranges=[named_range],
        has_vba=True,
    )


@pytest.fixture()
def empty_workbook() -> WorkbookData:
    return WorkbookData(file_path="C:/test/empty.xlsx", name="empty.xlsx")


# ── Conversation fixtures ─────────────────────────────────────────────────────

@pytest.fixture()
def conversation() -> Conversation:
    return Conversation()


@pytest.fixture()
def ai_response_with_changes() -> AIResponse:
    return AIResponse(
        message="I updated cell A1.",
        changes=[
            Change(type="set_cell", params={"sheet": "Sheet1", "cell": "A1", "value": 99}),
        ],
        diff_summary="- Set A1 to 99",
        raw_json='{"message":"I updated cell A1.","changes":[{"type":"set_cell","sheet":"Sheet1","cell":"A1","value":99}],"diff_summary":"- Set A1 to 99"}',
    )


@pytest.fixture()
def default_config() -> ContextConfig:
    return ContextConfig()
