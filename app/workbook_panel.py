# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Left-sidebar panel showing the loaded workbook's structure."""
from __future__ import annotations

from PySide6.QtCore import QRect, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QPainter, QPainterPath
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.config import get_settings
from models.workbook import WorkbookData


def _sync_tristate(item: QTreeWidgetItem) -> None:
    """Set *item*'s check state from its direct children's states.

    Best-practice tristate rules:
      - All Checked        → Checked
      - All Unchecked      → Unchecked
      - Mixed / any Partial → PartiallyChecked
    Called explicitly during setup (blockSignals active) so Qt's own
    auto-tristate propagation is not relied upon.
    """
    n = item.childCount()
    if n == 0:
        return
    n_checked = n_unchecked = 0
    for i in range(n):
        cs = item.child(i).checkState(0)
        if cs == Qt.CheckState.Checked:
            n_checked += 1
        elif cs == Qt.CheckState.Unchecked:
            n_unchecked += 1
    if n_checked == n:
        item.setCheckState(0, Qt.CheckState.Checked)
    elif n_unchecked == n:
        item.setCheckState(0, Qt.CheckState.Unchecked)
    else:
        item.setCheckState(0, Qt.CheckState.PartiallyChecked)


class _BranchTree(QTreeWidget):
    """QTreeWidget with QPainter-drawn branch arrows so they're visible on any theme."""

    _accent: str = "#89b4fa"

    def set_accent(self, color: str) -> None:
        self._accent = color
        self.viewport().update()

    def drawBranches(  # type: ignore[override]
        self, painter: QPainter, rect: QRect, index
    ) -> None:
        item = self.itemFromIndex(index)
        if item is None or item.childCount() == 0:
            return
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(QColor(self._accent)))
        cx = rect.right() - 10
        cy = rect.center().y()
        path = QPainterPath()
        if self.isExpanded(index):
            path.moveTo(cx - 5, cy - 3)
            path.lineTo(cx + 5, cy - 3)
            path.lineTo(cx, cy + 4)
        else:
            path.moveTo(cx - 3, cy - 5)
            path.lineTo(cx + 4, cy)
            path.lineTo(cx - 3, cy + 5)
        path.closeSubpath()
        painter.drawPath(path)
        painter.restore()


class WorkbookPanel(QWidget):
    """Sidebar showing sheets, VBA modules, and named ranges."""

    vba_view_requested = Signal(str, str)  # (module_name, code)
    context_changed = Signal()             # fired when any include/exclude checkbox changes

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("workbookPanel")
        self.setMinimumWidth(200)
        self.setMaximumWidth(320)
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── Header ────────────────────────────────────────────────────────────
        self._header = QLabel("No file loaded")
        self._header.setObjectName("sidebarHeader")
        self._header.setWordWrap(True)
        layout.addWidget(self._header)

        self._tree = _BranchTree()
        self._tree.setHeaderHidden(True)
        self._tree.setAnimated(True)
        self._tree.itemDoubleClicked.connect(self._on_double_click)
        self._tree.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self._tree)

    # ── Public API ─────────────────────────────────────────────────────────────

    def set_accent(self, color: str) -> None:
        """Update branch arrow color to match the current theme's accent."""
        self._tree.set_accent(color)

    def set_workbook(self, wb: WorkbookData) -> None:
        self._tree.blockSignals(True)
        self._header.setText(wb.name)
        self._tree.clear()
        self._wb = wb

        s = get_settings()
        excluded_sheets = set(s.value("context/excluded_sheets", []) or [])
        excluded_vba    = set(s.value("context/excluded_vba",    []) or [])

        # Parse excluded areas: ["SheetName||$A$1:$C$10", ...]
        excluded_area_pairs = list(s.value("context/excluded_areas", []) or [])
        excl_areas_by_sheet: dict[str, set[str]] = {}
        for pair in excluded_area_pairs:
            if "||" in pair:
                sname, aaddr = pair.split("||", 1)
                excl_areas_by_sheet.setdefault(sname, set()).add(aaddr)

        # ── Sheets root (3-state) ───────────────────────────────────────────────────────
        sheets_root = QTreeWidgetItem(["Sheets"])
        sheets_root.setFlags(
            sheets_root.flags()
            | Qt.ItemFlag.ItemIsUserCheckable
            | Qt.ItemFlag.ItemIsAutoTristate
        )
        sheets_root.setExpanded(True)

        for sheet in wb.sheets:
            vis  = "" if sheet.visible else " [hidden]"
            info = f"  {sheet.row_count}r × {sheet.col_count}c"
            item = QTreeWidgetItem([f"{sheet.name}{vis}{info}"])
            item.setData(0, Qt.ItemDataRole.UserRole, ("sheet", sheet.name))
            item.setFlags(
                item.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsAutoTristate
            )

            # Show each used area as a checkable child item.
            # UsedRange may have multiple areas for non-contiguous data.
            areas = [a.strip() for a in sheet.used_range_address.split(",") if a.strip()]
            for area_addr in areas:
                area_excluded = (
                    sheet.name in excluded_sheets
                    or area_addr in excl_areas_by_sheet.get(sheet.name, set())
                )
                addr_item = QTreeWidgetItem([area_addr])
                addr_item.setData(0, Qt.ItemDataRole.UserRole, ("area", sheet.name, area_addr))
                addr_item.setFlags(addr_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.addChild(addr_item)
                addr_item.setCheckState(
                    0,
                    Qt.CheckState.Unchecked if area_excluded else Qt.CheckState.Checked,
                )

            if areas:
                # Derive sheet item state from area children
                _sync_tristate(item)
            else:
                # No area sub-items — simple 2-state checkbox
                item.setCheckState(
                    0,
                    Qt.CheckState.Unchecked if sheet.name in excluded_sheets
                    else Qt.CheckState.Checked,
                )

            sheets_root.addChild(item)

        _sync_tristate(sheets_root)
        self._tree.addTopLevelItem(sheets_root)

        # ── VBA Modules root (3-state) ──────────────────────────────────────────────
        if wb.vba_modules:
            vba_root = QTreeWidgetItem(["VBA Modules"])
            vba_root.setFlags(
                vba_root.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsAutoTristate
            )
            vba_root.setExpanded(True)
            for m in wb.vba_modules:
                lines = len(m.code.splitlines()) if m.code else 0
                item = QTreeWidgetItem([f"{m.name}  ({m.type_name}, {lines} lines)"])
                item.setData(0, Qt.ItemDataRole.UserRole, ("vba", m.name, m.code))
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(
                    0,
                    Qt.CheckState.Unchecked if m.name in excluded_vba
                    else Qt.CheckState.Checked,
                )
                item.setToolTip(0, "Double-click to view source code")
                vba_root.addChild(item)
            _sync_tristate(vba_root)
            self._tree.addTopLevelItem(vba_root)

        if wb.named_ranges:
            nr_root = QTreeWidgetItem(["Named Ranges"])
            nr_root.setExpanded(False)
            for nr in wb.named_ranges:
                item = QTreeWidgetItem([f"{nr.name}  =  {nr.refers_to}"])
                item.setData(0, Qt.ItemDataRole.UserRole, ("named_range", nr.name))
                nr_root.addChild(item)
            self._tree.addTopLevelItem(nr_root)

        if wb.extraction_error:
            warn = QTreeWidgetItem([f"Warning: {wb.extraction_error}"])
            self._tree.addTopLevelItem(warn)

        self._tree.blockSignals(False)

    def clear(self) -> None:
        self._tree.clear()
        self._header.setText("No file loaded")

    def _on_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if column != 0:
            return
        data = item.data(0, Qt.ItemDataRole.UserRole)
        if data is None or data[0] not in ("sheet", "vba", "area"):
            return
        self._save_check_state()
        self.context_changed.emit()

    def _save_check_state(self) -> None:
        excluded_sheets: list[str] = []
        excluded_vba:    list[str] = []
        excluded_area_pairs: list[str] = []

        for i in range(self._tree.topLevelItemCount()):
            root = self._tree.topLevelItem(i)
            for j in range(root.childCount()):
                child = root.child(j)
                data = child.data(0, Qt.ItemDataRole.UserRole)
                if data is None:
                    continue
                if data[0] == "sheet":
                    sheet_name = data[1]
                    if child.checkState(0) == Qt.CheckState.Unchecked:
                        excluded_sheets.append(sheet_name)
                    else:
                        # Checked or PartiallyChecked — sheet is included;
                        # record which individual areas are excluded.
                        for k in range(child.childCount()):
                            area_item = child.child(k)
                            area_data = area_item.data(0, Qt.ItemDataRole.UserRole)
                            if area_data and area_data[0] == "area":
                                if area_item.checkState(0) == Qt.CheckState.Unchecked:
                                    excluded_area_pairs.append(
                                        f"{sheet_name}||{area_data[2]}"
                                    )
                elif data[0] == "vba":
                    if child.checkState(0) == Qt.CheckState.Unchecked:
                        excluded_vba.append(data[1])

        s = get_settings()
        s.setValue("context/excluded_sheets", excluded_sheets)
        s.setValue("context/excluded_vba", excluded_vba)
        s.setValue("context/excluded_areas", excluded_area_pairs)

    def _on_double_click(self, item: QTreeWidgetItem, _col: int) -> None:
        data = item.data(0, Qt.ItemDataRole.UserRole)
        if data and data[0] == "vba":
            _, name, code = data
            self.vba_view_requested.emit(name, code)


class VBAViewerDialog(QDialog):
    def __init__(self, module_name: str, code: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"VBA - {module_name}")
        self.setMinimumSize(720, 500)
        self.resize(800, 600)

        layout = QVBoxLayout(self)

        browser = QTextBrowser()
        browser.setObjectName("codeViewer")
        # Wrap code in HTML pre block for monospace display
        escaped = code.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        browser.setHtml(f"<pre>{escaped}</pre>")
        layout.addWidget(browser)

        btn = QPushButton("Close")
        btn.setFixedWidth(80)
        btn.clicked.connect(self.accept)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(btn)
        layout.addLayout(row)
