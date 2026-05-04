# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""VBA copy/paste dialog.

Shows each VBA change as a numbered step with a copy button.
No COM or file mutation is performed - the user pastes into Excel's VBA editor.
"""
from __future__ import annotations

import os
import subprocess

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from models.conversation import Change

_OP_LABEL: dict[str, str] = {
    "set_vba": "Replace the code in module",
    "add_vba_module": "Create a new module called",
    "delete_vba_module": "Delete the module called",
}

_BANNER_STYLE = (
    "background: #92400e; color: #fef3c7; border-radius: 8px; padding: 14px 18px;"
)
_STEP_STYLE = (
    "background: #1e3a5f; border: 2px solid #3b82f6; border-radius: 8px; padding: 14px;"
)
_DELETE_STEP_STYLE = (
    "background: #3b1f1f; border: 2px solid #ef4444; border-radius: 8px; padding: 14px;"
)
_COPY_BTN_STYLE = (
    "QPushButton { background: #16a34a; color: #f0fdf4; font-weight: 700; "
    "font-size: 14px; border-radius: 6px; padding: 8px 20px; }"
    "QPushButton:hover { background: #15803d; }"
    "QPushButton:pressed { background: #166534; }"
)
_COPIED_BTN_STYLE = (
    "QPushButton { background: #14532d; color: #bbf7d0; font-weight: 700; "
    "font-size: 14px; border-radius: 6px; padding: 8px 20px; }"
)


class VBADialog(QDialog):
    """Displays VBA changes as numbered copy/paste steps for the user."""

    def __init__(
        self,
        changes: list[Change],
        parent: QWidget | None = None,
        file_path: str | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("(!!)  Action Required - VBA Changes")
        self.setMinimumSize(720, 540)
        self.resize(820, 680)
        self.setModal(True)
        self._changes = changes
        self._file_path = file_path
        self._setup_ui()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setSpacing(16)
        root.setContentsMargins(20, 20, 20, 20)

        # ── Top banner ────────────────────────────────────────────────────────
        banner_frame = QFrame()
        banner_frame.setStyleSheet(_BANNER_STYLE)
        banner_layout = QVBoxLayout(banner_frame)
        banner_layout.setContentsMargins(0, 0, 0, 0)
        banner_layout.setSpacing(6)

        banner_title = QLabel("(!!)  YOU NEED TO DO SOMETHING!")
        banner_title.setStyleSheet("font-size: 17px; font-weight: 800; color: #fef3c7;")
        banner_layout.addWidget(banner_title)

        banner_body = QLabel(
            "PyVBAai cannot write VBA code directly into Excel for you.\n"
            "Follow the steps below to paste the code in yourself - it only takes a minute!"
        )
        banner_body.setWordWrap(True)
        banner_body.setStyleSheet("font-size: 13px; color: #fde68a;")
        banner_layout.addWidget(banner_body)

        root.addWidget(banner_frame)

        # ── How-to instructions ───────────────────────────────────────────────
        how_to = QLabel(
            "<b>How to open the VBA editor:</b>  "
            "In Excel, press <b>Alt + F11</b> on your keyboard "
            "(or go to <b>Developer → Visual Basic</b>)."
        )
        how_to.setWordWrap(True)
        how_to.setTextFormat(Qt.TextFormat.RichText)
        how_to.setStyleSheet("font-size: 13px; padding: 4px 0;")
        root.addWidget(how_to)

        # ── Scrollable steps ──────────────────────────────────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(scroll.Shape.NoFrame)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(14)
        container_layout.setContentsMargins(0, 0, 0, 0)

        code_changes = [c for c in self._changes if c.type != "delete_vba_module"]
        delete_changes = [c for c in self._changes if c.type == "delete_vba_module"]

        step = 1
        for change in code_changes:
            container_layout.addWidget(self._make_code_step(step, change))
            step += 1
        for change in delete_changes:
            container_layout.addWidget(self._make_delete_step(step, change))
            step += 1

        container_layout.addStretch()
        scroll.setWidget(container)
        root.addWidget(scroll, stretch=1)

        # ── Bottom button row ──────────────────────────────────────────────────
        btn_row = QHBoxLayout()

        if self._file_path and os.path.isfile(self._file_path):
            open_btn = QPushButton("Open in Excel")
            open_btn.setStyleSheet(
                "QPushButton { background: #334155; color: #e2e8f0; font-weight: 600; "
                "font-size: 13px; border-radius: 6px; padding: 10px 20px; "
                "border: 1px solid #475569; }"
                "QPushButton:hover { background: #475569; }"
                "QPushButton:pressed { background: #1e293b; }"
            )
            open_btn.setFixedHeight(44)
            open_btn.clicked.connect(self._open_in_excel)
            btn_row.addWidget(open_btn)

        btn_row.addStretch()

        close_btn = QPushButton("[OK]  Done - Close this window")
        close_btn.setStyleSheet(
            "QPushButton { background: #3b82f6; color: white; font-weight: 700; "
            "font-size: 14px; border-radius: 6px; padding: 10px 24px; }"
            "QPushButton:hover { background: #2563eb; }"
        )
        close_btn.setFixedHeight(44)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        root.addLayout(btn_row)

    def _open_in_excel(self) -> None:
        if self._file_path:
            os.startfile(os.path.normpath(self._file_path))  # noqa: S606

    def _make_code_step(self, step: int, change: Change) -> QFrame:
        p = change.params
        module_name = p.get("module") or p.get("name", "")
        op_label = _OP_LABEL.get(change.type, change.type)
        code: str = p.get("code", "")

        frame = QFrame()
        frame.setStyleSheet(_STEP_STYLE)
        layout = QVBoxLayout(frame)
        layout.setSpacing(10)
        layout.setContentsMargins(0, 0, 0, 0)

        # Step header
        header_row = QHBoxLayout()
        step_badge = QLabel(f"STEP {step}")
        step_badge.setStyleSheet(
            "background: #3b82f6; color: white; font-weight: 800; font-size: 12px; "
            "border-radius: 4px; padding: 2px 8px;"
        )
        step_badge.setFixedHeight(22)
        header_row.addWidget(step_badge)
        header_row.addSpacing(8)

        title = QLabel(f"{op_label} <b>{module_name}</b>")
        title.setTextFormat(Qt.TextFormat.RichText)
        title.setStyleSheet("font-size: 14px; font-weight: 600;")
        header_row.addWidget(title, stretch=1)
        layout.addLayout(header_row)

        # Numbered instructions
        if change.type == "set_vba":
            instructions = (
                f"1. In the VBA editor, find <b>{module_name}</b> in the left panel<br>"
                f"2. Click on it to open it<br>"
                f"3. Press <b>Ctrl+A</b> to select all existing code<br>"
                f"4. Click the <b>[Copy] Copy Code</b> button below, then press <b>Ctrl+V</b> to paste"
            )
        else:
            instructions = (
                f"1. In the VBA editor, right-click the project name in the left panel<br>"
                f"2. Choose <b>Insert → Module</b><br>"
                f"3. Rename it to <b>{module_name}</b> in the Properties panel (press F4)<br>"
                f"4. Click the <b>[Copy] Copy Code</b> button below, then press <b>Ctrl+V</b> to paste"
            )

        instr_lbl = QLabel(instructions)
        instr_lbl.setTextFormat(Qt.TextFormat.RichText)
        instr_lbl.setWordWrap(True)
        instr_lbl.setStyleSheet("font-size: 12px; line-height: 1.6;")
        layout.addWidget(instr_lbl)

        # Copy button row
        copy_btn = QPushButton("[Copy]  Copy Code")
        copy_btn.setStyleSheet(_COPY_BTN_STYLE)
        copy_btn.setFixedHeight(40)
        copy_btn.setFixedWidth(160)
        copy_btn.clicked.connect(lambda checked=False, b=copy_btn, c=code: self._copy(b, c))
        copy_row = QHBoxLayout()
        copy_row.addWidget(copy_btn)
        copy_row.addStretch()
        layout.addLayout(copy_row)

        # Code preview (collapsible)
        toggle_btn = QPushButton(">  Show code preview")
        toggle_btn.setObjectName("ghostBtn")
        toggle_btn.setFixedWidth(180)

        editor = QTextEdit()
        editor.setReadOnly(True)
        editor.setPlainText(code)
        mono = QFont()
        mono.setFamilies(["Cascadia Mono", "Consolas", "Courier New"])
        mono.setStyleHint(QFont.StyleHint.Monospace)
        mono.setPointSize(10)
        editor.setFont(mono)
        editor.setObjectName("codeViewer")
        editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        lines = max(6, min(25, code.count("\n") + 2))
        editor.setFixedHeight(lines * 18 + 12)
        editor.setVisible(False)

        def _toggle(checked=False, dv=editor, b=toggle_btn):
            dv.setVisible(not dv.isVisible())
            b.setText("v  Hide code preview" if dv.isVisible() else ">  Show code preview")

        toggle_btn.clicked.connect(_toggle)
        layout.addWidget(toggle_btn)
        layout.addWidget(editor)

        return frame

    def _make_delete_step(self, step: int, change: Change) -> QFrame:
        module_name = change.params.get("name", "")

        frame = QFrame()
        frame.setStyleSheet(_DELETE_STEP_STYLE)
        layout = QVBoxLayout(frame)
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)

        header_row = QHBoxLayout()
        step_badge = QLabel(f"STEP {step}")
        step_badge.setStyleSheet(
            "background: #ef4444; color: white; font-weight: 800; font-size: 12px; "
            "border-radius: 4px; padding: 2px 8px;"
        )
        step_badge.setFixedHeight(22)
        header_row.addWidget(step_badge)
        header_row.addSpacing(8)

        title = QLabel(f"Delete module <b>{module_name}</b>")
        title.setTextFormat(Qt.TextFormat.RichText)
        title.setStyleSheet("font-size: 14px; font-weight: 600;")
        header_row.addWidget(title, stretch=1)
        layout.addLayout(header_row)

        note = QLabel(
            f"1. In the VBA editor, find <b>{module_name}</b> in the left panel<br>"
            f"2. Right-click it and choose <b>Remove {module_name}...</b><br>"
            f"3. Click <b>No</b> when asked to export"
        )
        note.setTextFormat(Qt.TextFormat.RichText)
        note.setWordWrap(True)
        note.setStyleSheet("font-size: 12px; line-height: 1.6;")
        layout.addWidget(note)

        return frame

    @staticmethod
    def _copy(btn: QPushButton, text: str) -> None:
        QApplication.clipboard().setText(text)
        btn.setText("[OK]  Copied!")
        btn.setStyleSheet(_COPIED_BTN_STYLE)
        QTimer.singleShot(2500, lambda: (
            btn.setText("[Copy]  Copy Code"),
            btn.setStyleSheet(_COPY_BTN_STYLE),
        ))

