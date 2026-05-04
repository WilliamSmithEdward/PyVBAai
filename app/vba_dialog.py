# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""VBA copy/paste dialog.

Shows each VBA change as a read-only code block the user can copy and paste
into the Excel VBA editor (Alt+F11). No COM or file mutation is performed.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
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
    "set_vba": "Replace module code",
    "add_vba_module": "Add new module",
    "delete_vba_module": "Delete module",
}


class VBADialog(QDialog):
    """Displays VBA changes as copy/paste blocks for the user."""

    def __init__(self, changes: list[Change], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("VBA Changes — Copy & Paste into Excel")
        self.setMinimumSize(680, 480)
        self.resize(780, 600)
        self.setModal(True)
        self._changes = changes
        self._setup_ui()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setSpacing(12)
        root.setContentsMargins(16, 16, 16, 16)

        info = QLabel(
            "<b>VBA changes cannot be applied automatically.</b><br>"
            "Open the VBA editor in Excel (<b>Alt+F11</b>), then copy and paste "
            "each module below."
        )
        info.setWordWrap(True)
        info.setTextFormat(Qt.TextFormat.RichText)
        root.addWidget(info)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(scroll.Shape.NoFrame)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(16)
        container_layout.setContentsMargins(0, 0, 0, 0)

        for change in self._changes:
            container_layout.addWidget(self._make_card(change))

        container_layout.addStretch()
        scroll.setWidget(container)
        root.addWidget(scroll, stretch=1)

        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        root.addLayout(btn_row)

    def _make_card(self, change: Change) -> QWidget:
        p = change.params
        op_label = _OP_LABEL.get(change.type, change.type)

        if change.type == "delete_vba_module":
            module_name = p.get("name", "")
            card = QWidget()
            layout = QVBoxLayout(card)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(4)
            header = QLabel(f"<b>{op_label}</b>: <code>{module_name}</code>")
            header.setTextFormat(Qt.TextFormat.RichText)
            layout.addWidget(header)
            note = QLabel(
                f"In the VBA editor, right-click <b>{module_name}</b> "
                "in the Project pane and choose <b>Remove</b>."
            )
            note.setTextFormat(Qt.TextFormat.RichText)
            note.setWordWrap(True)
            layout.addWidget(note)
            return card

        module_name = p.get("module") or p.get("name", "")
        code: str = p.get("code", "")

        card = QWidget()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        header_row = QHBoxLayout()
        header_lbl = QLabel(f"<b>{op_label}</b>: <code>{module_name}</code>")
        header_lbl.setTextFormat(Qt.TextFormat.RichText)
        header_row.addWidget(header_lbl, stretch=1)

        copy_btn = QPushButton("Copy")
        copy_btn.setFixedWidth(70)
        copy_btn.setObjectName("secondaryBtn")
        copy_btn.clicked.connect(lambda checked=False, c=code: self._copy(c))
        header_row.addWidget(copy_btn)
        layout.addLayout(header_row)

        editor = QTextEdit()
        editor.setReadOnly(True)
        editor.setPlainText(code)
        mono = QFont("Cascadia Mono, Consolas, Courier New")
        mono.setStyleHint(QFont.StyleHint.Monospace)
        editor.setFont(mono)
        editor.setObjectName("codeViewer")
        editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        lines = max(6, min(30, code.count("\n") + 2))
        editor.setFixedHeight(lines * 18 + 12)
        layout.addWidget(editor)

        return card

    @staticmethod
    def _copy(text: str) -> None:
        QApplication.clipboard().setText(text)
