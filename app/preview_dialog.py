# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Preview dialog - shows AI-proposed changes before the user approves."""
from __future__ import annotations

import difflib

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from models.conversation import AIResponse, Change

# ── icons / labels per operation type ─────────────────────────────────────────
_OP_META: dict[str, tuple[str, str]] = {
    "set_cell":           ("", "Set Cell"),
    "set_range":          ("", "Set Range"),
    "set_format":         ("", "Format Range"),
    "clear_range":        ("", "Clear Range"),
    "add_sheet":          ("", "Add Sheet"),
    "delete_sheet":       ("", "Delete Sheet"),
    "rename_sheet":       ("", "Rename Sheet"),
    "move_sheet":         ("", "Move Sheet"),
    "copy_sheet":         ("", "Copy Sheet"),
    "hide_sheet":         ("", "Hide Sheet"),
    "unhide_sheet":       ("", "Unhide Sheet"),
    "merge_cells":        ("", "Merge Cells"),
    "unmerge_cells":      ("", "Unmerge Cells"),
    "set_vba":            ("", "Update VBA Module"),
    "add_vba_module":     ("", "Add VBA Module"),
    "delete_vba_module":  ("", "Delete VBA Module"),
    "set_named_range":    ("", "Set Named Range"),
    "add_named_range":    ("", "Add Named Range"),
    "delete_named_range": ("", "Delete Named Range"),
}


class ChangeCard(QFrame):
    """One card per change operation in the preview list."""

    def __init__(self, change: Change, old_vba: str | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(4)

        icon, label = _OP_META.get(change.type, ("", change.type))
        p = change.params

        # Title row
        title = QLabel(f"<b>{label}</b>" if not icon else f"{icon}  <b>{label}</b>")
        title.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(title)

        # Details
        details = _format_details(change.type, p)
        if details:
            det_lbl = QLabel(details)
            det_lbl.setWordWrap(True)
            det_lbl.setStyleSheet("color: #6c7086; font-size: 12px;")
            layout.addWidget(det_lbl)

        # VBA diff (collapsible)
        if change.type in ("set_vba", "add_vba_module") and p.get("code"):
            diff_html = _make_vba_diff(old_vba or "", p["code"])
            btn = QPushButton("Show code diff")
            btn.setObjectName("ghostBtn")
            btn.setFixedWidth(130)
            diff_view = QTextBrowser()
            diff_view.setObjectName("codeViewer")
            diff_view.setHtml(diff_html)
            diff_view.setMinimumHeight(160)
            diff_view.setVisible(False)

            def _toggle(checked=False, dv=diff_view, b=btn):
                dv.setVisible(not dv.isVisible())
                b.setText("Hide code diff" if dv.isVisible() else "Show code diff")

            btn.clicked.connect(_toggle)
            layout.addWidget(btn)
            layout.addWidget(diff_view)


def _format_details(op: str, p: dict) -> str:  # noqa: PLR0911
    if op == "set_cell":
        loc = f"{p.get('sheet','')}!{p.get('cell','')}"
        val = p.get("formula") or repr(p.get("value", ""))
        return f"Location: {loc}  →  {val}"
    if op == "set_range":
        rows = len(p.get("values", []))
        cols = len(p.get("values", [[]])[0]) if p.get("values") else 0
        return f"Sheet: {p.get('sheet','')}   Range: {p.get('range','')}   ({rows}×{cols} values)"
    if op == "set_format":
        rng = f"{p.get('sheet','')}!{p.get('range','')}"
        keys = [k for k in p if k not in ("type", "sheet", "range")]
        return f"Range: {rng}   Fields: {', '.join(keys)}"
    if op == "clear_range":
        return f"Sheet: {p.get('sheet','')}   Range: {p.get('range','')}"
    if op == "add_sheet":
        return f"Name: {p.get('name','')}   position: {p.get('position', 'end')}"
    if op == "delete_sheet":
        return f"Sheet: {p.get('name','')}"
    if op == "rename_sheet":
        return f"{p.get('old_name','')}  →  {p.get('new_name','')}"
    if op == "move_sheet":
        return f"Sheet: {p.get('name','')}   →  position {p.get('position','')}"
    if op == "copy_sheet":
        return f"'{p.get('source','')}' → '{p.get('dest','')}' at position {p.get('position', 'end')}"
    if op in ("hide_sheet", "unhide_sheet"):
        return f"Sheet: {p.get('name','')}"
    if op in ("merge_cells", "unmerge_cells"):
        return f"Sheet: {p.get('sheet','')}   Range: {p.get('range','')}"
    if op in ("set_vba", "add_vba_module"):
        lines = len((p.get("code") or "").splitlines())
        return f"Module: {p.get('module') or p.get('name','')}   ({lines} lines)"
    if op == "delete_vba_module":
        return f"Module: {p.get('name','')}"
    if op in ("set_named_range", "add_named_range"):
        return f"{p.get('name','')}  =  {p.get('refers_to','')}"
    if op == "delete_named_range":
        return f"Name: {p.get('name','')}"
    return ""


def _make_vba_diff(old: str, new: str) -> str:
    old_lines = old.splitlines(keepends=True)
    new_lines = new.splitlines(keepends=True)
    diff = list(difflib.unified_diff(old_lines, new_lines, fromfile="before", tofile="after", lineterm=""))

    html_lines = []
    for line in diff:
        esc = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if line.startswith("+") and not line.startswith("+++"):
            html_lines.append(f'<span style="color:#a6e3a1;">{esc}</span>')
        elif line.startswith("-") and not line.startswith("---"):
            html_lines.append(f'<span style="color:#f38ba8;">{esc}</span>')
        elif line.startswith("@@"):
            html_lines.append(f'<span style="color:#89b4fa;">{esc}</span>')
        else:
            html_lines.append(f'<span style="color:#6c7086;">{esc}</span>')

    body = "\n".join(html_lines) or "<span style='color:#6c7086;'>No diff (new module)</span>"
    return (
        "<pre style='font-family:\"Cascadia Mono\",Consolas,monospace;"
        "font-size:12px;background:#1e1e2e;padding:8px;border-radius:4px;'>"
        f"{body}</pre>"
    )


# ── Preview Dialog ─────────────────────────────────────────────────────────────

class PreviewDialog(QDialog):
    """Shows the AI response summary + all planned changes."""

    approved = Signal()
    declined = Signal()
    revise_requested = Signal(str)  # user's revision note

    def __init__(
        self,
        ai_response: AIResponse,
        vba_map: dict[str, str] | None = None,  # module_name -> current code
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Preview Changes")
        self.setMinimumSize(620, 480)
        self.resize(740, 580)
        self.setModal(True)
        self._response = ai_response
        self._vba_map = vba_map or {}
        self._setup_ui()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setSpacing(12)
        root.setContentsMargins(16, 16, 16, 16)

        # ── Summary ───────────────────────────────────────────────────────────
        summary_label = QLabel("AI Summary")
        summary_label.setStyleSheet("font-weight: 600; font-size: 14px;")
        root.addWidget(summary_label)

        summary_text = QTextBrowser()
        summary_text.setMaximumHeight(100)
        import markdown as md
        summary_text.setHtml(md.markdown(self._response.message))
        root.addWidget(summary_text)

        # ── Changes list ──────────────────────────────────────────────────────
        change_count = len(self._response.changes)
        changes_label = QLabel(
            f"Changes ({change_count})"
            if change_count
            else "No file modifications proposed"
        )
        changes_label.setStyleSheet("font-weight: 600; font-size: 14px;")
        root.addWidget(changes_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        cards_container = QWidget()
        cards_layout = QVBoxLayout(cards_container)
        cards_layout.setSpacing(8)
        cards_layout.setContentsMargins(0, 0, 4, 0)

        for change in self._response.changes:
            module_name = change.params.get("module") or change.params.get("name", "")
            old_code = self._vba_map.get(module_name, "")
            card = ChangeCard(change, old_vba=old_code)
            cards_layout.addWidget(card)

        if not self._response.changes:
            cards_layout.addWidget(QLabel("(nothing to apply)"))

        cards_layout.addStretch()
        scroll.setWidget(cards_container)
        root.addWidget(scroll)

        # ── Button row ────────────────────────────────────────────────────────
        btn_row = QHBoxLayout()

        revise_btn = QPushButton("Revise...")
        revise_btn.setObjectName("ghostBtn")
        revise_btn.clicked.connect(self._on_revise)

        decline_btn = QPushButton("Decline")
        decline_btn.setObjectName("secondaryBtn")
        decline_btn.clicked.connect(self._on_decline)

        apply_btn = QPushButton("Apply Changes")
        if not self._response.changes:
            apply_btn.setEnabled(False)
        apply_btn.clicked.connect(self._on_apply)

        btn_row.addWidget(revise_btn)
        btn_row.addStretch()
        btn_row.addWidget(decline_btn)
        btn_row.addWidget(apply_btn)
        root.addLayout(btn_row)

    def _on_apply(self) -> None:
        self.approved.emit()
        self.accept()

    def _on_decline(self) -> None:
        self.declined.emit()
        self.reject()

    def _on_revise(self) -> None:
        from PySide6.QtWidgets import QInputDialog
        note, ok = QInputDialog.getText(
            self,
            "Revise Request",
            "Describe what to change about the proposed plan:",
        )
        if ok and note.strip():
            self.revise_requested.emit(note.strip())
            self.reject()
