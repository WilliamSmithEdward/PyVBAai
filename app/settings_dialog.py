# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Settings dialog - API key status, model, context, backups, appearance."""
from __future__ import annotations

import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from app.config import get_settings
from core.ai_client import AIClient
from models.workbook import ALL_FMT_FIELDS, ContextConfig


def _qbool(val) -> bool:
    """QSettings on Windows stores booleans as strings 'true'/'false'."""
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val.lower() not in ("false", "0", "no", "")
    return bool(val)


class SettingsDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(480)
        self.setModal(True)
        self._settings = get_settings()
        self._setup_ui()
        self._load_settings()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setSpacing(12)
        root.setContentsMargins(16, 16, 16, 16)

        tabs = QTabWidget()

        tabs.addTab(self._build_ai_tab(),       "AI")
        tabs.addTab(self._build_context_tab(),  "Context")
        tabs.addTab(self._build_backup_tab(),   "Backups")
        tabs.addTab(self._build_appear_tab(),   "Appearance")

        root.addWidget(tabs)

        # ── Buttons ──────────────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("secondaryBtn")
        cancel_btn.clicked.connect(self.reject)

        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self._save)

        btn_row.addStretch()
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(save_btn)
        root.addLayout(btn_row)

    # ── AI tab ────────────────────────────────────────────────────────────────
    def _build_ai_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        # API key status
        api_key = os.environ.get("OPENAI_API_KEY", "").strip()
        if api_key:
            status_txt = "OPENAI_API_KEY is set"
            status_color = "#a6e3a1"
        else:
            status_txt = "OPENAI_API_KEY not found in environment"
            status_color = "#f38ba8"

        key_lbl = QLabel(status_txt)
        key_lbl.setStyleSheet(f"color: {status_color}; font-weight: 600;")
        layout.addWidget(key_lbl)

        hint = QLabel(
            "Set the <b>OPENAI_API_KEY</b> user environment variable in Windows Settings "
            "> System > About > Advanced system settings > Environment Variables, "
            "then restart PyVBAai."
        )
        hint.setWordWrap(True)
        hint.setTextFormat(Qt.TextFormat.RichText)
        hint.setStyleSheet("color: #6c7086; font-size: 12px;")
        layout.addWidget(hint)

        layout.addSpacing(12)

        # Model selector
        form = QFormLayout()
        self._model_combo = QComboBox()
        self._model_combo.setPlaceholderText("Select a model...")
        for m in AIClient.fetch_models_from_api():
            self._model_combo.addItem(m)
        form.addRow("Model:", self._model_combo)
        layout.addLayout(form)
        layout.addStretch()
        return w

    # ── Context tab ───────────────────────────────────────────────────────────
    def _build_context_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # -- Inclusions -------------------------------------------------------
        form = QFormLayout()
        self._include_formulas = QCheckBox("Include cell formulas in context")
        self._include_vba = QCheckBox("Include VBA code in context")
        self._include_named = QCheckBox("Include named ranges in context")
        form.addRow(self._include_formulas)
        form.addRow(self._include_vba)
        form.addRow(self._include_named)
        layout.addLayout(form)

        # -- Max rows per area ------------------------------------------------
        rows_group = QGroupBox("Row limit per data area")
        rows_layout = QHBoxLayout(rows_group)
        rows_layout.setSpacing(8)
        self._max_rows_enabled = QCheckBox("Limit to")
        self._max_rows_spin = QSpinBox()
        self._max_rows_spin.setRange(1, 10000)
        self._max_rows_spin.setValue(50)
        self._max_rows_spin.setSuffix(" rows")
        self._max_rows_spin.setFixedWidth(100)
        self._max_rows_enabled.toggled.connect(self._max_rows_spin.setEnabled)
        rows_layout.addWidget(self._max_rows_enabled)
        rows_layout.addWidget(self._max_rows_spin)
        rows_layout.addStretch()
        layout.addWidget(rows_group)

        # -- Format field toggles ---------------------------------------------
        fmt_group = QGroupBox("Cell formatting to include in context")
        fmt_layout = QVBoxLayout(fmt_group)
        fmt_layout.setSpacing(4)
        # ordered label -> key pairs
        self._fmt_checks: dict[str, QCheckBox] = {}
        fmt_fields = [
            ("number_format", "Number format  (e.g. 0.00%, dd/mm/yyyy)"),
            ("bold",          "Bold"),
            ("italic",        "Italic"),
            ("underline",     "Underline"),
            ("font_color",    "Font colour"),
            ("bg_color",      "Background colour"),
            ("h_align",       "Horizontal alignment"),
            ("v_align",       "Vertical alignment"),
            ("wrap_text",     "Wrap text"),
        ]
        row1 = QHBoxLayout()
        row2 = QHBoxLayout()
        for i, (key, label) in enumerate(fmt_fields):
            cb = QCheckBox(label)
            self._fmt_checks[key] = cb
            (row1 if i < 5 else row2).addWidget(cb)
        row1.addStretch()
        row2.addStretch()
        fmt_layout.addLayout(row1)
        fmt_layout.addLayout(row2)
        layout.addWidget(fmt_group)

        note = QLabel(
            "Per-sheet and per-module filtering is available "
            "via the <b>Context Filter</b> button in the workbook explorer sidebar."
        )
        note.setWordWrap(True)
        note.setTextFormat(Qt.TextFormat.RichText)
        note.setStyleSheet("color: #6c7086; font-size: 12px;")
        layout.addWidget(note)

        layout.addStretch()
        return w
    def _build_backup_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        form = QFormLayout()
        self._max_backups_spin = QSpinBox()
        self._max_backups_spin.setRange(1, 200)
        form.addRow("Max backups to keep:", self._max_backups_spin)

        note = QLabel(
            "Backups are stored in a <b>backups/</b> subfolder next to the original file, "
            "named <i>filename_YYYYMMDD_HHMMSS.ext</i>."
        )
        note.setWordWrap(True)
        note.setTextFormat(Qt.TextFormat.RichText)
        note.setStyleSheet("color: #6c7086; font-size: 12px;")

        layout.addLayout(form)
        layout.addWidget(note)
        layout.addStretch()
        return w

    # ── Appearance tab ────────────────────────────────────────────────────────
    def _build_appear_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        self._dark_mode = QCheckBox("Dark mode")
        layout.addWidget(self._dark_mode)
        layout.addStretch()
        return w

    # ── Load / Save ───────────────────────────────────────────────────────────
    def _load_settings(self) -> None:
        s = self._settings
        # AI
        model = s.value("ai/model") or ""
        if model:
            idx = self._model_combo.findText(model)
            if idx < 0:
                # Saved model not in current list (e.g. deprecated) — add it
                self._model_combo.addItem(model)
                idx = self._model_combo.count() - 1
            self._model_combo.setCurrentIndex(idx)
        else:
            self._model_combo.setCurrentIndex(-1)
        # Context
        self._include_formulas.setChecked(_qbool(s.value("context/include_formulas", True)))
        self._include_vba.setChecked(_qbool(s.value("context/include_vba", True)))
        self._include_named.setChecked(_qbool(s.value("context/include_named_ranges", True)))
        max_rows_enabled = _qbool(s.value("context/max_rows_enabled", False))
        self._max_rows_enabled.setChecked(max_rows_enabled)
        self._max_rows_spin.setValue(int(s.value("context/max_rows", 50)))
        self._max_rows_spin.setEnabled(max_rows_enabled)
        for key, cb in self._fmt_checks.items():
            cb.setChecked(_qbool(s.value(f"context/fmt_{key}", True)))

        # Backups
        self._max_backups_spin.setValue(int(s.value("backups/max_keep", 20)))
        # Appearance
        self._dark_mode.setChecked(_qbool(s.value("appearance/dark_mode", True)))

    def _save(self) -> None:
        s = self._settings
        s.setValue("ai/model", self._model_combo.currentText())
        s.setValue("context/include_formulas", self._include_formulas.isChecked())
        s.setValue("context/include_vba", self._include_vba.isChecked())
        s.setValue("context/include_named_ranges", self._include_named.isChecked())
        s.setValue("context/max_rows_enabled", self._max_rows_enabled.isChecked())
        s.setValue("context/max_rows", self._max_rows_spin.value())
        for key, cb in self._fmt_checks.items():
            s.setValue(f"context/fmt_{key}", cb.isChecked())

        s.setValue("backups/max_keep", self._max_backups_spin.value())
        s.setValue("appearance/dark_mode", self._dark_mode.isChecked())

        self.accept()

    # ── Convenience accessors for MainWindow ──────────────────────────────────
    def get_context_config(self) -> ContextConfig:
        return SettingsDialog.load_context_config()

    @staticmethod
    def load_context_config() -> ContextConfig:
        """Load context config from QSettings without opening the dialog."""
        s = get_settings()
        excluded_sheets = list(s.value("context/excluded_sheets", []) or [])
        excluded_vba    = list(s.value("context/excluded_vba",    []) or [])
        excluded_area_pairs = list(s.value("context/excluded_areas", []) or [])
        excluded_areas: dict[str, list[str]] = {}
        for pair in excluded_area_pairs:
            if "||" in pair:
                sname, aaddr = pair.split("||", 1)
                excluded_areas.setdefault(sname, []).append(aaddr)
        max_rows: int | None = None
        if _qbool(s.value("context/max_rows_enabled", False)):
            max_rows = int(s.value("context/max_rows", 50))
        fmt_include: set[str] = {
            key for key in ALL_FMT_FIELDS
            if _qbool(s.value(f"context/fmt_{key}", True))
        }
        return ContextConfig(
            include_formulas=_qbool(s.value("context/include_formulas", True)),
            include_vba=_qbool(s.value("context/include_vba", True)),
            include_named_ranges=_qbool(s.value("context/include_named_ranges", True)),
            excluded_sheets=excluded_sheets,
            excluded_vba_modules=excluded_vba,
            excluded_areas=excluded_areas,
            max_rows_per_area=max_rows,
            fmt_include=fmt_include,
        )

    @staticmethod
    def load_model() -> str:
        return get_settings().value("ai/model") or ""

    @staticmethod
    def load_max_backups() -> int:
        return int(get_settings().value("backups/max_keep", 20))

    @staticmethod
    def load_dark_mode() -> bool:
        return _qbool(get_settings().value("appearance/dark_mode", True))
