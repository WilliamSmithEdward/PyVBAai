# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""PyVBAai entry point."""
from __future__ import annotations

import os
import sys

# ── PyInstaller / Nuitka: ensure bundled resources are accessible ──────────────
if getattr(sys, "frozen", False):
    # When frozen, add the executable's directory to sys.path
    _bundle_dir = os.path.dirname(sys.executable)
    if _bundle_dir not in sys.path:
        sys.path.insert(0, _bundle_dir)

# Add project root to path so imports work when running from source
_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)


def _set_user_env_var(name: str, value: str) -> None:
    """Write a user-level environment variable to the Windows registry."""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Environment",
            0,
            winreg.KEY_SET_VALUE,
        )
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
        winreg.CloseKey(key)
        # Broadcast WM_SETTINGCHANGE so the taskbar / new processes pick it up
        import ctypes
        ctypes.windll.user32.SendMessageTimeoutW(
            0xFFFF,   # HWND_BROADCAST
            0x001A,   # WM_SETTINGCHANGE
            0,
            "Environment",
            0x0002,   # SMTO_ABORTIFHUNG
            1000,
            None,
        )
    except Exception:
        pass  # Best-effort; key is already set in os.environ for this session


def _ensure_api_key() -> bool:
    """
    If OPENAI_API_KEY is not set, show a dialog letting the user enter it.
    Sets the key in os.environ for the current process and persists it as a
    user-level Windows environment variable via the registry.
    Returns True to continue, False to quit.
    """
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        # Also check the Windows registry: if OPENAI_API_KEY was set as a user
        # env var after this process launched, os.environ won't have it yet.
        try:
            import winreg
            _reg = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Environment",
                0,
                winreg.KEY_READ,
            )
            _reg_val, _ = winreg.QueryValueEx(_reg, "OPENAI_API_KEY")
            winreg.CloseKey(_reg)
            key = (_reg_val or "").strip()
            if key:
                os.environ["OPENAI_API_KEY"] = key
        except (FileNotFoundError, OSError):
            pass
    if key:
        return True

    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import (
        QCheckBox,
        QDialog,
        QDialogButtonBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QVBoxLayout,
    )

    class _APIKeyDialog(QDialog):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("OpenAI API Key Required")
            self.setMinimumWidth(460)
            self.setWindowFlags(
                self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint
            )

            layout = QVBoxLayout(self)
            layout.setSpacing(12)
            layout.setContentsMargins(20, 20, 20, 20)

            icon_row = QHBoxLayout()
            icon_lbl = QLabel("[key]")
            icon_lbl.setStyleSheet("font-size: 28px;")
            title_lbl = QLabel("<b>No API key found</b>")
            title_lbl.setStyleSheet("font-size: 15px;")
            icon_row.addWidget(icon_lbl)
            icon_row.addWidget(title_lbl)
            icon_row.addStretch()
            layout.addLayout(icon_row)

            desc = QLabel(
                "PyVBAai requires an OpenAI API key to function.\n"
                "Enter your key below to save it as a permanent user\n"
                "environment variable (OPENAI_API_KEY) and continue."
            )
            desc.setWordWrap(True)
            layout.addWidget(desc)

            self._input = QLineEdit()
            self._input.setPlaceholderText("sk-...")
            self._input.setEchoMode(QLineEdit.EchoMode.Password)
            layout.addWidget(self._input)

            show_cb_row = QHBoxLayout()
            show_cb = QCheckBox("Show key")
            show_cb.toggled.connect(
                lambda checked: self._input.setEchoMode(
                    QLineEdit.EchoMode.Normal if checked
                    else QLineEdit.EchoMode.Password
                )
            )
            show_cb_row.addWidget(show_cb)
            show_cb_row.addStretch()
            layout.addLayout(show_cb_row)

            hint = QLabel(
                '<a href="https://platform.openai.com/api-keys">'
                "Get a key at platform.openai.com/api-keys</a>"
            )
            hint.setOpenExternalLinks(True)
            hint.setStyleSheet("font-size: 11px;")
            layout.addWidget(hint)

            buttons = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok
                | QDialogButtonBox.StandardButton.Cancel
            )
            buttons.accepted.connect(self._on_ok)
            buttons.rejected.connect(self.reject)
            layout.addWidget(buttons)

        def _on_ok(self) -> None:
            key = self._input.text().strip()
            if not key:
                self._input.setStyleSheet("border: 1px solid #f38ba8;")
                self._input.setPlaceholderText("Key cannot be empty")
                return
            if not key.startswith("sk-"):
                self._input.setStyleSheet("border: 1px solid #f38ba8;")
                self._input.setPlaceholderText('Key should start with "sk-"')
                self._input.clear()
                return
            os.environ["OPENAI_API_KEY"] = key
            _set_user_env_var("OPENAI_API_KEY", key)
            self.accept()

    dlg = _APIKeyDialog()
    return dlg.exec() == QDialog.DialogCode.Accepted


def _ensure_model() -> bool:
    """
    If no AI model is configured, show a dialog letting the user pick one.
    Fetches the live model list from the OpenAI API (key must already be set).
    Returns True to continue, False to quit.
    """
    from app.config import get_settings
    s = get_settings()
    if (s.value("ai/model") or "").strip():
        return True

    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import (
        QComboBox,
        QDialog,
        QDialogButtonBox,
        QLabel,
        QVBoxLayout,
    )

    from core.ai_client import AIClient

    class _ModelDialog(QDialog):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Select AI Model")
            self.setMinimumWidth(420)
            self.setWindowFlags(
                self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint
            )

            layout = QVBoxLayout(self)
            layout.setSpacing(12)
            layout.setContentsMargins(20, 20, 20, 20)

            title = QLabel("<b>Choose an AI model</b>")
            title.setStyleSheet("font-size: 14px;")
            layout.addWidget(title)

            self._status = QLabel("Fetching available models from OpenAI\u2026")
            layout.addWidget(self._status)

            self._combo = QComboBox()
            self._combo.setPlaceholderText("Select a model\u2026")
            layout.addWidget(self._combo)

            hint = QLabel(
                "The selected model will be saved and used for all AI requests.\n"
                "You can change it later in Settings."
            )
            hint.setWordWrap(True)
            hint.setStyleSheet("color: #6c7086; font-size: 11px;")
            layout.addWidget(hint)

            self._buttons = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok
                | QDialogButtonBox.StandardButton.Cancel
            )
            ok_btn = self._buttons.button(QDialogButtonBox.StandardButton.Ok)
            ok_btn.setEnabled(False)
            self._buttons.accepted.connect(self._on_ok)
            self._buttons.rejected.connect(self.reject)
            layout.addWidget(self._buttons)

            # Fetch models synchronously (API key is already validated)
            models = AIClient.fetch_models_from_api()
            if models:
                self._combo.addItems(models)
                self._combo.setCurrentIndex(-1)
                self._status.setText("Select a model to use:")
                self._combo.currentIndexChanged.connect(self._on_selection_changed)
            else:
                self._status.setText(
                    "Could not fetch model list. Check your API key and network.\n"
                    "You can type a model ID directly (e.g. gpt-4o)."
                )
                self._combo.setEditable(True)
                self._combo.currentTextChanged.connect(self._on_text_changed)

        def _on_selection_changed(self, idx: int) -> None:
            ok_btn = self._buttons.button(QDialogButtonBox.StandardButton.Ok)
            ok_btn.setEnabled(idx >= 0)

        def _on_text_changed(self, text: str) -> None:
            ok_btn = self._buttons.button(QDialogButtonBox.StandardButton.Ok)
            ok_btn.setEnabled(bool(text.strip()))

        def _on_ok(self) -> None:
            model = self._combo.currentText().strip()
            if not model:
                return
            get_settings().setValue("ai/model", model)
            self.accept()

    dlg = _ModelDialog()
    return dlg.exec() == QDialog.DialogCode.Accepted


def _make_app_icon():
    """Draw the PyVBAai app icon using Qt primitives. No image file required."""
    from PySide6.QtCore import QRectF
    from PySide6.QtGui import QBrush, QColor, QIcon, QPainter, QPainterPath, QPixmap

    size = 256
    px = QPixmap(size, size)
    px.fill(QColor(0, 0, 0, 0))

    p = QPainter(px)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)

    # Outer rounded square background
    bg_path = QPainterPath()
    bg_path.addRoundedRect(QRectF(0, 0, size, size), 52, 52)
    p.fillPath(bg_path, QBrush(QColor("#1e1e2e")))

    # Page body with clearly rounded corners (no fold - too small to read at icon sizes)
    margin = 50
    page_w = size - margin * 2
    page_h = int(page_w * 1.28)
    page_x = margin
    page_y = (size - page_h) // 2 + 2
    page_r = 26
    page_path = QPainterPath()
    page_path.addRoundedRect(QRectF(page_x, page_y, page_w, page_h), page_r, page_r)
    p.fillPath(page_path, QBrush(QColor("#e8e8f4")))

    # Three text lines
    dot_color = QBrush(QColor("#89b4fa"))
    h_pad = 42
    line_start_x = page_x + h_pad
    line_y_start = page_y + 28
    line_gap = 22
    line_w_long = page_w - h_pad * 2
    line_w_short = int(line_w_long * 0.55)
    dot_h = 10
    for i, lw in enumerate([line_w_long, line_w_long, line_w_short]):
        ly = line_y_start + i * line_gap
        line_path = QPainterPath()
        line_path.addRoundedRect(QRectF(line_start_x, ly, lw, dot_h), 5, 5)
        p.fillPath(line_path, dot_color)

    p.end()
    return QIcon(px)


def main() -> None:
    from PySide6.QtWidgets import QApplication

    # Must be created before any other Qt objects
    app = QApplication(sys.argv)
    app.setApplicationName("PyVBAai")
    app.setApplicationDisplayName("PyVBAai")
    app.setOrganizationName("PyVBAai")
    app.setOrganizationDomain("pyvbaai.local")
    app.setWindowIcon(_make_app_icon())

    # Initialise debug logging early so all subsequent code can log
    from app.logger import init_logging
    init_logging()

    # Apply theme before any widgets are constructed
    from app.settings_dialog import SettingsDialog
    from app.theme import apply_theme

    dark = SettingsDialog.load_dark_mode()
    apply_theme(app, dark=dark)

    # ── API key check ─────────────────────────────────────────────────────────
    if not _ensure_api_key():
        sys.exit(0)

    # ── Model selection (required before opening main window) ────────────────
    if not _ensure_model():
        sys.exit(0)

    from app.main_window import MainWindow

    window = MainWindow()
    window.show()

    # If an Excel file was passed on the command line, open it immediately
    if len(sys.argv) > 1:
        candidate = sys.argv[1]
        if os.path.isfile(candidate) and candidate.lower().endswith(
            (".xlsx", ".xlsm", ".xls")
        ):
            window.load_file(candidate)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

