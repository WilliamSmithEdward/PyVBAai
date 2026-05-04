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


def main() -> None:
    from PySide6.QtWidgets import QApplication

    # Must be created before any other Qt objects
    app = QApplication(sys.argv)
    app.setApplicationName("PyVBAai")
    app.setApplicationDisplayName("PyVBAai")
    app.setOrganizationName("PyVBAai")
    app.setOrganizationDomain("pyvbaai.local")

    # Apply theme before any widgets are constructed
    from app.settings_dialog import SettingsDialog
    from app.theme import apply_theme

    dark = SettingsDialog.load_dark_mode()
    apply_theme(app, dark=dark)

    # ── API key check ─────────────────────────────────────────────────────────
    if not _ensure_api_key():
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

