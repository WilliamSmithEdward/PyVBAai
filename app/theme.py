# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Dark / light QSS themes for PyVBAai."""
from __future__ import annotations

import pathlib
import tempfile

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QPainter, QPalette, QPen
from PySide6.QtWidgets import (
    QApplication,
    QProxyStyle,
    QStyle,
)

_style_ref: _AppStyle | None = None


class _AppStyle(QProxyStyle):
    """Fusion-based proxy that draws visible checkmarks on all checkboxes."""

    _DARK = {
        "accent":     "#89b4fa",
        "check_fg":   "#1e1e2e",
        "bg_off":     "#313244",
        "border_off": "#585b70",
    }
    _LIGHT = {
        "accent":     "#1e66f5",
        "check_fg":   "#ffffff",
        "bg_off":     "#dce0e8",
        "border_off": "#bcc0cc",
    }

    def __init__(self, dark: bool = True) -> None:
        super().__init__("fusion")
        self._palette = self._DARK if dark else self._LIGHT

    def drawPrimitive(self, element, option, painter, widget=None):  # type: ignore[override]
        _checkbox_elements = (
            QStyle.PrimitiveElement.PE_IndicatorCheckBox,
            QStyle.PrimitiveElement.PE_IndicatorItemViewItemCheck,
        )
        if element not in _checkbox_elements:
            super().drawPrimitive(element, option, painter, widget)
            return

        checked  = bool(option.state & QStyle.StateFlag.State_On)
        partial  = bool(option.state & QStyle.StateFlag.State_NoChange)
        r = option.rect.adjusted(2, 2, -2, -2)
        p = self._palette
        accent = QColor(p["accent"])
        check_fg = QColor(p["check_fg"])
        bg_off = QColor(p["bg_off"])
        border_off = QColor(p["border_off"])

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        if checked:
            # Fully checked: solid accent fill + white checkmark
            painter.setPen(QPen(accent, 1.0))
            painter.setBrush(QBrush(accent))
            painter.drawRoundedRect(r, 3, 3)
            pen = QPen(check_fg, 1.8, Qt.PenStyle.SolidLine,
                       Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            x, y, w, h = r.x(), r.y(), r.width(), r.height()
            painter.drawLine(int(x + w * 0.18), int(y + h * 0.50),
                             int(x + w * 0.42), int(y + h * 0.76))
            painter.drawLine(int(x + w * 0.42), int(y + h * 0.76),
                             int(x + w * 0.82), int(y + h * 0.24))
        elif partial:
            # Partially checked: accent border + accent dash bar
            painter.setPen(QPen(accent, 1.5))
            painter.setBrush(QBrush(bg_off))
            painter.drawRoundedRect(r, 3, 3)
            x, y, w, h = r.x(), r.y(), r.width(), r.height()
            bar_y = y + h // 2
            pen = QPen(accent, 2.5, Qt.PenStyle.SolidLine,
                       Qt.PenCapStyle.RoundCap)
            painter.setPen(pen)
            painter.drawLine(int(x + w * 0.22), bar_y,
                             int(x + w * 0.78), bar_y)
        else:
            # Unchecked: muted fill + border
            painter.setPen(QPen(border_off, 1.5))
            painter.setBrush(QBrush(bg_off))
            painter.drawRoundedRect(r, 3, 3)

        painter.restore()

# ── Catppuccin Mocha (dark) ───────────────────────────────────────────────
DARK_QSS = """
/* ── Base ──────────────────────────────────────────────────────────── */
QMainWindow, QDialog, QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 13px;
}

QSplitter::handle {
    background-color: #313244;
}
QSplitter::handle:horizontal { width: 2px; }
QSplitter::handle:vertical   { height: 2px; }

/* ── Toolbar ────────────────────────────────────────────────────────── */
QToolBar {
    background-color: #181825;
    border-bottom: 1px solid #313244;
    spacing: 4px;
    padding: 4px 8px;
}
QToolBar QLabel {
    color: #cdd6f4;
    font-weight: 600;
    font-size: 15px;
}
QToolButton {
    background: transparent;
    border: none;
    border-radius: 6px;
    padding: 5px 10px;
    color: #cdd6f4;
}
QToolButton:hover  { background-color: #313244; }
QToolButton:pressed { background-color: #45475a; }

/* ── Sidebar / Workbook Panel ───────────────────────────────────────── */
#workbookPanel {
    background-color: #181825;
    border-right: 1px solid #313244;
}
#sidebarHeader {
    background-color: #181825;
    border-bottom: 1px solid #313244;
    padding: 8px 12px;
    font-weight: 600;
    color: #cba6f7;
}
#filterBar {
    background-color: #181825;
    border-bottom: 1px solid #313244;
}
QToolButton#filterBtn {
    background-color: #313244;
    color: #cdd6f4;
    border: none;
    border-radius: 5px;
    padding: 3px 10px;
    font-size: 12px;
}
QToolButton#filterBtn:hover    { background-color: #45475a; }
QToolButton#filterBtn::menu-indicator { width: 0; image: none; }
QTreeWidget {
    background-color: #181825;
    border: none;
    color: #cdd6f4;
    outline: none;
    padding: 4px;
}
QTreeWidget::item { padding: 3px 2px; border-radius: 4px; }
QTreeWidget::item:hover    { background-color: #313244; }
QTreeWidget::item:selected { background-color: #45475a; color: #cdd6f4; }
QTreeWidget::branch { background: transparent; }

/* ── Chat Area ──────────────────────────────────────────────────────── */
#chatScrollArea        { background-color: #1e1e2e; border: none; }
#chatScrollArea > QWidget > QWidget { background-color: #1e1e2e; }
#chatContainer         { background-color: #1e1e2e; }

/* User bubble */
#userBubble {
    background-color: #1e2e4a;
    border-radius: 14px;
}
#userBubble QTextBrowser {
    background-color: transparent;
    color: #c8d8f0;
    border: none;
    font-size: 13px;
    font-weight: 500;
}

/* AI bubble */
#aiBubble {
    background-color: #313244;
    border-radius: 14px;
}
#aiBubble QTextBrowser {
    background-color: transparent;
    color: #cdd6f4;
    border: none;
    font-size: 13px;
}

/* System / info bubble */
#sysBubble {
    background-color: #1e1e2e;
}
#sysBubble QLabel {
    color: #6c7086;
    font-size: 11px;
}

/* Typing indicator */
#typingBubble {
    background-color: #313244;
    border-radius: 14px;
}
#typingLabel { color: #6c7086; font-size: 20px; letter-spacing: 2px; }

/* ── Input Area ─────────────────────────────────────────────────────── */
#inputArea {
    background-color: #181825;
    border-top: 1px solid #313244;
}
#messageInput {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 10px;
    color: #cdd6f4;
    padding: 8px 12px;
    font-size: 13px;
    selection-background-color: #89b4fa;
}
#messageInput:focus { border-color: #89b4fa; }
#messageInput:disabled { background-color: #1e1e2e; border-color: #313244; color: #45475a; }

/* ── Buttons ────────────────────────────────────────────────────────── */
QPushButton {
    background-color: #2a4a8a;
    color: #c8d8f0;
    border: none;
    border-radius: 8px;
    padding: 8px 18px;
    font-weight: 600;
    font-size: 13px;
}
QPushButton:hover   { background-color: #345aa0; }
QPushButton:pressed { background-color: #1e3870; }
QPushButton:disabled { background-color: #313244; color: #6c7086; }

QPushButton#dangerBtn {
    background-color: #7a2535;
    color: #f2b8c0;
}
QPushButton#dangerBtn:hover { background-color: #922d40; }

QPushButton#secondaryBtn {
    background-color: #45475a;
    color: #cdd6f4;
}
QPushButton#secondaryBtn:hover { background-color: #585b70; }

QPushButton#ghostBtn {
    background-color: transparent;
    color: #8aaad8;
    border: 1px solid #3a5a9a;
}
QPushButton#ghostBtn:hover { background-color: #1e2e4a; }

/* ── ComboBox ───────────────────────────────────────────────────────── */
QComboBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 6px;
    color: #cdd6f4;
    padding: 4px 8px;
    min-width: 120px;
}
QComboBox::drop-down { border: none; width: 28px; }
QComboBox::down-arrow { image: url(__ARROW_PATH__); width: 12px; height: 8px; }
QComboBox QAbstractItemView {
    border: 1px solid #45475a;
    color: #cdd6f4;
    selection-background-color: #45475a;
    outline: none;
}

/* ── ScrollBars ─────────────────────────────────────────────────────── */
QScrollBar:vertical {
    background-color: transparent;
    width: 8px;
    margin: 0;
    border: none;
}
QScrollBar::handle:vertical {
    background-color: #45475a;
    border-radius: 4px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover  { background-color: #585b70; }
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical { height: 0; background: none; border: none; }
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical { background: none; }

QScrollBar:horizontal {
    background-color: transparent;
    height: 8px;
    border: none;
}
QScrollBar::handle:horizontal {
    background-color: #45475a;
    border-radius: 4px;
    min-width: 24px;
}
QScrollBar::handle:horizontal:hover { background-color: #585b70; }
QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal { width: 0; background: none; border: none; }

/* ── StatusBar ──────────────────────────────────────────────────────── */
QStatusBar {
    background-color: #181825;
    color: #6c7086;
    border-top: 1px solid #313244;
    font-size: 11px;
}

/* ── TabWidget ──────────────────────────────────────────────────────── */
QTabWidget::pane  { border: 1px solid #313244; border-radius: 4px; background: #1e1e2e; }
QTabBar::tab {
    background: #313244;
    color: #6c7086;
    padding: 6px 16px;
    border-radius: 0;
}
QTabBar::tab:selected { background: #45475a; color: #cdd6f4; }
QTabBar::tab:hover    { background: #45475a; }

/* ── LineEdit ───────────────────────────────────────────────────────── */
QLineEdit {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 6px;
    color: #cdd6f4;
    padding: 5px 8px;
}
QLineEdit:focus  { border-color: #89b4fa; }

/* ── CheckBox / SpinBox ─────────────────────────────────────────────── */
QCheckBox { color: #cdd6f4; spacing: 6px; }

QSpinBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 6px;
    color: #cdd6f4;
    padding: 4px 8px;
}
QSpinBox:focus { border-color: #89b4fa; }

/* ── GroupBox ───────────────────────────────────────────────────────── */
QGroupBox {
    border: 1px solid #313244;
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 8px;
    color: #cdd6f4;
    font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
    color: #89b4fa;
}

/* ── ToolTip ────────────────────────────────────────────────────────── */
QToolTip {
    background-color: #313244;
    color: #cdd6f4;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 12px;
}

/* ── TextBrowser (code viewer) ──────────────────────────────────────── */
QTextBrowser#codeViewer {
    background-color: #181825;
    color: #cdd6f4;
    border: 1px solid #313244;
    border-radius: 6px;
    font-family: "Cascadia Mono", "Consolas", monospace;
    font-size: 12px;
    padding: 8px;
}
"""

# ── Catppuccin Latte (light) ───────────────────────────────────────────────
LIGHT_QSS = """
QMainWindow, QDialog, QWidget {
    background-color: #eff1f5;
    color: #4c4f69;
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 13px;
}
QSplitter::handle { background-color: #ccd0da; }
QSplitter::handle:horizontal { width: 2px; }
QSplitter::handle:vertical   { height: 2px; }

QToolBar {
    background-color: #e6e9ef;
    border-bottom: 1px solid #ccd0da;
    spacing: 4px;
    padding: 4px 8px;
}
QToolBar QLabel { color: #4c4f69; font-weight: 600; font-size: 15px; }
QToolButton { background: transparent; border: none; border-radius: 6px; padding: 5px 10px; color: #4c4f69; }
QToolButton:hover  { background-color: #ccd0da; }
QToolButton:pressed { background-color: #bcc0cc; }

#workbookPanel { background-color: #e6e9ef; border-right: 1px solid #ccd0da; }
#sidebarHeader { background-color: #e6e9ef; border-bottom: 1px solid #ccd0da; padding: 8px 12px; font-weight: 600; color: #8839ef; }
#filterBar { background-color: #e6e9ef; border-bottom: 1px solid #ccd0da; }
QToolButton#filterBtn { background-color: #ccd0da; color: #4c4f69; border: none; border-radius: 5px; padding: 3px 10px; font-size: 12px; }
QToolButton#filterBtn:hover { background-color: #bcc0cc; }
QToolButton#filterBtn::menu-indicator { width: 0; image: none; }
QTreeWidget { background-color: #e6e9ef; border: none; color: #4c4f69; outline: none; padding: 4px; }
QTreeWidget::item { padding: 3px 2px; border-radius: 4px; }
QTreeWidget::item:hover    { background-color: #ccd0da; }
QTreeWidget::item:selected { background-color: #bcc0cc; color: #4c4f69; }
QTreeWidget::branch { background: transparent; }

#chatScrollArea, #chatScrollArea > QWidget > QWidget, #chatContainer { background-color: #eff1f5; border: none; }

#userBubble { background-color: #1e66f5; border-radius: 14px; }
#userBubble QTextBrowser { background-color: transparent; color: #eff1f5; border: none; font-size: 13px; }

#aiBubble { background-color: #dce0e8; border-radius: 14px; }
#aiBubble QTextBrowser { background-color: transparent; color: #4c4f69; border: none; font-size: 13px; }

#sysBubble { background-color: #eff1f5; }
#sysBubble QLabel { color: #9ca0b0; font-size: 11px; }

#typingBubble { background-color: #dce0e8; border-radius: 14px; }
#typingLabel  { color: #9ca0b0; font-size: 20px; letter-spacing: 2px; }

#inputArea { background-color: #e6e9ef; border-top: 1px solid #ccd0da; }
#messageInput {
    background-color: #dce0e8;
    border: 1px solid #bcc0cc;
    border-radius: 10px;
    color: #4c4f69;
    padding: 8px 12px;
    font-size: 13px;
}
#messageInput:focus { border-color: #1e66f5; }
#messageInput:disabled { background-color: #eff1f5; border-color: #ccd0da; color: #9ca0b0; }

QPushButton { background-color: #1e66f5; color: #eff1f5; border: none; border-radius: 8px; padding: 8px 18px; font-weight: 600; font-size: 13px; }
QPushButton:hover   { background-color: #04a5e5; }
QPushButton:pressed { background-color: #7287fd; }
QPushButton:disabled { background-color: #dce0e8; color: #9ca0b0; }

QPushButton#dangerBtn  { background-color: #d20f39; color: #eff1f5; }
QPushButton#dangerBtn:hover { background-color: #e64553; }
QPushButton#secondaryBtn { background-color: #ccd0da; color: #4c4f69; }
QPushButton#secondaryBtn:hover { background-color: #bcc0cc; }
QPushButton#ghostBtn { background-color: transparent; color: #1e66f5; border: 1px solid #1e66f5; }
QPushButton#ghostBtn:hover { background-color: #dce0e8; }

QComboBox { background-color: #dce0e8; border: 1px solid #bcc0cc; border-radius: 6px; color: #4c4f69; padding: 4px 8px; min-width: 120px; }
QComboBox::drop-down { border: none; width: 28px; }
QComboBox::down-arrow { image: url(__ARROW_PATH__); width: 12px; height: 8px; }
QComboBox QAbstractItemView { background-color: #dce0e8; border: 1px solid #bcc0cc; color: #4c4f69; selection-background-color: #bcc0cc; outline: none; }

QScrollBar:vertical { background: transparent; width: 8px; border: none; }
QScrollBar::handle:vertical { background-color: #bcc0cc; border-radius: 4px; min-height: 24px; }
QScrollBar::handle:vertical:hover { background-color: #9ca0b0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; background: none; border: none; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }
QScrollBar:horizontal { background: transparent; height: 8px; border: none; }
QScrollBar::handle:horizontal { background-color: #bcc0cc; border-radius: 4px; min-width: 24px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; background: none; border: none; }

QStatusBar { background-color: #e6e9ef; color: #9ca0b0; border-top: 1px solid #ccd0da; font-size: 11px; }

QTabWidget::pane { border: 1px solid #ccd0da; border-radius: 4px; background: #eff1f5; }
QTabBar::tab { background: #dce0e8; color: #9ca0b0; padding: 6px 16px; }
QTabBar::tab:selected { background: #bcc0cc; color: #4c4f69; }
QTabBar::tab:hover    { background: #bcc0cc; }

QLineEdit { background-color: #dce0e8; border: 1px solid #bcc0cc; border-radius: 6px; color: #4c4f69; padding: 5px 8px; }
QLineEdit:focus { border-color: #1e66f5; }

QCheckBox { color: #4c4f69; spacing: 6px; }

QSpinBox { background-color: #dce0e8; border: 1px solid #bcc0cc; border-radius: 6px; color: #4c4f69; padding: 4px 8px; }
QSpinBox:focus { border-color: #1e66f5; }

QGroupBox { border: 1px solid #ccd0da; border-radius: 6px; margin-top: 12px; padding-top: 8px; color: #4c4f69; font-weight: 600; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 6px; color: #1e66f5; }

QToolTip { background-color: #dce0e8; color: #4c4f69; border: 1px solid #bcc0cc; border-radius: 4px; padding: 4px 8px; font-size: 12px; }

QTextBrowser#codeViewer { background-color: #e6e9ef; color: #4c4f69; border: 1px solid #ccd0da; border-radius: 6px; font-family: "Cascadia Mono","Consolas",monospace; font-size: 12px; padding: 8px; }
"""


def _chevron_svg_path(stroke: str) -> str:
    """Write a chevron SVG to a temp file and return a forward-slash path for QSS url()."""
    safe = stroke.lstrip("#")
    p = pathlib.Path(tempfile.gettempdir()) / f"pyvbaai_chevron_{safe}.svg"
    if not p.exists():
        p.write_text(
            f'<svg xmlns="http://www.w3.org/2000/svg" width="12" height="8" viewBox="0 0 12 8">'
            f'<polyline points="1,1 6,7 11,1" fill="none" stroke="#{safe}"'
            f' stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
            encoding="utf-8",
        )
    return str(p).replace("\\", "/")


def apply_theme(app: QApplication, dark: bool = True) -> None:
    global _style_ref
    _style_ref = _AppStyle(dark)
    app.setStyle(_style_ref)
    arrow = _chevron_svg_path("#89b4fa" if dark else "#1e66f5")
    qss = (DARK_QSS if dark else LIGHT_QSS).replace("__ARROW_PATH__", arrow)
    app.setStyleSheet(qss)

    # Set QPalette so QTextEdit / QAbstractScrollArea viewports use the right
    # Base (background) and Text colors regardless of OS dark-mode overrides.
    palette = app.palette()
    if dark:
        palette.setColor(QPalette.ColorRole.Base,   QColor("#313244"))
        palette.setColor(QPalette.ColorRole.Text,   QColor("#cdd6f4"))
        palette.setColor(QPalette.ColorRole.Window, QColor("#1e1e2e"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor("#cdd6f4"))
        palette.setColor(QPalette.ColorRole.PlaceholderText, QColor("#6c7086"))
    else:
        palette.setColor(QPalette.ColorRole.Base,   QColor("#dce0e8"))
        palette.setColor(QPalette.ColorRole.Text,   QColor("#4c4f69"))
        palette.setColor(QPalette.ColorRole.Window, QColor("#eff1f5"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor("#4c4f69"))
        palette.setColor(QPalette.ColorRole.PlaceholderText, QColor("#9ca0b0"))
    app.setPalette(palette)
