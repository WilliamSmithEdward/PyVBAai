# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Chat widget: message bubbles, typing indicator, and input bar."""
from __future__ import annotations

import markdown as md
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

_MD_EXTENSIONS = ["fenced_code", "tables", "nl2br"]


def _md_to_html(text: str) -> str:
    """Convert Markdown text to an HTML fragment."""
    return md.markdown(text, extensions=_MD_EXTENSIONS)


# ── Auto-sizing text browser used inside bubbles ──────────────────────────────

class _BubbleText(QTextBrowser):
    """QTextBrowser that auto-sizes its height to fit content."""

    def __init__(self, html: str, is_user: bool, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setReadOnly(True)
        self.setOpenExternalLinks(True)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setFrameStyle(QFrame.Shape.NoFrame)
        # Transparent so parent bubble frame shows through
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # Inline style for code blocks inside the HTML
        if is_user:
            code_bg, code_fg = "#1e1e2e", "#89b4fa"
        else:
            code_bg, code_fg = "#1e1e2e", "#a6e3a1"
        style_prefix = (
            f"<style>code, pre {{ background:{code_bg}; color:{code_fg}; "
            f"border-radius:4px; padding:2px 4px; "
            f"font-family:'Cascadia Mono','Consolas',monospace; font-size:12px; }}"
            f"pre {{ padding:8px; display:block; }}</style>"
        )
        self.setHtml(style_prefix + html)

    def _recalc_height(self) -> None:
        doc = self.document()
        doc.setTextWidth(self.viewport().width())
        margin = int(doc.documentMargin())
        h = int(doc.size().height()) + margin * 2 + 4
        self.setMinimumHeight(h)
        self.setMaximumHeight(h)

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._recalc_height()

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        self._recalc_height()


# ── Individual message bubble ─────────────────────────────────────────────────

class MessageBubble(QWidget):
    """A single chat message bubble (user right-aligned, AI left-aligned)."""

    MAX_BUBBLE_WIDTH = 680

    def __init__(self, content: str, role: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.role = role
        is_user = (role == "user")

        html = _md_to_html(content)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(12, 4, 12, 4)
        outer.setSpacing(0)

        if role == "system":
            # Centered system note rendered as HTML
            lbl = QLabel()
            lbl.setTextFormat(Qt.TextFormat.RichText)
            lbl.setText(_md_to_html(content))
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setObjectName("sysBubble")
            container = QFrame()
            container.setObjectName("sysBubble")
            cl = QHBoxLayout(container)
            cl.setContentsMargins(8, 4, 8, 4)
            cl.addWidget(lbl)
            outer.addWidget(container, alignment=Qt.AlignmentFlag.AlignCenter)
            return

        bubble = QFrame()
        bubble.setObjectName("userBubble" if is_user else "aiBubble")
        bubble.setMaximumWidth(self.MAX_BUBBLE_WIDTH)
        bubble.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)

        bl = QVBoxLayout(bubble)
        bl.setContentsMargins(12, 8, 12, 8)
        bl.setSpacing(0)

        text = _BubbleText(html, is_user, bubble)
        bl.addWidget(text)

        if is_user:
            outer.addStretch()
            outer.addWidget(bubble)
        else:
            outer.addWidget(bubble)
            outer.addStretch()


# ── Typing indicator ──────────────────────────────────────────────────────────

class TypingIndicator(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        outer = QHBoxLayout(self)
        outer.setContentsMargins(12, 4, 12, 4)

        bubble = QFrame()
        bubble.setObjectName("typingBubble")
        bl = QVBoxLayout(bubble)
        bl.setContentsMargins(16, 10, 16, 10)

        self._label = QLabel("...")
        self._label.setObjectName("typingLabel")
        bl.addWidget(self._label)

        outer.addWidget(bubble)
        outer.addStretch()

        self._dots = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(400)

    def _tick(self) -> None:
        self._dots = (self._dots + 1) % 4
        filled = "." * self._dots + " " * (3 - self._dots)
        self._label.setText(filled.strip() or "...")

    def stop(self) -> None:
        self._timer.stop()


# ── Multi-line input box that sends on Enter (Shift+Enter = newline) ──────────

class ChatInput(QTextEdit):
    send_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("messageInput")
        self.setPlaceholderText("Message PyVBAai...  (Enter to send, Shift+Enter for newline)")
        self.setMaximumHeight(120)
        self.setMinimumHeight(42)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.document().contentsChanged.connect(self._auto_resize)
        self.document().documentLayout().documentSizeChanged.connect(self._auto_resize)

    def _auto_resize(self, *_args) -> None:
        doc = self.document()
        doc.setTextWidth(self.viewport().width())
        doc_h = int(doc.size().height()) + 16
        clamped = max(42, min(doc_h, 120))
        self.setFixedHeight(clamped)

    def keyPressEvent(self, event: QKeyEvent) -> None:  # type: ignore[override]
        if (
            event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
            and not (event.modifiers() & Qt.KeyboardModifier.ShiftModifier)
        ):
            self.send_requested.emit()
        else:
            super().keyPressEvent(event)


# ── Main chat widget ──────────────────────────────────────────────────────────

class ChatWidget(QWidget):
    """Full chat UI: scrollable message list + input bar."""

    message_sent = Signal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._typing: TypingIndicator | None = None
        self._setup_ui()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Scroll area ──────────────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setObjectName("chatScrollArea")
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self._container = QWidget()
        self._container.setObjectName("chatContainer")
        self._msg_layout = QVBoxLayout(self._container)
        self._msg_layout.setContentsMargins(0, 8, 0, 8)
        self._msg_layout.setSpacing(2)
        self._msg_layout.addStretch()   # pushes messages downward initially

        self._scroll.setWidget(self._container)
        root.addWidget(self._scroll)

        # ── Input bar ────────────────────────────────────────────────────────
        input_area = QWidget()
        input_area.setObjectName("inputArea")
        input_layout = QHBoxLayout(input_area)
        input_layout.setContentsMargins(12, 8, 12, 8)
        input_layout.setSpacing(8)

        self._input = ChatInput()
        self._input.send_requested.connect(self._on_send)

        self._send_btn = QPushButton("Send")
        self._send_btn.setFixedWidth(72)
        self._send_btn.clicked.connect(self._on_send)

        input_layout.addWidget(self._input)
        input_layout.addWidget(self._send_btn, alignment=Qt.AlignmentFlag.AlignBottom)

        root.addWidget(input_area)

    # ── Public API ────────────────────────────────────────────────────────────

    def add_message(self, content: str, role: str = "assistant") -> None:
        """Add a message bubble (role: 'user', 'assistant', or 'system')."""
        bubble = MessageBubble(content, role, self._container)
        # Insert before the final stretch item
        count = self._msg_layout.count()
        self._msg_layout.insertWidget(count - 1, bubble)
        QTimer.singleShot(50, self._scroll_to_bottom)

    def show_typing(self) -> None:
        """Show the animated typing indicator."""
        self.hide_typing()
        self._typing = TypingIndicator(self._container)
        count = self._msg_layout.count()
        self._msg_layout.insertWidget(count - 1, self._typing)
        QTimer.singleShot(50, self._scroll_to_bottom)

    def hide_typing(self) -> None:
        """Remove the typing indicator."""
        if self._typing is not None:
            self._typing.stop()
            self._msg_layout.removeWidget(self._typing)
            self._typing.deleteLater()
            self._typing = None

    def set_input_enabled(self, enabled: bool) -> None:
        self._input.setEnabled(enabled)
        self._send_btn.setEnabled(enabled)

    def clear_input(self) -> None:
        self._input.clear()

    def focus_input(self) -> None:
        self._input.setFocus()

    # ── Private ───────────────────────────────────────────────────────────────

    def _on_send(self) -> None:
        text = self._input.toPlainText().strip()
        if not text:
            return
        self._input.clear()
        self.message_sent.emit(text)

    def _scroll_to_bottom(self) -> None:
        sb = self._scroll.verticalScrollBar()
        sb.setValue(sb.maximum())
