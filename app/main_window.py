# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Main application window."""
from __future__ import annotations

import os

from PySide6.QtCore import QSettings, Qt, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QLabel,
    QMainWindow,
    QPushButton,
    QSplitter,
    QStatusBar,
    QToolBar,
    QWidget,
)

from app.chat_widget import ChatWidget
from app.preview_dialog import PreviewDialog
from app.settings_dialog import SettingsDialog
from app.theme import apply_theme
from app.workbook_panel import VBAViewerDialog, WorkbookPanel
from app.workers import AIWorker, ExcelReaderWorker, ExcelWriterWorker, NewWorkbookWorker
from core.context_builder import build_context, estimate_tokens
from models.conversation import AIResponse, Conversation
from models.workbook import WorkbookData

_WELCOME = (
    "## Welcome to PyVBAai\n\n"
    "Load an Excel file to get started - drag and drop it onto this window, "
    "click **Open File** to open an existing workbook, or click **New Workbook** "
    "to create a fresh one.\n\n"
    "Once loaded, chat naturally to inspect, modify, or extend your workbook.\n\n"
    "_Powered by OpenAI GPT. Changes are always previewed before applying._"
)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PyVBAai")
        self.setMinimumSize(1000, 640)
        self.resize(1280, 780)
        self.setAcceptDrops(True)

        # State
        self._wb: WorkbookData | None = None
        self._conversation = Conversation()
        self._ai_worker: AIWorker | None = None
        self._reader_worker: ExcelReaderWorker | None = None
        self._writer_worker: ExcelWriterWorker | None = None
        self._new_wb_worker: NewWorkbookWorker | None = None
        self._pending_response: AIResponse | None = None
        self._dark_mode: bool = SettingsDialog.load_dark_mode()

        self._build_ui()
        self._apply_current_theme()
        self._update_status()

        # Show welcome message
        self._chat.add_message(_WELCOME, "assistant")

    # ── UI Construction ────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── Toolbar ───────────────────────────────────────────────────────────
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        self.addToolBar(toolbar)

        title_lbl = QLabel("  PyVBAai  ")
        title_lbl.setStyleSheet("font-size: 16px; font-weight: 700; letter-spacing: 1px;")
        toolbar.addWidget(title_lbl)

        toolbar.addSeparator()

        open_btn = QPushButton("Open File")
        open_btn.clicked.connect(self._open_file_dialog)
        toolbar.addWidget(open_btn)

        new_wb_btn = QPushButton("New Workbook")
        new_wb_btn.clicked.connect(self._new_workbook_dialog)
        toolbar.addWidget(new_wb_btn)

        toolbar.addSeparator()

        self._model_combo = QComboBox()
        from core.ai_client import AIClient
        for m in AIClient.available_models():
            self._model_combo.addItem(m)
        saved_model = SettingsDialog.load_model()
        idx = self._model_combo.findText(saved_model)
        if idx >= 0:
            self._model_combo.setCurrentIndex(idx)
        self._model_combo.currentTextChanged.connect(self._on_model_changed)
        toolbar.addWidget(QLabel("  Model: "))
        toolbar.addWidget(self._model_combo)

        toolbar.addSeparator()

        settings_btn = QPushButton("Settings")
        settings_btn.setObjectName("secondaryBtn")
        settings_btn.clicked.connect(self._open_settings)
        toolbar.addWidget(settings_btn)

        # Spacer
        spacer = QWidget()
        spacer.setSizePolicy(
            spacer.sizePolicy().horizontalPolicy(),
            spacer.sizePolicy().verticalPolicy(),
        )
        from PySide6.QtWidgets import QSizePolicy
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        theme_btn = QPushButton("Dark")
        theme_btn.setObjectName("ghostBtn")
        theme_btn.setFixedWidth(90)
        theme_btn.clicked.connect(self._toggle_theme)
        self._theme_btn = theme_btn
        toolbar.addWidget(theme_btn)

        new_chat_btn = QPushButton("New Chat")
        new_chat_btn.setObjectName("secondaryBtn")
        new_chat_btn.clicked.connect(self._new_chat)
        toolbar.addWidget(new_chat_btn)

        toolbar.addWidget(QLabel("  "))

        # ── Central splitter ──────────────────────────────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(4)

        self._wb_panel = WorkbookPanel()
        self._wb_panel.vba_view_requested.connect(self._show_vba_viewer)
        self._wb_panel.context_changed.connect(self._update_status)
        splitter.addWidget(self._wb_panel)

        self._chat = ChatWidget()
        self._chat.message_sent.connect(self._on_user_message)
        splitter.addWidget(self._chat)

        splitter.setStretchFactor(0, 0)   # sidebar: fixed
        splitter.setStretchFactor(1, 1)   # chat: stretches
        splitter.setSizes([240, 900])

        self.setCentralWidget(splitter)

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = QStatusBar()
        self._status.setSizeGripEnabled(False)
        self.setStatusBar(self._status)
        self._status_file_lbl = QLabel("  No file loaded")
        self._status_token_lbl = QLabel("")
        self._status.addWidget(self._status_file_lbl)
        self._status.addPermanentWidget(self._status_token_lbl)

        # Initially input is disabled until a file is loaded
        self._chat.set_input_enabled(False)

    # ── File Loading ───────────────────────────────────────────────────────────

    def _open_file_dialog(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*)",
        )
        if path:
            self.load_file(path)

    def _new_workbook_dialog(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "New Excel Workbook",
            "Book1.xlsx",
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return
        path = os.path.normpath(path)
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self._chat.add_message(f"_Creating new workbook **{os.path.basename(path)}**..._", "system")
        self._chat.set_input_enabled(False)
        self._new_wb_worker = NewWorkbookWorker(path)
        self._new_wb_worker.finished.connect(self._on_new_workbook_created)
        self._new_wb_worker.error.connect(self._on_new_workbook_error)
        self._new_wb_worker.start()

    def _on_new_workbook_created(self, path: str) -> None:
        self.load_file(path)

    def _on_new_workbook_error(self, err: str) -> None:
        self._chat.set_input_enabled(True)
        self._chat.add_message(f"**Failed to create workbook:**\n\n{err}", "system")

    def load_file(self, path: str) -> None:
        self._chat.add_message(f"_Loading **{os.path.basename(path)}**..._", "system")
        self._chat.set_input_enabled(False)
        self._status_file_lbl.setText(f"  Loading {os.path.basename(path)}...")

        config = SettingsDialog.load_context_config()
        self._reader_worker = ExcelReaderWorker(path, config)
        self._reader_worker.finished.connect(self._on_workbook_loaded)
        self._reader_worker.error.connect(self._on_reader_error)
        self._reader_worker.start()

    def _on_workbook_loaded(self, wb: WorkbookData) -> None:
        self._wb = wb
        self._wb_panel.set_workbook(wb)
        self._conversation.workbook_path = wb.file_path
        self._chat.set_input_enabled(True)
        self._update_status()

        has_vba = wb.has_vba
        sheet_count = len(wb.sheets)
        vba_count = len(wb.vba_modules)
        nr_count = len(wb.named_ranges)

        msg_parts = [f"**{wb.name}** loaded successfully."]
        msg_parts.append(f"- {sheet_count} sheet(s)")
        if has_vba:
            msg_parts.append(f"- {vba_count} VBA module(s)")
        if nr_count:
            msg_parts.append(f"- {nr_count} named range(s)")
        if wb.extraction_error:
            msg_parts.append(f"\nWarning: {wb.extraction_error}")

        msg_parts.append("\nHow can I help you with this workbook?")
        self._chat.add_message("\n".join(msg_parts), "assistant")

        # Start a fresh conversation for the new file
        self._conversation.clear()
        self._conversation.workbook_path = wb.file_path

    def _on_reader_error(self, err: str) -> None:
        self._chat.add_message(f"**Failed to load workbook:**\n\n{err}", "system")
        self._status_file_lbl.setText("  Load failed")

    # ── Chat flow ──────────────────────────────────────────────────────────────

    def _on_user_message(self, text: str) -> None:
        if self._wb is None:
            self._chat.add_message(
                "Please load an Excel file first.", "system"
            )
            return

        self._chat.add_message(text, "user")
        self._conversation.add_user(text)

        self._chat.set_input_enabled(False)
        self._chat.show_typing()

        config = SettingsDialog.load_context_config()
        context = build_context(self._wb, config)
        model = self._model_combo.currentText()

        self._ai_worker = AIWorker(self._conversation.api_messages(), context, model)
        self._ai_worker.finished.connect(self._on_ai_response)
        self._ai_worker.error.connect(self._on_ai_error)
        self._ai_worker.start()

    def _on_ai_response(self, response: AIResponse) -> None:
        self._chat.hide_typing()
        self._chat.set_input_enabled(True)

        # Add AI message to conversation history
        self._conversation.add_assistant(response.message, response)

        if response.changes:
            # Show message first, then open preview
            self._chat.add_message(response.message, "assistant")
            self._pending_response = response
            QTimer.singleShot(200, self._show_preview)
        else:
            self._chat.add_message(response.message, "assistant")

    def _on_ai_error(self, err: str) -> None:
        self._chat.hide_typing()
        self._chat.set_input_enabled(True)
        self._chat.add_message(f"**AI Error:**\n\n{err}", "system")

    # ── Preview & Apply ────────────────────────────────────────────────────────

    def _show_preview(self) -> None:
        if not self._pending_response or not self._wb:
            return

        vba_map = {m.name: m.code for m in self._wb.vba_modules}

        dlg = PreviewDialog(self._pending_response, vba_map, self)
        dlg.approved.connect(self._on_apply_approved)
        dlg.declined.connect(self._on_apply_declined)
        dlg.revise_requested.connect(self._on_revise_requested)
        dlg.exec()

    def _on_apply_approved(self) -> None:
        if not self._pending_response or not self._wb:
            return

        changes = self._pending_response.changes
        path = self._wb.file_path
        max_backups = SettingsDialog.load_max_backups()

        self._chat.add_message("_Applying changes and creating backup..._", "system")
        self._chat.set_input_enabled(False)

        self._writer_worker = ExcelWriterWorker(path, changes, max_backups)
        self._writer_worker.finished.connect(self._on_write_done)
        self._writer_worker.error.connect(self._on_write_error)
        self._writer_worker.start()

    def _on_apply_declined(self) -> None:
        self._chat.add_message("_Changes declined._", "system")
        self._pending_response = None

    def _on_revise_requested(self, note: str) -> None:
        self._pending_response = None
        self._on_user_message(f"Please revise the plan: {note}")

    def _on_write_done(self, backup_path: str) -> None:
        self._chat.set_input_enabled(True)
        bname = os.path.basename(backup_path)
        self._chat.add_message(
            f"Changes applied successfully.\nBackup saved as `{bname}`.\n\n"
            "_Re-reading workbook to update context..._",
            "system",
        )
        self._pending_response = None
        # Re-read to update context
        if self._wb:
            self.load_file(self._wb.file_path)

    def _on_write_error(self, err: str) -> None:
        self._chat.set_input_enabled(True)
        self._chat.add_message(f"**Failed to apply changes:**\n\n{err}", "system")
        self._pending_response = None

    # ── New Chat ───────────────────────────────────────────────────────────────

    def _new_chat(self) -> None:
        self._conversation.clear()
        # Re-add welcome
        self._chat.add_message("_New conversation started._", "system")

    # ── Settings ──────────────────────────────────────────────────────────────

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self)
        if dlg.exec():
            # Theme may have changed
            new_dark = SettingsDialog.load_dark_mode()
            if new_dark != self._dark_mode:
                self._dark_mode = new_dark
                self._apply_current_theme()

    def _on_model_changed(self, model: str) -> None:
        QSettings().setValue("ai/model", model)

    # ── Theme ──────────────────────────────────────────────────────────────────

    def _apply_current_theme(self) -> None:
        apply_theme(QApplication.instance(), self._dark_mode)
        self._theme_btn.setText("Light" if self._dark_mode else "Dark")
        self._wb_panel.set_accent("#89b4fa" if self._dark_mode else "#1e66f5")

    def _toggle_theme(self) -> None:
        self._dark_mode = not self._dark_mode
        QSettings().setValue("appearance/dark_mode", self._dark_mode)
        self._apply_current_theme()

    # ── Status bar ─────────────────────────────────────────────────────────────

    def _update_status(self) -> None:
        if self._wb:
            self._status_file_lbl.setText(
                f"  {self._wb.name}  |  "
                f"{len(self._wb.sheets)} sheets  |  "
                f"{len(self._wb.vba_modules)} VBA modules"
            )
            config = SettingsDialog.load_context_config()
            tokens = estimate_tokens(build_context(self._wb, config))
            self._status_token_lbl.setText(f"~{tokens:,} ctx tokens  |  {self._model_combo.currentText()}  ")
        else:
            self._status_file_lbl.setText("  No file loaded")
            self._status_token_lbl.setText("")

    # ── VBA Viewer ─────────────────────────────────────────────────────────────

    def _show_vba_viewer(self, name: str, code: str) -> None:
        dlg = VBAViewerDialog(name, code, self)
        dlg.exec()

    # ── Drag & Drop ────────────────────────────────────────────────────────────

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:  # type: ignore[override]
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
                self.load_file(path)
                event.acceptProposedAction()
                return
        event.ignore()
