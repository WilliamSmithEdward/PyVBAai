# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""QThread workers for non-blocking Excel I/O and AI API calls."""
from __future__ import annotations

import os

from PySide6.QtCore import QThread, Signal

from app.logger import get_logger
from models.conversation import Change
from models.workbook import ContextConfig

_log = get_logger(__name__)


class ExcelReaderWorker(QThread):
    """Read an Excel workbook on a background thread."""

    finished = Signal(object)   # WorkbookData
    error = Signal(str)

    def __init__(self, file_path: str, config: ContextConfig | None = None) -> None:
        super().__init__()
        self._path = file_path
        self._config = config or ContextConfig()

    def run(self) -> None:
        try:
            from core.excel_reader import read_workbook
            _log.debug("ExcelReaderWorker: reading %s", self._path)
            result = read_workbook(self._path, self._config)
            _log.debug(
                "ExcelReaderWorker: loaded %d sheets, error=%s",
                len(result.sheets), result.extraction_error,
            )
            self.finished.emit(result)
        except Exception as exc:
            _log.exception("ExcelReaderWorker error: %s", exc)
            self.error.emit(str(exc))


class AIWorker(QThread):
    """Call the OpenAI API on a background thread."""

    finished = Signal(object)   # AIResponse
    error = Signal(str)

    def __init__(
        self,
        conversation_messages: list[dict],
        context: str,
        model: str = "gpt-4o",
    ) -> None:
        super().__init__()
        self._messages = conversation_messages
        self._context = context
        self._model = model

    def run(self) -> None:
        try:
            from core.ai_client import AIClient
            _log.debug("AIWorker: sending %d messages, model=%s", len(self._messages), self._model)
            client = AIClient(model=self._model)
            result = client.send(self._messages, self._context)
            _log.debug(
                "AIWorker: response received, %d changes, message=%r",
                len(result.changes), result.message[:120] if result.message else "",
            )
            self.finished.emit(result)
        except Exception as exc:
            _log.exception("AIWorker error: %s", exc)
            self.error.emit(str(exc))


class ExcelWriterWorker(QThread):
    """Apply changes to an Excel workbook on a background thread."""

    finished = Signal(str, str)  # backup path, saved path (may differ if .xlsx→.xlsm)
    error = Signal(str)

    def __init__(self, file_path: str, changes: list[Change], max_backups: int = 20) -> None:
        super().__init__()
        self._path = file_path
        self._changes = changes
        self._max_backups = max_backups

    def run(self) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize()

            from core.backup_manager import create_backup
            from core.excel_writer import apply_changes

            _log.debug(
                "ExcelWriterWorker: applying %d changes to %s",
                len(self._changes), self._path,
            )
            for i, c in enumerate(self._changes):
                _log.debug("  change[%d]: type=%s params=%s", i, c.type, c.params)
            backup_path = create_backup(self._path, self._max_backups)
            _log.debug("ExcelWriterWorker: backup created at %s", backup_path)
            saved_path = apply_changes(self._path, self._changes)
            _log.debug("ExcelWriterWorker: saved to %s", saved_path)
            self.finished.emit(backup_path, saved_path)

        except Exception as exc:
            _log.exception("ExcelWriterWorker error: %s", exc)
            self.error.emit(str(exc))
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass


class NewWorkbookWorker(QThread):
    """Create a blank Excel workbook on a background thread."""

    finished = Signal(str)  # saved file path
    error = Signal(str)

    def __init__(self, file_path: str) -> None:
        super().__init__()
        self._path = file_path

    def run(self) -> None:
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        excel = None
        we_launched = False
        try:
            try:
                excel = win32.GetActiveObject("Excel.Application")
            except Exception:
                excel = win32.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                we_launched = True

            abs_path = os.path.abspath(os.path.normpath(self._path))
            wb = excel.Workbooks.Add()
            excel.DisplayAlerts = False
            wb.SaveAs(abs_path, 51)  # FileFormat 51 = xlOpenXMLWorkbook
            wb.Close(SaveChanges=False)
            self.finished.emit(abs_path)

        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            try:
                if we_launched and excel and excel.Workbooks.Count == 0:
                    excel.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
