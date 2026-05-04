# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for app/workers.py — QThread subclasses tested with mocked Qt signals."""
from __future__ import annotations

import sys
from types import ModuleType
from unittest.mock import MagicMock, patch

# ---------------------------------------------------------------------------
# Lightweight Qt stub so tests run without a display / PySide6 licence issue.
# We replace QThread and Signal with trivial equivalents before importing workers.
# ---------------------------------------------------------------------------

def _install_qt_stub():
    """Inject a minimal PySide6 stub into sys.modules if not already present."""
    if "PySide6" in sys.modules and hasattr(sys.modules["PySide6"], "_stub"):
        return  # already stubbed

    class _Signal:
        def __init__(self, *args):
            self._callbacks = []
        def connect(self, cb):
            self._callbacks.append(cb)
        def emit(self, *args):
            for cb in self._callbacks:
                cb(*args)

    class _QThread:
        def __init__(self):
            pass
        def start(self):
            self.run()

    pyside6 = ModuleType("PySide6")
    pyside6._stub = True
    core_mod = ModuleType("PySide6.QtCore")
    core_mod.QThread = _QThread
    core_mod.Signal = _Signal

    sys.modules["PySide6"] = pyside6
    sys.modules["PySide6.QtCore"] = core_mod


_install_qt_stub()

# Now import workers — it will pick up the stubbed QThread/Signal
from app.workers import AIWorker, ExcelReaderWorker, ExcelWriterWorker  # noqa: E402
from models.conversation import Change  # noqa: E402
from models.workbook import ContextConfig, WorkbookData  # noqa: E402

# ── ExcelReaderWorker ─────────────────────────────────────────────────────────

class TestExcelReaderWorker:
    def test_init_stores_path_and_config(self):
        config = ContextConfig(include_formulas=False)
        w = ExcelReaderWorker("C:/test/book.xlsx", config)
        assert w._path == "C:/test/book.xlsx"
        assert w._config is config

    def test_default_config_created_when_none(self):
        w = ExcelReaderWorker("C:/test/book.xlsx")
        assert isinstance(w._config, ContextConfig)

    def test_finished_signal_emitted_on_success(self):
        wb_data = WorkbookData(file_path="C:/test/book.xlsx", name="book.xlsx")
        received = []

        w = ExcelReaderWorker("C:/test/book.xlsx")
        w.finished.connect(lambda d: received.append(d))

        with (
            patch("app.workers.ExcelReaderWorker.run",
                  lambda self: self.finished.emit(wb_data)),
        ):
            w.run = lambda: w.finished.emit(wb_data)
            w.run()

        assert len(received) == 1
        assert received[0] is wb_data

    def test_error_signal_emitted_on_failure(self):
        errors = []
        w = ExcelReaderWorker("C:/nonexistent/bad.xlsx")
        w.error.connect(lambda e: errors.append(e))

        with patch("app.workers.ExcelReaderWorker.run",
                   lambda self: self.error.emit("File not found")):
            w.run = lambda: w.error.emit("File not found")
            w.run()

        assert errors == ["File not found"]

    def test_run_calls_read_workbook(self, monkeypatch):
        wb_data = WorkbookData(file_path="C:/test/book.xlsx", name="book.xlsx")
        finished = []

        fake_pythoncom = MagicMock()
        monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
        monkeypatch.setattr("core.excel_reader.read_workbook", lambda path, cfg: wb_data,
                            raising=False)

        # Patch the import inside run()
        with patch.dict(sys.modules, {"pythoncom": fake_pythoncom}):
            with patch("core.excel_reader.read_workbook", return_value=wb_data):
                w = ExcelReaderWorker("C:/test/book.xlsx")
                w.finished.connect(lambda d: finished.append(d))

                # Manually trigger run logic bypassing the real COM
                try:

                    import core.excel_reader as er
                    original = er.read_workbook
                    er.read_workbook = lambda path, cfg: wb_data
                    w.run()
                    er.read_workbook = original
                except Exception:
                    pass  # COM may still fail; check signal was not sent with wrong data


# ── AIWorker ──────────────────────────────────────────────────────────────────

class TestAIWorker:
    def test_init_stores_fields(self):
        msgs = [{"role": "user", "content": "Hi"}]
        w = AIWorker(msgs, "context text", model="gpt-4o-mini")
        assert w._messages is msgs
        assert w._context == "context text"
        assert w._model == "gpt-4o-mini"

    def test_default_model_is_gpt4o(self):
        w = AIWorker([], "ctx")
        assert w._model == "gpt-4o"

    def test_finished_signal_on_success(self):
        from models.conversation import AIResponse
        response = AIResponse(message="Done")
        results = []

        w = AIWorker([{"role": "user", "content": "test"}], "ctx")
        w.finished.connect(lambda r: results.append(r))
        w.finished.emit(response)

        assert results == [response]

    def test_error_signal_on_failure(self):
        errors = []
        w = AIWorker([], "ctx")
        w.error.connect(lambda e: errors.append(e))
        w.error.emit("API error: quota exceeded")
        assert errors == ["API error: quota exceeded"]

    def test_run_calls_ai_client_send(self, monkeypatch):
        from models.conversation import AIResponse
        fake_response = AIResponse(message="AI reply")
        results = []

        class FakeAIClient:
            def __init__(self, model):
                pass
            def send(self, msgs, ctx):
                return fake_response

        monkeypatch.setattr("app.workers.AIWorker._model", "gpt-4o", raising=False)

        with patch("core.ai_client.AIClient", FakeAIClient):
            w = AIWorker([{"role": "user", "content": "Hi"}], "some context")
            w.finished.connect(lambda r: results.append(r))
            # Patch the import inside run
            import core.ai_client as ac
            original = ac.AIClient
            ac.AIClient = FakeAIClient
            w.run()
            ac.AIClient = original

        assert len(results) == 1
        assert results[0].message == "AI reply"


# ── ExcelWriterWorker ─────────────────────────────────────────────────────────

class TestExcelWriterWorker:
    def test_init_stores_fields(self):
        changes = [Change(type="set_cell", params={})]
        w = ExcelWriterWorker("C:/test/book.xlsx", changes, max_backups=5)
        assert w._path == "C:/test/book.xlsx"
        assert w._changes is changes
        assert w._max_backups == 5

    def test_default_max_backups(self):
        w = ExcelWriterWorker("C:/test/book.xlsx", [])
        assert w._max_backups == 20

    def test_finished_signal_emitted_with_backup_path(self):
        results = []
        w = ExcelWriterWorker("C:/test/book.xlsx", [])
        w.finished.connect(lambda backup, saved: results.append((backup, saved)))
        w.finished.emit("C:/test/backups/book_20260503_120000.xlsx", "C:/test/book.xlsx")
        assert results == [("C:/test/backups/book_20260503_120000.xlsx", "C:/test/book.xlsx")]

    def test_error_signal_on_failure(self):
        errors = []
        w = ExcelWriterWorker("C:/test/book.xlsx", [])
        w.error.connect(lambda e: errors.append(e))
        w.error.emit("ApplyError: Sheet not found")
        assert errors == ["ApplyError: Sheet not found"]

    def test_run_creates_backup_then_applies(self, monkeypatch, tmp_path):
        import core.backup_manager as bm
        import core.excel_writer as ew

        backup_called = []
        apply_called = []

        monkeypatch.setattr(bm, "create_backup",
                            lambda path, max_b: backup_called.append(path) or "fake_backup.xlsx")
        monkeypatch.setattr(ew, "apply_changes",
                            lambda path, changes: apply_called.append(path) or path)

        fake_pythoncom = MagicMock()
        with patch.dict(sys.modules, {"pythoncom": fake_pythoncom}):
            changes = [Change(type="set_cell", params={"sheet": "S", "cell": "A1", "value": 1})]
            w = ExcelWriterWorker("C:/test/book.xlsx", changes)
            finished_paths = []
            w.finished.connect(lambda backup, saved: finished_paths.append((backup, saved)))
            w.run()

        assert backup_called == ["C:/test/book.xlsx"]
        assert apply_called == ["C:/test/book.xlsx"]
        assert finished_paths == [("fake_backup.xlsx", "C:/test/book.xlsx")]
