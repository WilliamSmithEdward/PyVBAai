# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for core/ai_client.py — focuses on pure Python logic (no OpenAI calls)."""
from __future__ import annotations

import json

import pytest

from core.ai_client import AIClient, _parse_response
from models.conversation import AIResponse

# ── _parse_response ───────────────────────────────────────────────────────────

class TestParseResponse:
    def test_basic_message(self):
        raw = json.dumps({"message": "Hello", "changes": [], "diff_summary": ""})
        r = _parse_response(raw)
        assert isinstance(r, AIResponse)
        assert r.message == "Hello"
        assert r.changes == []
        assert r.diff_summary == ""

    def test_single_change_parsed(self):
        raw = json.dumps({
            "message": "Updated A1",
            "changes": [{"type": "set_cell", "sheet": "Sheet1", "cell": "A1", "value": 99}],
            "diff_summary": "- Set A1=99",
        })
        r = _parse_response(raw)
        assert len(r.changes) == 1
        assert r.changes[0].type == "set_cell"
        assert r.changes[0].params == {"sheet": "Sheet1", "cell": "A1", "value": 99}

    def test_type_removed_from_params(self):
        raw = json.dumps({
            "message": "x",
            "changes": [{"type": "add_sheet", "name": "NewSheet", "position": 1}],
            "diff_summary": "",
        })
        r = _parse_response(raw)
        assert "type" not in r.changes[0].params
        assert r.changes[0].params == {"name": "NewSheet", "position": 1}

    def test_multiple_changes(self):
        changes = [
            {"type": "set_cell", "sheet": "S1", "cell": "A1", "value": 1},
            {"type": "set_cell", "sheet": "S1", "cell": "B1", "value": 2},
            {"type": "add_sheet", "name": "NewSheet"},
        ]
        raw = json.dumps({"message": "done", "changes": changes, "diff_summary": ""})
        r = _parse_response(raw)
        assert len(r.changes) == 3
        assert r.changes[2].type == "add_sheet"

    def test_invalid_json_returns_error_message(self):
        r = _parse_response("not valid json {{{")
        assert "[Parse error]" in r.message
        assert r.changes == []
        assert r.raw_json == "not valid json {{{"

    def test_empty_json_object(self):
        r = _parse_response("{}")
        assert r.message == ""
        assert r.changes == []
        assert r.diff_summary == ""

    def test_raw_json_stored(self):
        raw = '{"message":"Hi","changes":[],"diff_summary":""}'
        r = _parse_response(raw)
        assert r.raw_json == raw

    def test_missing_changes_key(self):
        raw = json.dumps({"message": "No changes key"})
        r = _parse_response(raw)
        assert r.changes == []

    def test_changes_with_missing_type_skipped(self):
        raw = json.dumps({
            "message": "x",
            "changes": [{"sheet": "S1", "cell": "A1"}],  # no "type"
            "diff_summary": "",
        })
        r = _parse_response(raw)
        assert r.changes == []

    def test_changes_non_dict_items_skipped(self):
        raw = json.dumps({
            "message": "x",
            "changes": ["invalid", 42, None],
            "diff_summary": "",
        })
        r = _parse_response(raw)
        assert r.changes == []

    def test_diff_summary_preserved(self):
        raw = json.dumps({
            "message": "done",
            "changes": [],
            "diff_summary": "- Added sheet\n- Set A1",
        })
        r = _parse_response(raw)
        assert r.diff_summary == "- Added sheet\n- Set A1"

    @pytest.mark.parametrize("op_type", [
        "set_cell", "set_range", "clear_range",
        "add_sheet", "delete_sheet", "rename_sheet", "move_sheet", "copy_sheet",
        "set_vba", "add_vba_module", "delete_vba_module",
        "add_named_range", "delete_named_range",
    ])
    def test_all_operation_types_parsed(self, op_type):
        raw = json.dumps({
            "message": "ok",
            "changes": [{"type": op_type, "extra_param": "value"}],
            "diff_summary": "",
        })
        r = _parse_response(raw)
        assert len(r.changes) == 1
        assert r.changes[0].type == op_type


# ── AIClient ──────────────────────────────────────────────────────────────────

class TestAIClient:
    def test_default_model(self):
        c = AIClient()
        assert c.model == "gpt-4o"

    def test_custom_model(self):
        c = AIClient(model="gpt-4o-mini")
        assert c.model == "gpt-4o-mini"

    def test_available_models_returns_list(self, monkeypatch):
        """fetch_models_from_api should return a sorted list when the API responds."""
        import json
        import io
        import core.ai_client as ai_mod

        fake_payload = json.dumps({
            "object": "list",
            "data": [
                {"id": "gpt-4o", "object": "model"},
                {"id": "gpt-4o-mini", "object": "model"},
                {"id": "gpt-4", "object": "model"},
            ],
        }).encode()

        class _FakeResp:
            def read(self):
                return fake_payload
            def __enter__(self):
                return self
            def __exit__(self, *_):
                pass

        monkeypatch.setattr("urllib.request.urlopen", lambda *a, **kw: _FakeResp())
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        ai_mod._cached_models = None  # reset cache

        models = AIClient.fetch_models_from_api()
        assert isinstance(models, list)
        assert len(models) > 0

    def test_available_models_contains_gpt4o(self, monkeypatch):
        import json
        import core.ai_client as ai_mod

        fake_payload = json.dumps({
            "object": "list",
            "data": [{"id": "gpt-4o"}, {"id": "gpt-4o-mini"}],
        }).encode()

        class _FakeResp:
            def read(self):
                return fake_payload
            def __enter__(self):
                return self
            def __exit__(self, *_):
                pass

        monkeypatch.setattr("urllib.request.urlopen", lambda *a, **kw: _FakeResp())
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        ai_mod._cached_models = None

        assert "gpt-4o" in AIClient.fetch_models_from_api()

    def test_available_models_contains_gpt4o_mini(self, monkeypatch):
        import json
        import core.ai_client as ai_mod

        fake_payload = json.dumps({
            "object": "list",
            "data": [{"id": "gpt-4o"}, {"id": "gpt-4o-mini"}],
        }).encode()

        class _FakeResp:
            def read(self):
                return fake_payload
            def __enter__(self):
                return self
            def __exit__(self, *_):
                pass

        monkeypatch.setattr("urllib.request.urlopen", lambda *a, **kw: _FakeResp())
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        ai_mod._cached_models = None

        assert "gpt-4o-mini" in AIClient.fetch_models_from_api()

    def test_available_models_returns_empty_without_key(self, monkeypatch):
        import core.ai_client as ai_mod
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        ai_mod._cached_models = None
        assert AIClient.fetch_models_from_api() == []

    def test_available_models_returns_empty_on_network_error(self, monkeypatch):
        import urllib.error
        import core.ai_client as ai_mod
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        monkeypatch.setattr(
            "urllib.request.urlopen",
            lambda *a, **kw: (_ for _ in ()).throw(urllib.error.URLError("timeout")),
        )
        ai_mod._cached_models = None
        assert AIClient.fetch_models_from_api() == []

    def test_available_models_filters_non_gpt(self, monkeypatch):
        """Non-gpt- prefixed models should be excluded."""
        import json
        import core.ai_client as ai_mod

        fake_payload = json.dumps({
            "object": "list",
            "data": [{"id": "gpt-4o"}, {"id": "dall-e-3"}, {"id": "whisper-1"}],
        }).encode()

        class _FakeResp:
            def read(self):
                return fake_payload
            def __enter__(self):
                return self
            def __exit__(self, *_):
                pass

        monkeypatch.setattr("urllib.request.urlopen", lambda *a, **kw: _FakeResp())
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        ai_mod._cached_models = None

        models = AIClient.fetch_models_from_api()
        assert models == ["gpt-4o"]

    def test_available_models_filters_dated_snapshots(self, monkeypatch):
        """Versioned snapshot models (date suffixes) should be excluded."""
        import json
        import core.ai_client as ai_mod

        fake_payload = json.dumps({
            "object": "list",
            "data": [
                {"id": "gpt-4o"},
                {"id": "gpt-4o-mini"},
                {"id": "gpt-4o-2024-05-13"},       # full date suffix
                {"id": "gpt-4-0314"},               # short date suffix
                {"id": "gpt-4-0613"},
                {"id": "gpt-3.5-turbo-0301"},
                {"id": "gpt-4-turbo-2024-04-09"},
                {"id": "gpt-4-turbo"},              # stable alias - keep
            ],
        }).encode()

        class _FakeResp:
            def read(self):
                return fake_payload
            def __enter__(self):
                return self
            def __exit__(self, *_):
                pass

        monkeypatch.setattr("urllib.request.urlopen", lambda *a, **kw: _FakeResp())
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        ai_mod._cached_models = None

        models = AIClient.fetch_models_from_api()
        assert models == ["gpt-4-turbo", "gpt-4o", "gpt-4o-mini"]

    def test_available_models_caches_result(self, monkeypatch):
        """Second call should return cached result without hitting the network."""
        import json
        import core.ai_client as ai_mod

        call_count = [0]

        def _fake_urlopen(*a, **kw):
            call_count[0] += 1
            class _R:
                def read(self):
                    return json.dumps({"object": "list", "data": [{"id": "gpt-4o"}]}).encode()
                def __enter__(self):
                    return self
                def __exit__(self, *_):
                    pass
            return _R()

        monkeypatch.setattr("urllib.request.urlopen", _fake_urlopen)
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        ai_mod._cached_models = None

        AIClient.fetch_models_from_api()
        AIClient.fetch_models_from_api()
        assert call_count[0] == 1  # second call used cache

    def test_client_not_created_without_call(self):
        """Client should be lazily initialized."""
        c = AIClient()
        assert c._client is None

    def test_client_raises_without_api_key(self, monkeypatch):
        """Accessing .client without OPENAI_API_KEY should raise ValueError."""
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        c = AIClient()
        with pytest.raises(ValueError, match="OPENAI_API_KEY"):
            _ = c.client

    def test_send_uses_parse_response(self, monkeypatch):
        """send() should call the OpenAI client and parse the JSON response."""
        fake_json = json.dumps({
            "message": "Done",
            "changes": [{"type": "set_cell", "sheet": "S1", "cell": "A1", "value": 5}],
            "diff_summary": "- Set A1=5",
        })

        # Mock the OpenAI client
        class FakeChoice:
            class message:
                content = fake_json

        class FakeCompletion:
            choices = [FakeChoice()]

        class FakeOpenAIClient:
            class chat:
                class completions:
                    @staticmethod
                    def create(**kwargs):
                        return FakeCompletion()

        monkeypatch.setenv("OPENAI_API_KEY", "sk-fake-key-for-testing")
        client = AIClient(model="gpt-4o")
        # Inject the fake OpenAI client directly
        client._client = FakeOpenAIClient()

        result = client.send([{"role": "user", "content": "Update A1 to 5"}], "context text")
        assert result.message == "Done"
        assert len(result.changes) == 1
        assert result.changes[0].type == "set_cell"
