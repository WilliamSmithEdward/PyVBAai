# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for models/conversation.py dataclasses."""
from __future__ import annotations

from datetime import datetime

import pytest

from models.conversation import AIResponse, Change, Conversation, Message

# ── Change ────────────────────────────────────────────────────────────────────

class TestChange:
    def test_fields(self):
        c = Change(type="set_cell", params={"sheet": "Sheet1", "cell": "A1", "value": 42})
        assert c.type == "set_cell"
        assert c.params["sheet"] == "Sheet1"
        assert c.params["value"] == 42

    def test_empty_params(self):
        c = Change(type="add_sheet", params={})
        assert c.params == {}

    @pytest.mark.parametrize("op_type", [
        "set_cell", "set_range", "clear_range",
        "add_sheet", "delete_sheet", "rename_sheet", "move_sheet", "copy_sheet",
        "set_vba", "add_vba_module", "delete_vba_module",
        "add_named_range", "delete_named_range",
    ])
    def test_all_valid_types(self, op_type):
        c = Change(type=op_type, params={})
        assert c.type == op_type


# ── AIResponse ────────────────────────────────────────────────────────────────

class TestAIResponse:
    def test_defaults(self):
        r = AIResponse(message="Hello")
        assert r.changes == []
        assert r.diff_summary == ""
        assert r.raw_json is None

    def test_with_changes(self, ai_response_with_changes):
        r = ai_response_with_changes
        assert r.message == "I updated cell A1."
        assert len(r.changes) == 1
        assert r.changes[0].type == "set_cell"
        assert r.changes[0].params["value"] == 99

    def test_no_changes(self):
        r = AIResponse(message="Just a question answer.")
        assert r.changes == []

    def test_raw_json_stored(self):
        raw = '{"message":"Hi","changes":[],"diff_summary":""}'
        r = AIResponse(message="Hi", raw_json=raw)
        assert r.raw_json == raw


# ── Message ───────────────────────────────────────────────────────────────────

class TestMessage:
    def test_user_message(self):
        m = Message(role="user", content="Hello AI")
        assert m.role == "user"
        assert m.content == "Hello AI"
        assert m.applied is False
        assert m.ai_response is None

    def test_assistant_message_with_response(self, ai_response_with_changes):
        m = Message(role="assistant", content="Done.", ai_response=ai_response_with_changes)
        assert m.role == "assistant"
        assert m.ai_response is ai_response_with_changes

    def test_timestamp_is_set(self):
        before = datetime.now()
        m = Message(role="user", content="test")
        after = datetime.now()
        assert before <= m.timestamp <= after

    def test_applied_flag(self):
        m = Message(role="assistant", content="Applied.", applied=True)
        assert m.applied is True


# ── Conversation ──────────────────────────────────────────────────────────────

class TestConversation:
    def test_starts_empty(self, conversation):
        assert conversation.messages == []
        assert conversation.workbook_path is None

    def test_add_user(self, conversation):
        msg = conversation.add_user("Hi")
        assert msg.role == "user"
        assert msg.content == "Hi"
        assert len(conversation.messages) == 1

    def test_add_assistant(self, conversation):
        msg = conversation.add_assistant("Hello back")
        assert msg.role == "assistant"
        assert len(conversation.messages) == 1

    def test_add_assistant_with_ai_response(self, conversation, ai_response_with_changes):
        msg = conversation.add_assistant("Done.", ai_response=ai_response_with_changes)
        assert msg.ai_response is ai_response_with_changes

    def test_add_system(self, conversation):
        msg = conversation.add_system("System context")
        assert msg.role == "system"
        assert len(conversation.messages) == 1

    def test_message_ordering(self, conversation):
        conversation.add_user("first")
        conversation.add_assistant("second")
        conversation.add_user("third")
        assert [m.content for m in conversation.messages] == ["first", "second", "third"]

    def test_api_messages_excludes_system(self, conversation):
        conversation.add_system("Ignore me")
        conversation.add_user("user msg")
        conversation.add_assistant("ai msg")
        api = conversation.api_messages()
        assert len(api) == 2
        assert api[0] == {"role": "user", "content": "user msg"}
        assert api[1] == {"role": "assistant", "content": "ai msg"}

    def test_api_messages_format(self, conversation):
        conversation.add_user("Q")
        conversation.add_assistant("A")
        api = conversation.api_messages()
        for item in api:
            assert set(item.keys()) == {"role", "content"}

    def test_api_messages_empty(self, conversation):
        assert conversation.api_messages() == []

    def test_clear(self, conversation):
        conversation.add_user("a")
        conversation.add_user("b")
        conversation.clear()
        assert conversation.messages == []

    def test_workbook_path(self):
        c = Conversation(workbook_path="C:/test/book.xlsx")
        assert c.workbook_path == "C:/test/book.xlsx"

    def test_multiple_rounds(self, conversation):
        for i in range(5):
            conversation.add_user(f"question {i}")
            conversation.add_assistant(f"answer {i}")
        assert len(conversation.messages) == 10
        assert len(conversation.api_messages()) == 10
