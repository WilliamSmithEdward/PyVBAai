# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Conversation and AI response models."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


@dataclass
class Change:
    """A single change operation to apply to the workbook."""
    type: str
    params: dict[str, Any]


@dataclass
class AIResponse:
    """Parsed response from the AI."""
    message: str
    changes: list[Change] = field(default_factory=list)
    diff_summary: str = ""
    raw_json: str | None = None


@dataclass
class Message:
    role: str          # "user" | "assistant" | "system"
    content: str
    timestamp: datetime = field(default_factory=datetime.now)
    ai_response: AIResponse | None = None   # populated for assistant messages
    applied: bool = False                       # whether changes were applied


@dataclass
class Conversation:
    messages: list[Message] = field(default_factory=list)
    workbook_path: str | None = None

    def add_user(self, content: str) -> Message:
        msg = Message(role="user", content=content)
        self.messages.append(msg)
        return msg

    def add_assistant(self, content: str, ai_response: AIResponse | None = None) -> Message:
        msg = Message(role="assistant", content=content, ai_response=ai_response)
        self.messages.append(msg)
        return msg

    def add_system(self, content: str) -> Message:
        msg = Message(role="system", content=content)
        self.messages.append(msg)
        return msg

    def api_messages(self) -> list[dict[str, str]]:
        """Return only user/assistant messages in OpenAI API format."""
        return [
            {"role": m.role, "content": m.content}
            for m in self.messages
            if m.role in ("user", "assistant")
        ]

    def clear(self) -> None:
        self.messages.clear()
