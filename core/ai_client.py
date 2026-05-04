# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""OpenAI API client for PyVBAai.

Reads OPENAI_API_KEY from the environment.
The system prompt instructs GPT to always reply with JSON containing
'message', 'changes', and 'diff_summary' keys.
"""
from __future__ import annotations

import json
import os
from typing import Any

from models.conversation import AIResponse, Change

# ── System Prompt ─────────────────────────────────────────────────────────────
_SYSTEM_PROMPT = """\
You are PyVBAai, an expert Excel/VBA developer assistant embedded in a Windows desktop application.
You receive context about an open Excel workbook and help the user modify it through natural language.

## Context Format
Workbook context is provided inside <workbook_context> tags using this compact notation:
- SHEETS: sheet list with row×col dimensions
- CELLS: SheetName section, each non-empty cell as ADDR=value or ADDR={=formula}
- VBA: module source code sections
- NAMED RANGES: workbook-level names

## Response Format
You MUST ALWAYS respond with valid JSON and nothing else:
{
  "message": "Friendly explanation of what you are doing or answering",
  "changes": [],
  "diff_summary": ""
}

## Available Change Operations
### Cell / Data
{"type":"set_cell",      "sheet":"Name", "cell":"A1", "value": <any>}
{"type":"set_cell",      "sheet":"Name", "cell":"A1", "formula": "=SUM(B1:B10)"}
{"type":"set_range",     "sheet":"Name", "range":"A1:C3", "values":[[r1c1,r1c2],[r2c1,r2c2]]}
{"type":"clear_range",   "sheet":"Name", "range":"A1:Z100"}

### Sheets
{"type":"add_sheet",     "name":"NewSheet", "position":1}
{"type":"delete_sheet",  "name":"SheetName"}
{"type":"rename_sheet",  "old_name":"Old",  "new_name":"New"}
{"type":"move_sheet",    "name":"Sheet",    "position":2}
{"type":"copy_sheet",    "source":"Sheet1", "dest":"Sheet1_Copy", "position":2}

### VBA
{"type":"set_vba",           "module":"ModuleName", "code":"Sub Foo()\\n  ...\\nEnd Sub"}
{"type":"add_vba_module",    "name":"NewModule",    "code":"..."}
{"type":"delete_vba_module", "name":"ModuleName"}

### Named Ranges
{"type":"add_named_range",    "name":"MyRange",  "refers_to":"=Sheet1!$A$1:$D$10"}
{"type":"delete_named_range", "name":"MyRange"}

## Rules
1. Sheet names and addresses must exactly match the context (case-sensitive).
2. Use standard Excel formula syntax (= prefix).
3. VBA code must be syntactically complete; escape newlines as \\n in JSON strings.
4. If the user asks a question or the action requires no changes, set changes to [].
5. Be concise in 'message'. Use 'diff_summary' for a bullet-point list of changes.
6. If the context is truncated, acknowledge this and ask if more detail is needed.
7. Never invent sheet names or cell references not present in the context.
"""


class AIClient:
    def __init__(self, model: str = "gpt-4o") -> None:
        self.model = model
        self._client = None

    @property
    def client(self):
        if self._client is None:
            from openai import OpenAI
            api_key = os.environ.get("OPENAI_API_KEY", "").strip()
            if not api_key:
                raise ValueError(
                    "OPENAI_API_KEY environment variable is not set.\n"
                    "Set it in Windows Environment Variables and restart the application."
                )
            self._client = OpenAI(api_key=api_key)
        return self._client

    def send(
        self,
        conversation_messages: list[dict[str, str]],
        context: str,
    ) -> AIResponse:
        """
        Send the conversation to GPT and return a parsed AIResponse.

        conversation_messages: list of {"role": "user"|"assistant", "content": str}
        context: compact workbook context string
        """
        full_messages: list[dict[str, str]] = [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {
                "role": "user",
                "content": f"<workbook_context>\n{context}\n</workbook_context>",
            },
            {
                "role": "assistant",
                "content": (
                    '{"message": "Context loaded. How can I help with this workbook?", '
                    '"changes": [], "diff_summary": ""}'
                ),
            },
            *conversation_messages,
        ]

        response = self.client.chat.completions.create(
            model=self.model,
            messages=full_messages,
            response_format={"type": "json_object"},
            temperature=0.1,
        )

        raw_json = response.choices[0].message.content or "{}"
        return _parse_response(raw_json)

    @staticmethod
    def available_models() -> list[str]:
        return ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-4", "gpt-3.5-turbo"]


# ── response parsing ─────────────────────────────────────────────────────────

def _parse_response(raw_json: str) -> AIResponse:
    try:
        data: dict[str, Any] = json.loads(raw_json)
    except json.JSONDecodeError as exc:
        return AIResponse(
            message=f"[Parse error] Could not decode AI response: {exc}\n\nRaw:\n{raw_json}",
            raw_json=raw_json,
        )

    message = str(data.get("message", ""))
    diff_summary = str(data.get("diff_summary", ""))

    changes: list[Change] = []
    for item in data.get("changes", []):
        if isinstance(item, dict) and "type" in item:
            op_type = item.pop("type")
            changes.append(Change(type=op_type, params=item))

    return AIResponse(
        message=message,
        changes=changes,
        diff_summary=diff_summary,
        raw_json=raw_json,
    )
