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
from typing import Any, cast

from models.conversation import AIResponse, Change

# Cache for model list fetched from the API (populated on first call)
_cached_models: list[str] | None = None

# ── System Prompt ─────────────────────────────────────────────────────────────
_SYSTEM_PROMPT = """\
You are PyVBAai, an expert Excel/VBA developer assistant embedded in a Windows desktop application.
You receive context about an open Excel workbook and help the user modify it through natural language.

## Cell Notation (compact format — what you receive)
Cells are shown as:  R<row>: <COL>=<val>[flags]
  val types: "string" | number | {formula without leading =} | None
  flags (comma-separated inside [ ]):
    B=bold  I=italic  S=strikethrough  U=underline  W=wrap_text
    #<spec>=number_format  (e.g. #0.00%  #$#,##0  #dd/mm/yyyy)
    fn:<name>=font_name    fs:<pts>=font_size
    ^<RRGGBB>=font_color   ~<RRGGBB>=bg_color
    ha:<l|c|r>=h_align     va:<t|c|b>=v_align
    bt/bb/bl/br:<style>[:<RRGGBB>]=border top/bottom/left/right
      border styles: thin  medium  thick  dashed  dotted  double  hair
  Example:  R1: A="Revenue"[B,fn:Calibri,fs:14,bt:thin,bb:thin]  B=42000[#$#,##0,~92D050]

## Response Format
You MUST ALWAYS respond with valid JSON and nothing else:
{
  "message": "Friendly explanation of what you are doing or answering",
  "changes": [],
  "diff_summary": ""
}

## Available Change Operations

### Cell Values
{"type":"set_cell",  "sheet":"Name", "cell":"A1", "value": <any>}
{"type":"set_cell",  "sheet":"Name", "cell":"A1", "formula": "=SUM(B1:B10)"}
{"type":"set_range", "sheet":"Name", "range":"A1:C3", "values":[[r1c1,r1c2],[r2c1,r2c2]]}
{"type":"clear_range","sheet":"Name","range":"A1:Z100"}

### Inline Cell Formatting (embed flags or explicit keys in set_cell / set_range)
Both forms are equivalent — use whichever is more token-efficient:
  compact flags string:  {"type":"set_cell","sheet":"S","cell":"A1","value":"Rev","flags":"B,#$#,##0,bt:thin"}
  explicit keys:         {"type":"set_cell","sheet":"S","cell":"A1","value":"Rev","bold":true,"number_format":"$#,##0","border_top":"thin"}
set_range also accepts "flags" or explicit keys to apply uniform format to all cells in the range.

### Formatting Only (no value change)
{"type":"set_format", "sheet":"Name", "range":"A1:D1",
 "bold":true, "italic":false, "strikethrough":false, "underline":false, "wrap_text":false,
 "number_format":"$#,##0", "font_name":"Calibri", "font_size":11,
 "font_color":"FF0000", "bg_color":"D9E1F2",
 "h_align":"center", "v_align":"center",
 "border_top":"thin", "border_bottom":"medium:0070C0", "border_left":"thin", "border_right":"thin"}
All format keys are optional — only include keys you want to change.
border values: "<style>" or "<style>:<RRGGBB>"  (e.g. "thin", "medium:FF0000")

### Sheets
{"type":"add_sheet",    "name":"NewSheet", "position":1}
{"type":"delete_sheet", "name":"SheetName"}
{"type":"rename_sheet", "old_name":"Old", "new_name":"New"}
{"type":"move_sheet",   "name":"Sheet",   "position":2}
{"type":"copy_sheet",   "source":"Sheet1","dest":"Sheet1_Copy","position":2}
{"type":"hide_sheet",   "name":"Sheet"}
{"type":"unhide_sheet", "name":"Sheet"}

### Merging
{"type":"merge_cells",   "sheet":"Name", "range":"A1:D1"}
{"type":"unmerge_cells", "sheet":"Name", "range":"A1:D1"}

### VBA
{"type":"set_vba",           "module":"ModuleName", "code":"Sub Foo()\\n  ...\\nEnd Sub"}
{"type":"add_vba_module",    "name":"NewModule",    "code":"..."}
{"type":"delete_vba_module", "name":"ModuleName"}

### Rows & Columns
{"type":"insert_rows",   "sheet":"S", "row":3, "count":1}
{"type":"delete_rows",   "sheet":"S", "row":3, "count":1}
{"type":"insert_cols",   "sheet":"S", "col":"B", "count":1}
{"type":"delete_cols",   "sheet":"S", "col":"B", "count":1}
{"type":"set_col_width", "sheet":"S", "columns":"A",   "width":20}  -- columns can be "A" or "A:D"
{"type":"set_row_height","sheet":"S", "row":1,   "height":30}  -- height in points
{"type":"freeze_panes",  "sheet":"S", "cell":"B2"}  -- freeze rows above and cols left of cell; cell="" to unfreeze
{"type":"auto_filter",   "sheet":"S", "range":"A1:D1"}  -- range="" to clear
{"type":"set_tab_color", "name":"Sheet1", "color":"FF0000"}  -- color is RRGGBB; color="" to clear

### Tables (Excel ListObject)
{"type":"create_table", "sheet":"Sheet1", "range":"A1:D101", "name":"PostsTable", "style":"TableStyleMedium9"}
{"type":"delete_table", "sheet":"Sheet1", "name":"PostsTable"}
- "style" is optional (default TableStyleMedium9). Valid values: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11.
- When the user asks to create a table, list object, or structured table — always use create_table, NOT set_named_range.

### Named Ranges
{"type":"set_named_range",    "name":"MyRange",  "refers_to":"Sheet1!$A$1:$D$10"}
{"type":"delete_named_range", "name":"MyRange"}

## Rules
1. Sheet names and addresses must exactly match the context (case-sensitive).
2. Use standard Excel formula syntax (= prefix in "formula" key, no = inside { } notation).
3. VBA code must be syntactically complete; escape newlines as \\n in JSON strings.
4. If the user asks a question or the action requires no changes, set changes to [].
5. Be concise in 'message'. Use 'diff_summary' for a bullet-point list of changes made.
6. If the context is truncated, acknowledge this and ask if more detail is needed.
7. Never invent sheet names or cell references not present in the context.
8. hex colors are RRGGBB (6 chars) — do not include alpha prefix.
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

        from openai.types.chat import ChatCompletionMessageParam

        response = self.client.chat.completions.create(
            model=self.model,
            messages=cast(list[ChatCompletionMessageParam], full_messages),
            response_format={"type": "json_object"},
            temperature=0.1,
        )

        raw_json = response.choices[0].message.content or "{}"
        return _parse_response(raw_json)

    @staticmethod
    def fetch_models_from_api() -> list[str]:
        """Fetch available GPT model IDs from the OpenAI API.

        Keeps only GPT-5 and above, in their base or -mini forms.
        Excludes older generations (gpt-4, gpt-3.5, ...), nano, turbo,
        instruct, dated snapshots, and any other variants.
        Returns an empty list if the API key is absent or the request fails.
        Results are cached for the lifetime of the process.
        """
        global _cached_models
        if _cached_models is not None:
            return _cached_models

        import json
        import re
        import urllib.error
        import urllib.request

        # Allow: gpt-5, gpt-5.4, gpt-5.5, gpt-5.4-mini, gpt-10.0-mini, ...
        # Major version must be >= 5; only base or -mini suffix permitted.
        _keep = re.compile(r"^gpt-([5-9]\d*|\d{2,})(\.\d+)?(-mini)?$")

        api_key = os.environ.get("OPENAI_API_KEY", "").strip()
        if not api_key:
            return []
        try:
            req = urllib.request.Request(
                "https://api.openai.com/v1/models",
                headers={"Authorization": f"Bearer {api_key}"},
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read())
            models = sorted(
                m["id"]
                for m in data.get("data", [])
                if isinstance(m.get("id"), str) and _keep.match(m["id"])
            )
            _cached_models = models
            return models
        except (urllib.error.URLError, OSError, ValueError):
            return []

    @staticmethod
    def available_models() -> list[str]:
        """Alias for fetch_models_from_api(); kept for call-site compatibility."""
        return AIClient.fetch_models_from_api()


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
