# PyVBAai — Token-Efficient Cell Notation

## Goal
Minimise tokens in both directions (context sent to LLM and changes returned).
Key principle: row number appears once per row, column letters are abbreviated, format flags are single-char or short-prefix codes.

## Context format (read — what LLM receives)
```
R<row>: <COL>=<val>[flags]
```
- `"string"` quoted | bare number | `{formula without =}` | `None`
- flags (comma-sep inside `[...]`):
  - `B` bold · `I` italic · `S` strikethrough · `U` underline · `W` wrap_text
  - `#<spec>` number_format (e.g. `#0.00%`)
  - `fn:<name>` font_name · `fs:<pts>` font_size
  - `^<RRGGBB>` font_color · `~<RRGGBB>` bg_color
  - `ha:<l|c|r>` h_align · `va:<t|c|b>` v_align
  - `bt/bb/bl/br:<style>[:<RRGGBB>]` borders (thin|medium|thick|dashed|dotted|double|hair)

## Write-back (LLM → app)
LLM may use **compact flags string** or **explicit JSON keys** in `set_cell` / `set_range`:
```json
{"type":"set_cell","sheet":"S","cell":"A1","value":"Rev","flags":"B,#$#,##0,bt:thin"}
{"type":"set_cell","sheet":"S","cell":"A1","value":"Rev","bold":true,"number_format":"$#,##0","border_top":"thin"}
```
Both are equivalent. `_expand_flags()` in `excel_writer.py` parses the compact form.
`set_format` applies format to a range without changing values (all format keys optional).

## Supported format fields
`bold` `italic` `strikethrough` `underline` `wrap_text`
`number_format` `font_name` `font_size` `font_color` `bg_color`
`h_align` (left|center|right|fill|justify)  `v_align` (top|center|bottom)
`border_top` `border_bottom` `border_left` `border_right`  — value: `"<style>"` or `"<style>:<RRGGBB>"`

## Implementation files
- `models/workbook.py` — `CellFormat` dataclass + `ALL_FMT_FIELDS`
- `core/excel_reader.py` — `_cell_format()` reads all fields from openpyxl
- `core/context_builder.py` — `_fmt_hints()` + `_CELL_LEGEND` emits flags
- `core/excel_writer.py` — `_expand_flags()` + `_set_format()` applies all fields
- `core/ai_client.py` — `_SYSTEM_PROMPT` documents notation to LLM
- `app/settings_dialog.py` — per-field toggles in Context tab
