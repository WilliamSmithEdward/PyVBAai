# PyVBAai

**Chat naturally with your Excel workbooks — powered by OpenAI GPT.**

PyVBAai is a Windows desktop application that lets you load any `.xlsx` /
`.xlsm` file, send natural-language requests to an AI, and have it safely
modify cells, formulas, sheets, charts, tables, and VBA macros — with a full
preview before anything changes.

---

## Features

| Feature | Detail |
|---|---|
| **No Python required at runtime** | Single `.exe` (PyInstaller), no installer, no admin rights |
| **Pure-Python workbook I/O** | All reads and writes use `openpyxl` — no Excel COM dependency for editing |
| **Full workbook context** | Cells, formulas, formatting, named ranges, tables, charts, pivot tables, VBA modules |
| **Dynamic-array formulas** | FILTER, UNIQUE, SORT, XLOOKUP, SEQUENCE, LET, LAMBDA, etc. spill correctly in Excel 365 (see [docs/dynamic-array-formulas.md](docs/dynamic-array-formulas.md)) |
| **Natural chat UI** | Multi-turn conversation, Markdown rendering, typing indicator |
| **Safe change preview** | Every change shown with diff before applying |
| **Automatic backups** | Timestamped backup created before every apply |
| **Drag & drop** | Drop any `.xlsx` / `.xlsm` onto the window to load it |
| **Dark / Light mode** | Toggle in toolbar or Settings |
| **Selective context** | Include/exclude sheets and VBA modules via Settings |

---

## Prerequisites

### Runtime
- **Windows 10 or 11 (x64)**
- **OPENAI_API_KEY** set as a user-level Windows environment variable
- **Microsoft Excel** (any modern version) — required only for opening the
  resulting workbooks; PyVBAai itself does not need Excel installed to read
  or write the files. Excel 365 is required if you want dynamic-array
  formulas to spill on open.

### Build-time (only if building from source)
- Python 3.11+
- `pip install -r requirements.txt`

---

## Quick Start

### Option A — pre-built executable
1. Download `PyVBAai.exe` from the Releases page.
2. Set `OPENAI_API_KEY` (see below).
3. Double-click `PyVBAai.exe`.

### Option B — build from source
```bat
git clone https://github.com/WilliamSmithEdward/PyVBAai.git
cd PyVBAai
build.bat
:: Output: dist\PyVBAai.exe
```

### Option C — run from source
```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

---

## Setting OPENAI_API_KEY

1. Press `Win + R` → type `sysdm.cpl` → **Advanced** → **Environment Variables**
2. Under *User variables*, click **New**:
   - Name: `OPENAI_API_KEY`
   - Value: `sk-...` (your OpenAI key)
3. Click OK and **restart PyVBAai**.

---

## How It Works

```
User loads .xlsx / .xlsm
        │
        ▼
openpyxl reads sheets, formulas, formats, named ranges,
charts, tables, pivot tables, and VBA modules (from vbaProject.bin)
        │
        ▼
Compact, token-efficient context built
        │
User types a message in the chat window
        │
        ▼
Context + conversation history → GPT (JSON mode)
        │
        ▼
AI returns { message, changes[], diff_summary }
        │
        ▼
Preview dialog shows each change with optional VBA diff
        │
User clicks "Apply Changes"
        │
        ▼
Versioned backup created  →  changes applied via openpyxl
        │
Workbook re-read to refresh context for the next turn
```

---

## Supported Change Operations

| Category | Operations |
|---|---|
| **Cell content** | `set_cell`, `set_range`, `clear_range`, `clear_format` |
| **Formatting** | `set_format`, `set_col_width`, `set_row_height`, `merge_cells`, `unmerge_cells`, `set_tab_color` |
| **Sheet structure** | `add_sheet`, `delete_sheet`, `rename_sheet`, `move_sheet`, `copy_sheet`, `hide_sheet`, `unhide_sheet` |
| **Rows / columns** | `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`, `hide_rows`, `unhide_rows`, `hide_cols`, `unhide_cols` |
| **Tables** | `create_table`, `delete_table` |
| **Charts** | `create_chart`, `delete_chart` (bar, col, line, pie, doughnut, scatter, area, radar) |
| **View** | `freeze_panes`, `set_zoom`, `set_print_area`, `auto_filter` |
| **Protection** | `protect_sheet`, `unprotect_sheet` |
| **Named ranges** | `set_named_range`, `delete_named_range` |
| **VBA** | `set_vba`, `add_vba_module`, `delete_vba_module` |

Pivot tables and pivot charts cannot be created directly — they must be
generated via a VBA macro (`set_vba`).

---

## Settings

| Setting | Default | Description |
|---|---|---|
| Model | _none_ | Pick from the toolbar dropdown (live list filtered to GPT-5+ base / -mini variants). The selection is persisted; sending a message with no model selected is blocked. |
| Max context characters | `150,000` | Hard cap on the size of the snapshot sent to the AI per request |
| Row limit per data area | `20` (toggleable) | When enabled, each detected data area is truncated to this many rows in context |
| Include formulas | ✅ | Show cell formulas in context |
| Include VBA | ✅ | Include VBA source in context |
| Include named ranges | ✅ | Include named ranges in context |
| Cell formatting fields | all on | Per-field toggle: number_format, bold/italic/strike/underline/wrap, font name/size/colour, bg colour, h-/v-align, four borders |
| Max backups | `20` | Oldest backups in `backups/` pruned automatically |
| Dark mode | ✅ | Light/dark theme toggle (also available from the toolbar) |

---

## Project Structure

```
PyVBAai/
├── main.py                     # Entry point
├── app/
│   ├── main_window.py          # Main QMainWindow
│   ├── chat_widget.py          # Chat UI (bubbles, input bar)
│   ├── workbook_panel.py       # Sidebar tree + VBA viewer
│   ├── preview_dialog.py       # Change preview + diff
│   ├── settings_dialog.py      # Settings tabs
│   ├── vba_dialog.py           # VBA module viewer
│   ├── theme.py                # Dark / light QSS
│   ├── config.py               # QSettings wrapper
│   ├── logger.py               # Rotating-file logger
│   └── workers.py              # QThread workers
├── core/
│   ├── excel_reader.py         # openpyxl workbook extraction
│   ├── excel_writer.py         # openpyxl change application + dynamic-array post-processing
│   ├── context_builder.py      # Token-efficient context
│   ├── ai_client.py            # OpenAI wrapper + system prompt
│   └── backup_manager.py       # Versioned backups
├── models/
│   ├── workbook.py             # Workbook data classes
│   └── conversation.py         # Conversation / AIResponse
├── docs/
│   ├── cell-notation.md
│   └── dynamic-array-formulas.md
├── tests/                      # pytest suite (~310 tests)
├── requirements.txt
├── PyVBAai.spec                # PyInstaller spec
└── build.bat                   # One-click build script
```

---

## Development

```bat
:: Run the test suite
python -m pytest tests\ -q

:: Lint
python -m ruff check .
```

---

## License

See [LICENSE](LICENSE).
