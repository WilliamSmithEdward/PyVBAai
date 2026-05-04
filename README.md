# PyVBAai

**Chat naturally with your Excel workbooks — powered by OpenAI GPT.**

PyVBAai is a standalone Windows desktop application that lets you load any `.xlsx` / `.xlsm` file, send natural-language requests to an AI, and have it safely modify cells, formulas, sheets, and VBA macros — with full preview before anything changes.

---

## Features

| Feature | Detail |
|---|---|
| **No Python required** | Single `.exe` (PyInstaller), no installer, no admin rights |
| **Full workbook context** | Reads cells, formulas, VBA modules, named ranges via Excel COM |
| **Natural chat UI** | Multi-turn conversation, Markdown rendering, typing indicator |
| **Safe change preview** | All proposed changes shown with diff before applying |
| **Automatic backups** | Timestamped `.xlsm` backup created before every apply |
| **Drag & drop** | Drop any `.xlsx` / `.xlsm` onto the window to load it |
| **Dark / Light mode** | Toggle in toolbar or Settings |
| **Selective context** | Include/exclude sheets and VBA modules via Settings |

---

## Prerequisites

### Runtime
- **Windows 10 or 11 (x64)**
- **Microsoft Excel 365 x64** installed
- **OPENAI_API_KEY** set as a user-level Windows environment variable

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
git clone https://github.com/YourName/PyVBAai.git
cd PyVBAai
build.bat
# Output: dist\PyVBAai.exe
```

---

## Setting OPENAI_API_KEY

1. Press `Win + R` → type `sysdm.cpl` → **Advanced** → **Environment Variables**
2. Under *User variables*, click **New**:
   - Name: `OPENAI_API_KEY`
   - Value: `sk-...` (your OpenAI key)
3. Click OK and **restart PyVBAai**.

---

## Enabling VBA Access (required for VBA read/write)

By default Excel blocks programmatic access to the VBA project.

1. Open Excel → **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Click **Macro Settings**
3. Check **"Trust access to the VBA project object model"**
4. Click OK

---

## How It Works

```
User loads .xlsx/.xlsm
        │
        ▼
Excel COM reads all sheets + VBA modules + named ranges
        │
        ▼
Compact token-efficient context built
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
Versioned backup created  →  changes applied via COM
        │
Workbook re-read to update context for next turn
```

---

## Context Format sent to GPT

```
=== WORKBOOK: Budget2024.xlsm ===
SHEETS (3): Sheet1(150r×10c), Data(500r×6c), Summary(15r×4c)
VBA MODULES (2): Module1, ThisWorkbook
NAMED RANGES (1):
  BudgetTotal = Sheet1!$A$1:$D$10

--- CELLS: Sheet1 (100/150 rows shown, 50 rows truncated) ---
  A1="Category", B1="Jan", C1="Feb"
  A2="Revenue", B2=1000, C2=1200
  A10="Total", B10={=SUM(B2:B9)}
...

--- VBA: Module1 (Module) ---
Sub CalculateTotals()
    ...
End Sub
```

---

## Supported Change Operations

| Operation | Description |
|---|---|
| `set_cell` | Set a cell value or formula |
| `set_range` | Write a 2-D array of values |
| `clear_range` | Clear cell contents |
| `add_sheet` / `delete_sheet` | Add or remove a worksheet |
| `rename_sheet` / `move_sheet` / `copy_sheet` | Structural sheet changes |
| `set_vba` | Replace a VBA module's code |
| `add_vba_module` / `delete_vba_module` | Add or remove VBA modules |
| `add_named_range` / `delete_named_range` | Manage named ranges |

---

## Settings

| Setting | Default | Description |
|---|---|---|
| Model | `gpt-4o` | OpenAI model to use |
| Max rows per sheet | `100` | Context truncation limit |
| Include formulas | ✅ | Show cell formulas in context |
| Include VBA | ✅ | Include VBA source in context |
| Include named ranges | ✅ | Include named ranges |
| Max backups | `20` | Oldest backups pruned automatically |

---

## Project Structure

```
PyVBAai/
├── main.py                  # Entry point
├── app/
│   ├── main_window.py       # Main QMainWindow
│   ├── chat_widget.py       # Chat UI (bubbles, input bar)
│   ├── workbook_panel.py    # Sidebar tree + VBA viewer
│   ├── preview_dialog.py    # Change preview + diff
│   ├── settings_dialog.py   # Settings tabs
│   ├── theme.py             # Dark / light QSS
│   └── workers.py           # QThread workers
├── core/
│   ├── excel_reader.py      # COM workbook extraction
│   ├── excel_writer.py      # COM change application
│   ├── context_builder.py   # Token-efficient context
│   ├── ai_client.py         # OpenAI wrapper
│   └── backup_manager.py    # Versioned backups
├── models/
│   ├── workbook.py          # Workbook data classes
│   └── conversation.py      # Conversation / AIResponse
├── requirements.txt
├── PyVBAai.spec             # PyInstaller spec
└── build.bat                # One-click build script
```