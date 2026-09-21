# Excel MCP Server (`excel-mcp-rw`)

> **High-Performance Read & Write Model Context Protocol (MCP) Server for Microsoft Excel on Windows.**  
> Powered by **Pandas**, **OpenPyXL**, and **Native Windows Microsoft Excel Automation (COM / PyWin32)**.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-windows-lightgrey.svg)](https://microsoft.com/windows)
[![MCP](https://img.shields.io/badge/MCP-Standard-green.svg)](https://modelcontextprotocol.io/)

---

## ⚡ Why This Server?

Most Excel MCP servers are lightweight JavaScript wrappers that fail on real-world spreadsheets:
* ❌ They crash or exceed LLM context windows on large multi-megabyte files.
* ❌ They **cannot calculate dynamic Excel formulas** (`=XLOOKUP`, `=FILTER()`, dynamic arrays, complex nesting).
* ❌ They cannot refresh PivotTables or external database connections.
* ❌ They cannot run VBA macros (`.xlsm`) or export pixel-perfect PDFs.

**`excel-mcp-rw`** solves this with a **Dual-Engine Architecture**:

1. **Engine 1: Fast Python Data Science Stack (`pandas` + `openpyxl`)**
   * Instant metadata streaming without loading multi-GB spreadsheets into RAM.
   * Vectorized filtering (`df.query()`) and column statistics at compiled C-speed.
   * Bounded pagination so the model never blows its context window.
2. **Engine 2: Native Microsoft Excel Engine (COM Automation via `pywin32`)**
   * Uses real Windows Microsoft Excel to execute `Application.CalculateFull()`.
   * Refreshes PivotTables and data connections (`wb.RefreshAll()`).
   * Native print-engine PDF export.
   * Runs VBA macros and inspects active workbooks open on the desktop.

---

## 🛠️ Available Tools

### 📊 Data Analysis & Manipulation (`pandas` / `openpyxl`)

| Tool | Description |
| :--- | :--- |
| `get_workbook_info` | Read sheet names, dimensions, column header preview, and file size in MB in read-only streaming mode. |
| `preview_sheet` | Preview first $N$ rows (default 10) with column data types and shape. |
| `query_rows` | Filter rows with vectorized Pandas expressions (e.g. `Age > 30 and Status == 'Active'`), select specific columns, with `limit` and `offset` pagination. |
| `read_range` | Windowed slice read (e.g. rows 100 to 200) without loading the entire spreadsheet. |
| `summarize_column` | Comprehensive column statistics: min, max, mean, median, std, sum (numeric) or top 10 value frequencies (categorical). |
| `search_text` | Fast pattern and keyword search across all rows and columns with row number coordinates. |
| `create_workbook` | Create formatted `.xlsx` files with auto-fitted column widths from JSON records. |
| `add_sheet` | Add a brand new sheet tab to an existing workbook without touching other sheets (with optional data and headers). |
| `rename_sheet` | Rename an existing sheet tab. |
| `delete_sheet` | Delete a specific worksheet tab from an existing workbook. |
| `append_rows` | In-place row append without rewriting the workbook. |
| `update_cells` | Surgical updates to specific cells or formulas (e.g. `[{"cell": "C10", "value": "=SUM(C1:C9)"}]`). |
| `export_to_csv` | Export large sheets to clean CSV format. |

### 🪟 Native Windows Microsoft Excel Tools (COM / PyWin32)

| Tool | Description |
| :--- | :--- |
| `recalculate_and_save` | Opens workbook in native Microsoft Excel, executes `CalculateFull()`, saves, and closes. |
| `export_to_pdf` | Generates a pixel-perfect PDF using Microsoft Excel's internal print engine. |
| `refresh_data_and_pivots` | Refreshes all external data connections and PivotTables (`RefreshAll()`). |
| `run_vba_macro` | Runs any VBA macro inside `.xlsm` or `.xlsb` workbooks. |
| `get_active_excel_window` | Detects if Excel is open on your desktop and returns the active workbook, active sheet, and selected range. |

---

## 🚀 Quickstart & Installation

### 1. Clone & Install Dependencies

```bash
git clone https://github.com/indala/excel-mcp-rw.git
cd excel-mcp-rw
pip install -r requirements.txt
```

---

## ⚙️ Client Configurations

### 1. Claude Desktop

Add this to your `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "C:/path/to/excel-mcp-rw/src/excel_mcp/server.py"
      ]
    }
  }
}
```

### 2. Claude Code CLI

Add globally to your user scope:

```bash
claude mcp add -s user excel -- python C:/path/to/excel-mcp-rw/src/excel_mcp/server.py
```

### 3. Google Antigravity

Add to `~/.gemini/config/mcp_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "C:/path/to/excel-mcp-rw/src/excel_mcp/server.py"
      ]
    }
  }
}
```

### 4. Cursor / Windsurf

Add a new STDIO server in your MCP settings:
* **Name**: `excel`
* **Type**: `command` (or `stdio`)
* **Command**: `python C:/path/to/excel-mcp-rw/src/excel_mcp/server.py`

---

## 📝 Example Prompts

Once connected to your assistant, you can ask:

* *"Give me an overview of all sheets in `C:/reports/financials_q3.xlsx` without loading the whole file."*
* *"Filter the `Transactions` sheet for `Amount > 10000 and Status == 'Pending'`, and give me the top 20 rows."*
* *"What is the statistical distribution of the 'Salary' column in our employee database?"*
* *"Search for customer ID 'CUST-8492' across all sheets and tell me which row it appears on."*
* *"Recalculate all formulas in this sheet using native Excel and export it to PDF."*
* *"What Excel spreadsheet do I currently have open on my desktop?"*

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
