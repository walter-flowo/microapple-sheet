# MicroApple Sheet

**The only Excel MCP server that won't corrupt your workbook.**

Most Excel MCP servers use openpyxl for everything. That silently drops every
cached formula value the moment it saves — your `=SUM(...)` cells come back as 0
in Excel. MicroApple Sheet routes each write through the correct engine
automatically, and is the only implementation with a proven Rule-0-safe path for
computed workbooks (320 → naive openpyxl = 20, 97% loss → safe two-step = 320,
100% restored).

## Why this one is different

| Capability | microapple-sheet | typical Excel MCP |
|---|---|---|
| Reads cached formula values back correctly | Yes | Often 0 / blank |
| Safe to edit computed workbooks | Yes (LO two-step) | No — drops cached `<v>` |
| Dynamic-array / LAMBDA / XLOOKUP workbooks | Refuses with clear error | May corrupt silently |
| Live edit while workbook is open in Excel | Yes (AppleScript) | No |
| Cross-platform file engine | Yes | Varies |

## Platform matrix

| Feature | macOS | Linux |
|---|---|---|
| File engine (create / read / write / recalc / convert / audit) | Yes | Yes |
| LibreOffice recalc (computed workbooks) | Yes | Yes (requires libreoffice-calc) |
| Live bridge (instant in-Excel edits via AppleScript) | **macOS + Microsoft Excel only** | Returns structured error |

## Install

**Via uvx (recommended — no install required):**

```bash
uvx --from git+https://github.com/walter-flowo/microapple-sheet microapple-sheet
```

**Via pip from GitHub:**

```bash
pip install git+https://github.com/walter-flowo/microapple-sheet
```

**Dev install:**

```bash
git clone https://github.com/walter-flowo/microapple-sheet
cd microapple-sheet
python3 -m venv .venv && source .venv/bin/activate
pip install -e .
```

## MCP client config

### Claude Code

```bash
claude mcp add microapple-sheet -- python -m microapple_sheet
```

Or with an explicit venv:

```bash
claude mcp add microapple-sheet -- /path/to/.venv/bin/python -m microapple_sheet
```

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "microapple-sheet": {
      "command": "python",
      "args": ["-m", "microapple_sheet"]
    }
  }
}
```

## Tool catalogue

### Lifecycle

| Tool | Description |
|---|---|
| `excel_ping` | Health check — returns server name and version |
| `excel_open` | Inspect workbook metadata and engine classification |
| `excel_info` | Lightweight file metadata without full parse |
| `excel_recalc` | Force LibreOffice recalculation (Rule-0-safe) |
| `excel_convert` | Convert to PDF, CSV, or XLSX via LibreOffice |
| `excel_check_automation` | Probe macOS TCC permissions for live bridge |

### Read

| Tool | Description |
|---|---|
| `excel_read` | Read a range in values / formulas / both modes |
| `excel_read_table` | Read as header-keyed records |
| `excel_list_sheets` | Return ordered sheet names |
| `excel_list_names` | Return all defined names / named ranges |

### Write (engine-routed)

| Tool | Description |
|---|---|
| `excel_create` | Create a new blank workbook |
| `excel_write` | Write a 2-D array of values (auto-detects live / file) |
| `excel_set_cell` | Write a single cell (auto-detects live / file) |
| `excel_write_formula` | Write a single formula (validated + auto-detects) |
| `excel_write_linked_formula` | Compose a formula from Cell.coordinate references |
| `excel_format` | Apply font / fill / border / number format / alignment |
| `excel_define_name` | Define or update a named range |

### Live bridge (macOS + Microsoft Excel only)

| Tool | Description |
|---|---|
| `excel_is_open` | Check whether a workbook is open in Excel |
| `excel_live_read` | Read live cell content (sees unsaved edits) |
| `excel_live_set` | Set values via one atomic Apple Event |
| `excel_live_formula` | Set a formula (triggers instant recalc in Excel) |

### Audit / config

| Tool | Description |
|---|---|
| `excel_audit_hardcoded` | Lint for numeric literals that should be formulas |
| `excel_config_get` | Read a config value |
| `excel_config_set` | Write a config value |

## Rule-0 engine routing

Before every write, `classify_workbook()` scans the xlsx ZIP (no openpyxl load)
and picks the engine:

```
workbook state                    → engine chosen
─────────────────────────────────────────────────
New / nonexistent                 → openpyxl (safe)
Exists, no cached formula <v>     → openpyxl (safe)
Exists, has cached formula <v>    → LibreOffice two-step (Rule 0)
Has _xlfn / XLOOKUP / LAMBDA      → refuse + clear error (Rule 0b)
Open in Excel (lock file / probe) → live bridge only
```

**The two-step (Rule 0 safe path):**

1. openpyxl writes value/formula edits — this intentionally drops cached `<v>`
   values, but preserves all formula strings in other cells.
2. Immediate LibreOffice `--convert-to` recalc restores every `<v>` from its
   formula.

Result: your edits land, and every dependent formula updates. Proven:
320 → naive openpyxl = 20 (97% loss) → safe path = 320 (100% restored).

**Why LibreOffice is refused for dynamic-array workbooks (Rule 0b):**

LibreOffice `--convert-to` recalc on a workbook containing XLOOKUP / LAMBDA /
dynamic-array formulas cascades to `#N/A` across dependent cells (measured:
6545 → 2197 cached cells, whole blocks gone). The only safe paths for these
workbooks are the live bridge (open in Excel, use `excel_live_set` /
`excel_live_formula`) or direct XML surgery.

## `excel_audit_hardcoded`

Lints a sheet for numeric literals that look calculated:

```python
# Whole-sheet mode (column-peer heuristic):
# A column where at least one cell is a formula → any numeric literal is flagged.
# Pure-input columns (no formula peers) are not flagged.
excel_audit_hardcoded(path="/path/to/model.xlsx", sheet="Calcs")

# Strict-range mode — every non-formula numeric in the range is flagged:
excel_audit_hardcoded(path="/path/to/model.xlsx", sheet="Output", range_="H3:J15")
```

Returns `{hardcoded: [...], formula_cells: N, input_literal_cells: N, verdict}`.

## macOS Automation permission

The live bridge requires Automation permission for your terminal emulator (or
Claude Code process) to send Apple Events to Microsoft Excel.

If `excel_check_automation()` returns `automation_ok: false` with error -1743:

1. Open **System Settings → Privacy & Security → Automation**.
2. Find your terminal (e.g. Terminal, iTerm2, or the process running Claude Code).
3. Enable the toggle next to **Microsoft Excel**.

## Roadmap

- **B001** — Blast-radius preview animation (pending-glyph markers, fill+border
  highlight, drain, content-aware)
- **B002** — XML surgery for dynamic-array / LAMBDA workbooks (cache-preserving
  file-mode edits without LO recalc)
- **B003** — Change-aware preview server (SSE + MCP resources for side-by-side
  on closed files)
- **B004** — Graph co-authoring backend (cloud / remote / multi-editor)
- **B005** — House-style module (palette, charts, pivots, CF breadth)

## License

MIT — Copyright (c) 2026 Walter Wong. See [LICENSE](LICENSE).
