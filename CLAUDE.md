# Excel MCP — Project Doctrine

## Purpose

Engine-routed, Rule-0-safe, live-bridge Excel MCP server for Claude Code.

Every off-the-shelf Excel MCP server uses a single openpyxl load-modify-save path.
That path **silently drops every cached formula value on save** (Rule 0: measured 2247→22
cells, 99% loss, on the Bewl HVAC calc). This server makes engine-routing,
live-formula enforcement, recalc-preservation, and per-save change-awareness first-class.

---

## Engine routing table (spine of the server)

`engine.classify_workbook(path)` is called before EVERY write:

| Target state | Engine | Why |
|---|---|---|
| New / nonexistent | **openpyxl** | Fresh build; Excel recalcs on open |
| Exists, no cached `<v>` under `<f>` | **openpyxl** | No cached values to lose — safe |
| Exists, has cached formula values | **LibreOffice headless macro** | Recalc-preserving (Rule 0) |
| Exists, has `_xlfn`/XLOOKUP/LAMBDA/IFS/array | **XML surgery** + Excel AppleScript fallback | LO cascades these to `#N/A` (Rule 0b) |
| File open in Excel | **AppleScript bridge** ONLY | Excel owns the file; never dual-write |

---

## Tool catalogue (MVP)

| Tool | Purpose |
|---|---|
| `excel_ping` | Health / version check — always implemented |
| `excel_check_automation` | TCC AppleScript permission preflight — always implemented |
| `excel_create` | Build a new blank workbook |
| `excel_open` | Classify + return metadata (engine, hash, dims…) |
| `excel_info` | Lightweight metadata without full parse |
| `excel_read` | Read values / formulas / both |
| `excel_list_sheets` | List sheet names |
| `excel_list_names` | List defined names |
| `excel_write` | Write literal values (engine-routed) |
| `excel_write_formula` | Write a live formula (validated) |
| `excel_write_linked_formula` | Compose formula from Cell.coordinate refs (Rule #1/#4) |
| `excel_format` | Apply formatting (engine-routed) |
| `excel_define_name` | Add / update a named range |
| `excel_live_read` | Read live (unsaved) values via AppleScript bridge |
| `excel_live_set` | Write values via AppleScript (atomic AE + read-back) |
| `excel_live_formula` | Write formula via AppleScript (triggers recalc) |
| `excel_is_open` | Probe whether workbook is open in Excel |
| `excel_audit_hardcoded` | Lint workbook for literal values in calc columns |
| `excel_config_get` | Read config.toml key |
| `excel_config_set` | Write config.toml key (persists; preserves literal marker) |
| `excel_recalc` | Force LibreOffice full recalculation |
| `excel_convert` | Export to PDF / CSV / XLSX |

**Deferred (later phases):** `excel_watch`, `excel_changes`, `excel_clear_stray_markers`,
blast-radius animation (⏳ markers), XML surgery (`ops_xml.py`), `watcher.py`, `preview.py`,
`graph.py`, `style.py`.

---

## Module map

| Module | Phase | Role |
|---|---|---|
| `server.py` | 0 ✅ | FastMCP app + all tool registrations |
| `config.py` | 0 ✅ | config.toml load/validate/write — IMPLEMENTED |
| `engine.py` | 1a | classify_workbook + disk_hash + is_open_in_excel |
| `libreoffice.py` | 1a | Headless macro invocation (recalc / edit / convert) |
| `backups.py` | 1a | Timestamped snapshots + prune |
| `ops_openpyxl.py` | 1b | Create / read / write / format (openpyxl safe path) |
| `liveformula.py` | 3 | validate_formula + audit_hardcoded + compose_linked |
| `bridge.py` | 2 | AppleScript live read/write via ScriptingBridge |
| `libreoffice/Module1.xba` | 1a | StarBasic macro (EditCells, RecalcOnly) |

---

## Build & test

```bash
# Create venv (Python 3.14)
python3 -m venv .venv
.venv/bin/pip install -e .

# Run smoke tests
.venv/bin/python -m pytest tests/ -q
# or without pytest
.venv/bin/python tests/test_smoke.py

# Import sanity
.venv/bin/python -c "import microapple_sheet.server; print('OK')"

# Register with Claude Code (stdio transport)
claude mcp add microapple-sheet -- .venv/bin/python -m microapple_sheet
```

---

## Key invariants (never break)

1. **Rule 0**: Never openpyxl-load-modify-save a workbook with cached `<v>` values.
2. **Rule 0b**: Never LibreOffice-convert a workbook with `_xlfn`/XLOOKUP/LAMBDA/IFS/array formulas.
3. **Rule #1/#4**: Every calculated cell holds a live formula referencing `Cell.coordinate` captures — never a hardcoded number.
4. **Backup before write**: `backups.snapshot(path)` before every file-engine write.
5. **Staleness guard**: refuse stale write if disk hash changed since last read.
6. **Never dual-write**: if Excel has the file open, refuse file-engine writes.

---

## Dependencies

- Python 3.14+, `mcp[cli]`, `openpyxl`, `xlsxwriter`, `pandas`, `watchdog`, `lxml`, `Pillow`
- System: LibreOffice 26.2.3.2 at `/Applications/LibreOffice.app`
- System: Microsoft Excel (optional — needed for live-bridge mode only)
