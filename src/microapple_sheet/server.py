"""FastMCP server for microapple-sheet.

Tools fully implemented:
  Phase 0:   excel_ping, excel_check_automation
  Phase 1a:  excel_open, excel_info, excel_recalc, excel_convert
  Phase 1b:  excel_create, excel_read, excel_read_table, excel_list_sheets,
             excel_list_names, excel_write, excel_set_cell, excel_write_formula,
             excel_write_linked_formula, excel_format, excel_define_name
  Phase 2:   excel_live_read, excel_live_set, excel_live_formula, excel_is_open
  Phase 3:   excel_audit_hardcoded, excel_config_get, excel_config_set
             + auto-detect routing in excel_write, excel_set_cell, excel_write_formula
"""
from __future__ import annotations

import os
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from microapple_sheet import __version__

mcp = FastMCP(
    "microapple-sheet",
    instructions="""
# MicroApple Sheet MCP Server

Engine-routed, Rule-0-safe Excel automation for Claude Code.

## Key behaviours
- **Engine routing**: classify_workbook() picks openpyxl (new/safe) vs LibreOffice (computed) vs XML surgery (dynamic-array) on every write.
- **Rule 0 safe**: never silently drops cached formula values.
- **Live bridge**: when Excel has the file open, all writes go via AppleScript (not disk).
- **Live-formula enforcement**: write_formula validates; audit_hardcoded lints.

## Preflight
Run excel_check_automation() to verify TCC permissions before live-mode tools.
""",
)


# ---------------------------------------------------------------------------
# Implemented tools (Phase 0 — proof of life)
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_ping() -> dict[str, Any]:
    """Return server health and version.

    Returns:
        {ok: bool, server: str, version: str}
    """
    return {"ok": True, "server": "microapple-sheet", "version": __version__}


@mcp.tool()
def excel_check_automation() -> dict[str, Any]:
    """Probe macOS Automation (TCC) permissions needed by the live AppleScript bridge.

    Sends a benign Apple Event to Microsoft Excel to verify automation access.
    Detects TCC denial (error -1743) and Excel-not-running states separately.
    Does NOT launch Excel if it is not already running.

    Returns:
        {automation_ok: bool, excel_running: bool, detail: str}
    """
    from microapple_sheet import bridge
    return bridge.check_automation()


# ---------------------------------------------------------------------------
# Stub tools — Phase 1: lifecycle & file engine
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_create(path: str, sheets: list[str] | None = None) -> dict[str, Any]:
    """Create a new blank workbook at *path*.

    Args:
        path:   Absolute path for the new .xlsx file (parent dir must exist).
        sheets: Optional list of sheet names (default: ['Sheet1']).

    Returns:
        {path: str, sheets: list[str], engine: str}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.create(path, sheets)


@mcp.tool()
def excel_open(path: str) -> dict[str, Any]:
    """Inspect a workbook and return metadata needed for routing decisions.

    Scans the xlsx ZIP to classify the workbook (no openpyxl load, no cached-value risk).

    Args:
        path: Absolute path to the .xlsx / .xlsm file.

    Returns:
        {path, sheets, engine, has_cached_values, has_dynamic_arrays,
         disk_hash, is_open_in_excel}
    """
    from microapple_sheet.engine import classify_workbook
    cls = classify_workbook(path)
    return cls.as_dict()


@mcp.tool()
def excel_info(path: str) -> dict[str, Any]:
    """Lightweight file metadata for a workbook (no full parse, no openpyxl load).

    Args:
        path: Absolute path to the workbook.

    Returns:
        {path, exists, size_bytes, modified_iso, sheets, engine, disk_hash}
    """
    from microapple_sheet.engine import classify_workbook, disk_hash as _hash
    p = Path(path)
    if not p.exists():
        return {
            "path": str(p),
            "exists": False,
            "size_bytes": None,
            "modified_iso": None,
            "sheets": [],
            "engine": "openpyxl",
            "disk_hash": None,
        }
    stat = p.stat()
    modified_iso = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat()
    cls = classify_workbook(p)
    return {
        "path": str(p),
        "exists": True,
        "size_bytes": stat.st_size,
        "modified_iso": modified_iso,
        "sheets": cls.sheets,
        "engine": cls.engine.value,
        "disk_hash": cls.disk_hash,
    }


@mcp.tool()
def excel_recalc(path: str) -> dict[str, Any]:
    """Force a full LibreOffice recalculation of a workbook (Rule-0-safe).

    Preserves all cached formula values — proven safe: 320 populated <v> before = 320 after.
    Creates a timestamped backup before touching the file.

    Args:
        path: Absolute path to the workbook.

    Returns:
        {path, engine: 'libreoffice', recalc_ok: bool, backup_path: str, detail: str}
    """
    from microapple_sheet import libreoffice
    result = libreoffice.recalc(path)
    result["engine"] = "libreoffice"
    return result


@mcp.tool()
def excel_convert(
    path: str,
    to: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Convert a workbook to PDF, CSV, or another format via LibreOffice.

    Args:
        path:  Absolute path to the source workbook.
        to:    Target format — 'pdf', 'csv', or 'xlsx'.
        sheet: Sheet name for CSV export (optional; ignored for pdf/xlsx).

    Returns:
        {output_path: str, engine: str, ok: bool, detail: str}
    """
    from microapple_sheet import libreoffice
    result = libreoffice.convert(path, to=to, sheet=sheet)
    result["engine"] = "libreoffice"
    return result


# ---------------------------------------------------------------------------
# Stub tools — Phase 1: read
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_read(
    path: str,
    sheet: str,
    range: str | None = None,
    mode: str = "values",
) -> dict[str, Any]:
    """Read cells from a worksheet.

    Args:
        path:  Absolute path to the workbook.
        sheet: Sheet name.
        range: A1-notation range (e.g. 'A1:D10'). If omitted, reads entire used range.
        mode:  'values' (cached <v>; None for un-recalced formulas) |
               'formulas' (formula strings like '=SUM(A1)') |
               'both' (returns both 'values' and 'formulas' keys).

    Returns:
        mode='values'|'formulas': {sheet, range, mode, data: list[list[Any]]}
        mode='both': {sheet, range, mode, values: 2D, formulas: 2D}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.read_range(path, sheet, range_=range, mode=mode)


@mcp.tool()
def excel_read_table(
    path: str,
    sheet: str,
    range: str | None = None,
) -> dict[str, Any]:
    """Read a worksheet table as header-keyed records.

    The first row is treated as column headers. Returns a list of dicts
    (one per data row) suitable for direct use in reports or further analysis.

    Args:
        path:  Absolute path to the workbook.
        sheet: Sheet name.
        range: A1-notation range. None → entire used range.

    Returns:
        {sheet, range, headers: list[str], records: list[dict[str, Any]], rows: int}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.read_table(path, sheet, range_=range)


@mcp.tool()
def excel_list_sheets(path: str) -> dict[str, Any]:
    """Return the ordered list of sheet names in a workbook.

    Args:
        path: Absolute path to the workbook.

    Returns:
        {path, sheets: list[str]}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.list_sheets(path)


@mcp.tool()
def excel_list_names(path: str) -> dict[str, Any]:
    """Return all defined names (named ranges) in a workbook.

    Args:
        path: Absolute path to the workbook.

    Returns:
        {path, names: list[{name, ref, scope}]}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.list_names(path)


# ---------------------------------------------------------------------------
# Stub tools — Phase 1: write (engine-routed)
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_write(
    path: str,
    sheet: str,
    range: str,
    values: list[list[Any]],
) -> dict[str, Any]:
    """Write a 2-D array of values to a range (engine-routed, backup + clobber-safety).

    Auto-detects whether the workbook is open in Excel:
    - Open → routes through the live AppleScript bridge (excel_live_set).
    - Closed → routes through the file engine (openpyxl / LibreOffice).

    Values that are strings starting with '=' are treated as formulas.
    For LIBREOFFICE-engine workbooks, routes through libreoffice.apply_edits
    to preserve all cached formula values (Rule-0 safe).

    Args:
        path:   Absolute path to the workbook.
        sheet:  Sheet name.
        range:  Top-left cell or full A1 range (e.g. 'B4' or 'B4:D10').
        values: Row-major 2-D list of values. Use None for empty cells.

    Returns:
        {path, sheet, range, engine, cells_written, backup_path, recalc_ok, routed}
    """
    from microapple_sheet import bridge, ops_openpyxl
    probe = bridge.is_open(path)
    if probe.get("is_open"):
        result = bridge.live_set(path, sheet, range, values)
        result["routed"] = "live"
        return result
    result = ops_openpyxl.write_range(path, sheet, range, values)
    result["routed"] = "file"
    return result


@mcp.tool()
def excel_set_cell(
    path: str,
    sheet: str,
    cell: str,
    value: Any,
) -> dict[str, Any]:
    """Write a single cell value (engine-routed convenience tool).

    Auto-detects whether the workbook is open in Excel:
    - Open → routes through the live AppleScript bridge.
    - Closed → routes through the file engine.

    Args:
        path:  Absolute path to the workbook.
        sheet: Sheet name.
        cell:  A1 cell reference (e.g. 'B4').
        value: Literal value or formula string (if starts with '=').

    Returns:
        {path, sheet, cell, engine, backup_path, recalc_ok, routed}
    """
    from microapple_sheet import bridge, ops_openpyxl
    probe = bridge.is_open(path)
    if probe.get("is_open"):
        result = bridge.live_set(path, sheet, cell, [[value]])
        result["routed"] = "live"
        result["cell"] = cell
        return result
    result = ops_openpyxl.set_cell(path, sheet, cell, value)
    result["routed"] = "file"
    return result


@mcp.tool()
def excel_write_formula(
    path: str,
    sheet: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Write a single Excel formula to a cell (engine-routed, live-formula validated).

    Auto-detects whether the workbook is open in Excel:
    - Open → routes through the live AppleScript bridge.
    - Closed → routes through the file engine.

    Validates the formula via liveformula.validate_formula. If the formula
    contains no cell references (e.g. '=42'), a formula_warning is included in
    the result but the write proceeds.

    Args:
        path:    Absolute path to the workbook.
        sheet:   Sheet name.
        cell:    Target cell in A1 notation (e.g. 'B4').
        formula: Formula string — with or without leading '='.

    Returns:
        {path, sheet, cell, formula, engine, backup_path, recalc_ok, routed,
         formula_warning?}
    """
    from microapple_sheet import bridge, liveformula, ops_openpyxl

    # Normalise formula
    norm = formula.strip()
    if not norm.startswith("="):
        norm = "=" + norm

    # Validate — include warning if formula is a bare constant
    _, formula_warning = liveformula.validate_formula(norm)

    probe = bridge.is_open(path)
    if probe.get("is_open"):
        result = bridge.live_formula(path, sheet, cell, norm)
        result["routed"] = "live"
    else:
        result = ops_openpyxl.write_formula(path, sheet, cell, norm)
        result["routed"] = "file"

    if formula_warning:
        result["formula_warning"] = formula_warning

    return result


@mcp.tool()
def excel_write_linked_formula(
    path: str,
    sheet: str,
    cell: str,
    template: str,
    refs: dict[str, str],
) -> dict[str, Any]:
    """Compose and write a live formula using captured Cell.coordinate references.

    Implements Rule #1/#4: the caller supplies coordinates captured from
    Cell.coordinate (e.g. {'peak': 'B4', 'fraction': '$B$2'}) so the formula
    is built from live cell references, never from hardcoded strings.

    Args:
        path:     Absolute path to the workbook.
        sheet:    Sheet name.
        cell:     Target cell in A1 notation.
        template: Formula template with {name} placeholders.
                  Example: '={peak}*{fraction}'
        refs:     Map of placeholder → A1 coordinate.
                  Example: {'peak': 'B4', 'fraction': '$B$2'}

    Returns:
        {path, sheet, cell, resolved_formula, engine, backup_path, recalc_ok}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.write_linked_formula(path, sheet, cell, template, refs)


@mcp.tool()
def excel_format(
    path: str,
    sheet: str,
    range: str,
    font: dict[str, Any] | None = None,
    fill: dict[str, Any] | None = None,
    border: dict[str, Any] | None = None,
    number_format: str | None = None,
    alignment: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Apply formatting to a cell range (engine-routed).

    For LIBREOFFICE-engine workbooks, runs an LO recalc after saving to restore
    cached formula values. Note: x14 CF (databars/icon-sets) may be corrupted by
    the openpyxl round-trip; plain dxf cellIs/expression CF survives.

    Args:
        path:          Absolute path to the workbook.
        sheet:         Sheet name.
        range:         A1-notation range.
        font:          {name?, size?, bold?, italic?, color?} (color = hex RRGGBB).
        fill:          {fgColor?} (hex RRGGBB solid fill).
        border:        {top?, bottom?, left?, right?} each {style?, color?}.
        number_format: Excel number-format string (e.g. '#,##0.00').
        alignment:     {horizontal?, vertical?, wrap_text?}.

    Returns:
        {path, sheet, range, engine, backup_path, recalc_ok, recalc_detail}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.format_range(
        path, sheet, range,
        font=font, fill=fill, border=border,
        number_format=number_format, alignment=alignment,
    )


@mcp.tool()
def excel_define_name(
    path: str,
    name: str,
    ref: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Define or update a named range / constant in the workbook.

    For LIBREOFFICE-engine workbooks, runs an LO recalc after saving.

    Args:
        path:  Absolute path to the workbook.
        name:  Defined name (must be a valid Excel identifier, no spaces).
        ref:   Reference in A1 notation (e.g. 'Sheet1!$B$4:$B$10' or '42').
        sheet: If provided, scope the name to this sheet only.

    Returns:
        {path, name, ref, scope, engine, backup_path, recalc_ok}
    """
    from microapple_sheet import ops_openpyxl
    return ops_openpyxl.define_name(path, name, ref, sheet)


# ---------------------------------------------------------------------------
# Stub tools — Phase 2: live bridge (AppleScript)
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_live_read(
    path: str,
    sheet: str,
    range: str,
    mode: str = "values",
) -> dict[str, Any]:
    """Read live cell content from an open workbook via AppleScript.

    Sees unsaved edits — reads directly from Excel's in-memory model.
    Requires Excel to have *path* open (use excel_is_open() to confirm first).

    Args:
        path:  Absolute path to the open workbook.
        sheet: Sheet name.
        range: A1-notation range (e.g. 'B4:D10').
        mode:  ``"values"`` (default) or ``"formulas"``.

    Returns:
        {path, sheet, range, data: list[list[Any]], source: 'live', mode}
    """
    from microapple_sheet import bridge
    try:
        return bridge.live_read(path, sheet, range, mode=mode)
    except RuntimeError as exc:
        return {"ok": False, "error": str(exc)}


@mcp.tool()
def excel_live_set(
    path: str,
    sheet: str,
    range: str,
    values: list[list[Any]],
) -> dict[str, Any]:
    """Set cell values in an open workbook via a single atomic Apple Event.

    Bypasses disk entirely — Excel sees the change instantly without a save.
    Never calls Select or Activate. Read-back verifies the write.
    Shows a macOS notification reminding the user to ⌘S.

    Args:
        path:   Absolute path to the open workbook.
        sheet:  Sheet name.
        range:  Top-left cell or full range in A1 notation.
        values: Row-major 2-D list of values.

    Returns:
        {path, sheet, range, cells_written: int, verified: bool, detail}
    """
    from microapple_sheet import bridge
    try:
        return bridge.live_set(path, sheet, range, values)
    except RuntimeError as exc:
        return {"ok": False, "error": str(exc)}


@mcp.tool()
def excel_live_formula(
    path: str,
    sheet: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Set a formula in an open workbook via AppleScript (triggers instant recalc).

    Args:
        path:    Absolute path to the open workbook.
        sheet:   Sheet name.
        cell:    Target cell in A1 notation.
        formula: Formula string — must start with ``=``.

    Returns:
        {path, sheet, cell, formula, verified_value: Any, detail}
    """
    from microapple_sheet import bridge
    try:
        return bridge.live_formula(path, sheet, cell, formula)
    except RuntimeError as exc:
        return {"ok": False, "error": str(exc)}


@mcp.tool()
def excel_is_open(path: str) -> dict[str, Any]:
    """Probe whether a workbook is currently open in Microsoft Excel.

    Matches by full resolved path, with basename fallback for iCloud path divergence.

    Args:
        path: Absolute path to the workbook.

    Returns:
        {path, is_open: bool, workbook_name: str|None, detail: str}
    """
    from microapple_sheet import bridge
    return bridge.is_open(path)


# ---------------------------------------------------------------------------
# Stub tools — Phase 3: live-formula enforcement + config
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_audit_hardcoded(
    path: str,
    sheet: str | None = None,
    range_: str | None = None,
) -> dict[str, Any]:
    """Lint a workbook for cells that look calculated but contain literal values.

    Uses the column-peer heuristic: in a column where at least one cell is a formula,
    any numeric literal in that column is flagged as hardcoded. Pure-input columns
    (no formula cells at all) are NOT flagged — they are recognised as legitimate
    input areas.

    If range_ is given (e.g. 'H3:J5'), strict mode applies: every non-formula
    numeric cell in that range is flagged regardless of column composition.

    Args:
        path:   Absolute path to the workbook.
        sheet:  Sheet name to audit. Required.
        range_: Optional A1-notation range for strict-mode audit.

    Returns:
        {sheet, scope, hardcoded, formula_cells, input_literal_cells, verdict}
    """
    from microapple_sheet import liveformula
    if sheet is None:
        raise ValueError("sheet is required for excel_audit_hardcoded")
    return liveformula.audit_hardcoded(path, sheet, range_=range_)


@mcp.tool()
def excel_config_get(key: str) -> dict[str, Any]:
    """Read a configuration value from config.toml.

    Args:
        key: Dot-separated key path (e.g. 'preview.marker').

    Returns:
        {key, value}
    """
    from microapple_sheet import config
    value = config.get(key)
    return {"key": key, "value": value}


@mcp.tool()
def excel_config_set(key: str, value: Any) -> dict[str, Any]:
    """Write a configuration value to config.toml (persists immediately).

    Preserves the literal TOML single-quoted string for preview.marker.
    Reads the old value before writing so the response includes the diff.

    Args:
        key:   Dot-separated key path (e.g. 'preview.enabled').
        value: New value (must be JSON-serialisable).

    Returns:
        {key, old_value, new_value, written: bool}
    """
    from microapple_sheet import config
    try:
        old_value: Any = config.get(key)
    except (KeyError, FileNotFoundError):
        old_value = None
    config.set(key, value)
    return {"key": key, "old_value": old_value, "new_value": value, "written": True}


# ---------------------------------------------------------------------------
# Sheet management (engine-routed: live when open, file engine when closed)
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_add_sheet(path: str, name: str) -> dict[str, Any]:
    """Add a worksheet to a workbook (engine-routed).

    Open in Excel → live AppleScript bridge (added at end, no save needed).
    Closed → file engine (openpyxl, Rule-0 safe via LO recalc for computed books).

    Args:
        path: Absolute path to the workbook.
        name: New sheet name (must not already exist).

    Returns:
        {path, sheet, sheets, routed, ...}
    """
    from microapple_sheet import bridge, ops_openpyxl
    if bridge.is_open(path).get("is_open"):
        result = bridge.live_add_sheet(path, name)
        result["routed"] = "live"
        return result
    result = ops_openpyxl.add_sheet(path, name)
    result["routed"] = "file"
    return result


@mcp.tool()
def excel_delete_sheet(path: str, name: str) -> dict[str, Any]:
    """Delete a worksheet from a workbook (engine-routed).

    Open in Excel → live bridge (Excel confirm dialog suppressed).
    Closed → file engine. Refuses to delete the only sheet.

    Args:
        path: Absolute path to the workbook.
        name: Sheet to delete.

    Returns:
        {path, sheet, sheets, routed, ...}
    """
    from microapple_sheet import bridge, ops_openpyxl
    if bridge.is_open(path).get("is_open"):
        result = bridge.live_delete_sheet(path, name)
        result["routed"] = "live"
        return result
    result = ops_openpyxl.delete_sheet(path, name)
    result["routed"] = "file"
    return result


@mcp.tool()
def excel_rename_sheet(path: str, old_name: str, new_name: str) -> dict[str, Any]:
    """Rename a worksheet (engine-routed: live when open, file when closed).

    Args:
        path:     Absolute path to the workbook.
        old_name: Current sheet name.
        new_name: New sheet name (must not already exist).

    Returns:
        {path, old_name, new_name, sheets, routed, ...}
    """
    from microapple_sheet import bridge, ops_openpyxl
    if bridge.is_open(path).get("is_open"):
        result = bridge.live_rename_sheet(path, old_name, new_name)
        result["routed"] = "live"
        return result
    result = ops_openpyxl.rename_sheet(path, old_name, new_name)
    result["routed"] = "file"
    return result


@mcp.tool()
def excel_move_sheet(
    path: str,
    name: str,
    before: str | None = None,
    after: str | None = None,
) -> dict[str, Any]:
    """Reorder a worksheet before/after an anchor sheet (engine-routed).

    Provide exactly one of *before* / *after* (anchor sheet name).

    Args:
        path:   Absolute path to the workbook.
        name:   Sheet to move.
        before: Move *name* immediately before this anchor sheet.
        after:  Move *name* immediately after this anchor sheet.

    Returns:
        {path, sheet, sheets, routed, ...}
    """
    from microapple_sheet import bridge, ops_openpyxl
    if bridge.is_open(path).get("is_open"):
        result = bridge.live_move_sheet(path, name, before=before, after=after)
        result["routed"] = "live"
        return result
    result = ops_openpyxl.move_sheet(path, name, before=before, after=after)
    result["routed"] = "file"
    return result


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Run the excel-mcp server (stdio transport)."""
    mcp.run()


if __name__ == "__main__":
    main()
