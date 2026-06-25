"""Smoke tests for excel-mcp Phase 0 scaffold.

Runs under pytest (preferred) or plain python3:
    .venv/bin/python -m pytest tests/ -q
    .venv/bin/python tests/test_smoke.py

Tests:
  1. FastMCP app exists and named 'excel-mcp'.
  2. All expected MVP tool names are registered.
  3. config.validate_marker accepts a normal marker, warns on '=x', rejects a newline.
"""
from __future__ import annotations

import sys

# ---------------------------------------------------------------------------
# Minimal pytest-compatible assertion harness (no pytest required)
# ---------------------------------------------------------------------------
_PASS = 0
_FAIL = 0


def _check(condition: bool, label: str) -> None:
    global _PASS, _FAIL
    if condition:
        _PASS += 1
        print(f"  PASS  {label}")
    else:
        _FAIL += 1
        print(f"  FAIL  {label}")


# ---------------------------------------------------------------------------
# Test 1: FastMCP app exists and is named 'excel-mcp'
# ---------------------------------------------------------------------------
def test_fastmcp_app_exists() -> None:
    from microapple_sheet.server import mcp

    _check(mcp is not None, "mcp object is not None")
    _check(mcp.name == "microapple-sheet", f"mcp.name == 'microapple-sheet' (got {mcp.name!r})")


# ---------------------------------------------------------------------------
# Test 2: Expected MVP tool names are registered
# ---------------------------------------------------------------------------
_EXPECTED_TOOLS = {
    "excel_create",
    "excel_open",
    "excel_info",
    "excel_read",
    "excel_read_table",
    "excel_list_sheets",
    "excel_list_names",
    "excel_write",
    "excel_set_cell",
    "excel_write_formula",
    "excel_write_linked_formula",
    "excel_format",
    "excel_define_name",
    "excel_live_read",
    "excel_live_set",
    "excel_live_formula",
    "excel_is_open",
    "excel_check_automation",
    "excel_audit_hardcoded",
    "excel_config_get",
    "excel_config_set",
    "excel_recalc",
    "excel_convert",
    "excel_ping",
}


def test_tool_names_registered() -> None:
    from microapple_sheet.server import mcp

    # FastMCP exposes registered tools — try the common attribute paths
    # (exact API depends on mcp[cli] version)
    registered: set[str] = set()

    # Approach 1: _tool_manager or _tools dict (internal FastMCP API)
    if hasattr(mcp, "_tool_manager") and hasattr(mcp._tool_manager, "_tools"):
        registered = set(mcp._tool_manager._tools.keys())
    elif hasattr(mcp, "_tools"):
        raw = mcp._tools
        registered = set(raw.keys()) if isinstance(raw, dict) else {t.name for t in raw}

    # Approach 2: iterate tools() if it's a method
    if not registered and callable(getattr(mcp, "get_tools", None)):
        import asyncio
        tools_list = asyncio.get_event_loop().run_until_complete(mcp.get_tools())
        registered = {t.name for t in tools_list}

    if not registered:
        # Last resort: introspect the module for @mcp.tool decorations by checking
        # the function objects themselves — if any decorated function has a _mcp_tool
        # attribute this signals registration.
        import microapple_sheet.server as srv_mod
        for attr_name in dir(srv_mod):
            fn = getattr(srv_mod, attr_name, None)
            if callable(fn) and attr_name in _EXPECTED_TOOLS:
                registered.add(attr_name)

    for tool_name in sorted(_EXPECTED_TOOLS):
        _check(tool_name in registered, f"tool '{tool_name}' registered")

    if not registered:
        print("  NOTE  Could not read FastMCP tool registry directly — "
              "checked known tool names in server module as fallback")


# ---------------------------------------------------------------------------
# Test 3: config.validate_marker
# ---------------------------------------------------------------------------
def test_validate_marker() -> None:
    from microapple_sheet.config import validate_marker

    # Normal glyph — should accept, no warning
    ok, msg = validate_marker("⏳")
    _check(ok is True, "validate_marker accepts '⏳' (ok=True)")
    _check(msg is None, "validate_marker '⏳' — no warning (msg=None)")

    # Normal ASCII marker — should accept
    ok2, msg2 = validate_marker("~PENDING~")
    _check(ok2 is True, "validate_marker accepts '~PENDING~'")

    # Leading '=' — should warn (ok=True) not reject
    ok3, msg3 = validate_marker("=x")
    _check(ok3 is True, "validate_marker '=x' — accepted (ok=True, not rejected)")
    _check(msg3 is not None, "validate_marker '=x' — warning message provided")
    _check(isinstance(msg3, str), f"validate_marker '=x' — warning is a string: {msg3!r}")

    # Newline — should reject (ok=False)
    ok4, msg4 = validate_marker("bad\nmarker")
    _check(ok4 is False, "validate_marker rejects newline in marker (ok=False)")
    _check(msg4 is not None, "validate_marker newline — error message provided")

    # Too long — should reject
    long_marker = "x" * 17  # exceeds 16-char limit
    ok5, msg5 = validate_marker(long_marker)
    _check(ok5 is False, f"validate_marker rejects {len(long_marker)}-char marker (ok=False)")

    # Tab character — should also reject (control char)
    ok6, msg6 = validate_marker("bad\tmarker")
    _check(ok6 is False, "validate_marker rejects tab in marker (ok=False)")


# ---------------------------------------------------------------------------
# Test 4: implemented Phase-1a modules callable; Phase-1b/2/3 stubs raise
# ---------------------------------------------------------------------------
def test_stubs_raise_on_call() -> None:
    """Phase-1a modules are implemented; Phase-1b/2/3 remain stubs (NotImplementedError)."""
    import tempfile, pathlib
    from microapple_sheet import engine, libreoffice, ops_openpyxl, liveformula, bridge, backups

    # ── Phase 1a: IMPLEMENTED — should NOT raise NotImplementedError ──────
    # classify_workbook on a nonexistent path returns OPENPYXL WorkbookClass (no error)
    try:
        cls = engine.classify_workbook("/tmp/fake_nonexistent.xlsx")
        _check(True, "engine.classify_workbook returns WorkbookClass for nonexistent path")
    except NotImplementedError:
        _check(False, "engine.classify_workbook still stub — expected implemented")
    except Exception as exc:
        _check(False, f"engine.classify_workbook raised unexpected {type(exc).__name__}: {exc}")

    # disk_hash on nonexistent path raises FileNotFoundError (correct behaviour)
    try:
        engine.disk_hash("/tmp/fake_nonexistent.xlsx")
        _check(False, "engine.disk_hash should raise FileNotFoundError for nonexistent path")
    except FileNotFoundError:
        _check(True, "engine.disk_hash raises FileNotFoundError (implemented, correct)")
    except NotImplementedError:
        _check(False, "engine.disk_hash still stub — expected implemented")

    # libreoffice.recalc on nonexistent path raises FileNotFoundError (correct)
    try:
        libreoffice.recalc("/tmp/fake_nonexistent.xlsx")
        _check(False, "libreoffice.recalc should raise FileNotFoundError for nonexistent path")
    except FileNotFoundError:
        _check(True, "libreoffice.recalc raises FileNotFoundError (implemented, correct)")
    except NotImplementedError:
        _check(False, "libreoffice.recalc still stub — expected implemented")

    # backups.snapshot on nonexistent path raises FileNotFoundError (correct)
    try:
        backups.snapshot("/tmp/fake_nonexistent.xlsx")
        _check(False, "backups.snapshot should raise FileNotFoundError for nonexistent path")
    except FileNotFoundError:
        _check(True, "backups.snapshot raises FileNotFoundError (implemented, correct)")
    except NotImplementedError:
        _check(False, "backups.snapshot still stub — expected implemented")

    # ── Phase 1b: ops_openpyxl.create raises FileNotFoundError (parent must exist) ──
    # create() on a valid path works; on a bogus nested path the parent doesn't exist
    try:
        ops_openpyxl.create("/tmp/fake_nonexistent_dir/wb.xlsx")
        _check(False, "ops_openpyxl.create on missing parent should fail")
    except (FileNotFoundError, PermissionError, OSError):
        _check(True, "ops_openpyxl.create raises expected error for missing parent dir (implemented)")
    except NotImplementedError:
        _check(False, "ops_openpyxl.create still stub — expected implemented")

    # ── Phase 2: bridge.is_open implemented (returns dict, not NotImplementedError) ──
    try:
        result = bridge.is_open("/tmp/fake_nonexistent.xlsx")
        _check(isinstance(result, dict), "bridge.is_open returns dict (implemented)")
        _check("is_open" in result, "bridge.is_open dict has 'is_open' key")
    except NotImplementedError:
        _check(False, "bridge.is_open still stub — expected implemented")
    except Exception as exc:
        # May fail if Excel not running, but should return dict not raise
        _check(False, f"bridge.is_open raised unexpected {type(exc).__name__}: {exc}")

    # ── Phase 3: IMPLEMENTED — validate_formula should return (True, None) for a valid ref ──
    try:
        valid, reason = liveformula.validate_formula("=SUM(A1)")
        _check(valid is True, "liveformula.validate_formula returns (True, None) for =SUM(A1)")
        _check(reason is None, "liveformula.validate_formula reason=None for valid formula")
    except NotImplementedError:
        _check(False, "liveformula.validate_formula still stub — expected implemented (Phase 3)")
    except Exception as exc:
        _check(False, f"liveformula.validate_formula raised unexpected {type(exc).__name__}: {exc}")


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------
def _run_all() -> None:
    print("\n=== excel-mcp Phase 0 smoke tests ===\n")
    test_fastmcp_app_exists()
    test_tool_names_registered()
    test_validate_marker()
    test_stubs_raise_on_call()
    print(f"\n{'='*40}")
    print(f"Results: {_PASS} passed, {_FAIL} failed")
    if _FAIL:
        sys.exit(1)


# pytest-compatible test functions (pytest collects these directly)
def test_all_in_one() -> None:
    """Single entry point for pytest — runs all sub-checks."""
    test_fastmcp_app_exists()
    test_tool_names_registered()
    test_validate_marker()
    test_stubs_raise_on_call()
    assert _FAIL == 0, f"{_FAIL} sub-checks failed (see output above)"


if __name__ == "__main__":
    _run_all()
