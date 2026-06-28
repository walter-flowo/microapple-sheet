"""bridge.py — AppleScript live bridge for Microsoft Excel.

Transport: osascript subprocess with `on run argv` scripts.
All cell refs, sheet names, values, and formulas are passed as argv strings —
never string-interpolated into the script source (injection-safe).

Error taxonomy:
  tcc_denied   — macOS Automation permission refused (error -1743)
  not_running  — Excel process is not running
  not_open     — workbook not found among open workbooks
  timeout      — subprocess timed out (often indicates cell-edit mode)
  generic      — other AppleScript / AE error

Edit-mode handling:
  Transient failures (generic / timeout) trigger a retry loop (≤3 attempts,
  0.8 s spacing, ≤~2.5 s total). After exhausting retries, the response
  includes an ``edit_mode_hint`` flag and a user-readable message.

Routing axiom: open in Excel → live bridge ONLY; closed → file engine ONLY.
Never dual-write a file that Excel has open.
"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path
from typing import Any

# ── Platform check ────────────────────────────────────────────────────────────
# The AppleScript live bridge is macOS + Microsoft Excel only.
_IS_MACOS: bool = sys.platform == "darwin"

_NOT_MACOS_DETAIL = (
    "The live bridge requires macOS + Microsoft Excel. "
    "File-engine tools (excel_create, excel_read, excel_write, …) work cross-platform."
)

# ── Constants ─────────────────────────────────────────────────────────────────

_OSASCRIPT = "/usr/bin/osascript"
_DEFAULT_TIMEOUT = 15   # seconds
_EDIT_RETRY_COUNT = 3
_EDIT_RETRY_DELAY = 0.8  # seconds between attempts

# Error-type tokens
_ERR_TCC_DENIED = "tcc_denied"
_ERR_NOT_RUNNING = "not_running"
_ERR_NOT_OPEN = "not_open"
_ERR_TIMEOUT = "timeout"
_ERR_GENERIC = "generic"

# ── AppleScript scripts ───────────────────────────────────────────────────────

# Benign Excel AE probe.  Returns "NOT_RUNNING" or "OK:<version>".
_SCRIPT_CHECK_AUTOMATION = """\
on run argv
    if application "Microsoft Excel" is not running then
        return "NOT_RUNNING"
    end if
    tell application "Microsoft Excel"
        return "OK:" & (version as text)
    end tell
end run
"""

# List every open workbook: one line per wb → "full_path<TAB>name".
_SCRIPT_LIST_WORKBOOKS = """\
on run argv
    if application "Microsoft Excel" is not running then
        return "NOT_RUNNING"
    end if
    tell application "Microsoft Excel"
        set fullNames to full name of every workbook
        set wbNames to name of every workbook
        if fullNames is missing value then set fullNames to {}
        if wbNames is missing value then set wbNames to {}
    end tell
    set outStr to ""
    repeat with i from 1 to (count of fullNames)
        set outStr to outStr & (item i of fullNames as text) & tab & (item i of wbNames as text) & linefeed
    end repeat
    return outStr
end run
"""

# Read cell VALUES from an open workbook.
# argv: wbFullPath, sheetName, rangeName
# Output: "nrows<TAB>ncols<LF>row1col1<TAB>row1col2<LF>..."
_SCRIPT_LIVE_READ_VALUES = """\
on run argv
    set wbFullPath to item 1 of argv
    set sheetName to item 2 of argv
    set rangeName to item 3 of argv
    if application "Microsoft Excel" is not running then
        error "Microsoft Excel is not running" number 1001
    end if
    tell application "Microsoft Excel"
        set fullNames to full name of every workbook
        set wbNames to name of every workbook
        if fullNames is missing value then set fullNames to {}
        if wbNames is missing value then set wbNames to {}
    end tell
    set targetName to missing value
    repeat with i from 1 to (count of fullNames)
        if (item i of fullNames as text) is wbFullPath then
            set targetName to item i of wbNames
            exit repeat
        end if
    end repeat
    if targetName is missing value then
        error "Workbook not open: " & wbFullPath number 1002
    end if
    tell application "Microsoft Excel"
        set ws to worksheet sheetName of workbook targetName
        set r to range rangeName of ws
        set rawVals to value of r
    end tell
    set tabCh to tab
    set lfCh to linefeed
    if class of rawVals is list then
        if (count of rawVals) > 0 and class of item 1 of rawVals is list then
            set vals to rawVals
        else
            set vals to {rawVals}
        end if
    else
        set vals to {{rawVals}}
    end if
    set nRows to count of vals
    set nCols to count of item 1 of vals
    set outStr to (nRows as text) & tabCh & (nCols as text) & lfCh
    repeat with theRow in vals
        set rowStr to ""
        set isFirst to true
        repeat with cellVal in theRow
            if not isFirst then set rowStr to rowStr & tabCh
            if cellVal is missing value then
                set rowStr to rowStr & "MISSING"
            else
                set rowStr to rowStr & (cellVal as text)
            end if
            set isFirst to false
        end repeat
        set outStr to outStr & rowStr & lfCh
    end repeat
    return outStr
end run
"""

# Read cell FORMULAS from an open workbook (same shape as values script).
_SCRIPT_LIVE_READ_FORMULAS = """\
on run argv
    set wbFullPath to item 1 of argv
    set sheetName to item 2 of argv
    set rangeName to item 3 of argv
    if application "Microsoft Excel" is not running then
        error "Microsoft Excel is not running" number 1001
    end if
    tell application "Microsoft Excel"
        set fullNames to full name of every workbook
        set wbNames to name of every workbook
        if fullNames is missing value then set fullNames to {}
        if wbNames is missing value then set wbNames to {}
    end tell
    set targetName to missing value
    repeat with i from 1 to (count of fullNames)
        if (item i of fullNames as text) is wbFullPath then
            set targetName to item i of wbNames
            exit repeat
        end if
    end repeat
    if targetName is missing value then
        error "Workbook not open: " & wbFullPath number 1002
    end if
    tell application "Microsoft Excel"
        set ws to worksheet sheetName of workbook targetName
        set r to range rangeName of ws
        set rawVals to formula of r
    end tell
    -- Build the TSV OUTSIDE the Excel tell block: inside it, `tab`/`linefeed`
    -- resolve to Excel dictionary terms (the literal word "tab") not the
    -- AppleScript character constants. Capture them here, as the values script does.
    set tabCh to tab
    set lfCh to linefeed
    if class of rawVals is list then
        if (count of rawVals) > 0 and class of item 1 of rawVals is list then
            set vals to rawVals
        else
            set vals to {rawVals}
        end if
    else
        set vals to {{rawVals}}
    end if
    set nRows to count of vals
    set nCols to count of item 1 of vals
    set outStr to (nRows as text) & tabCh & (nCols as text) & lfCh
    repeat with theRow in vals
        set rowStr to ""
        set isFirst to true
        repeat with cellVal in theRow
            if not isFirst then set rowStr to rowStr & tabCh
            if cellVal is missing value then
                set rowStr to rowStr & ""
            else
                set rowStr to rowStr & (cellVal as text)
            end if
            set isFirst to false
        end repeat
        set outStr to outStr & rowStr & lfCh
    end repeat
    return outStr
end run
"""

# Set VALUES in an open workbook.  ONE Apple Event for the whole range.
# argv: wbFullPath, sheetName, rangeName, nRows, nCols, enc_val1, enc_val2, ...
# Each value is type-tagged: "N:<num>", "S:<str>", "B:<true|false>", "E:" (empty).
# Returns "OK" on success.
_SCRIPT_LIVE_SET = """\
on decodeVal(taggedStr)
    if length of taggedStr < 2 then return ""
    set tag to text 1 thru 2 of taggedStr
    if length of taggedStr > 2 then
        set v to text 3 thru -1 of taggedStr
    else
        set v to ""
    end if
    if tag is "N:" then
        return v as real
    else if tag is "B:" then
        if v is "true" then return true
        return false
    else if tag is "E:" then
        return ""
    else
        return v
    end if
end decodeVal

on run argv
    set wbFullPath to item 1 of argv
    set sheetName to item 2 of argv
    set rangeName to item 3 of argv
    set nRows to item 4 of argv as integer
    set nCols to item 5 of argv as integer
    if (count of argv) >= 6 then
        set flatVals to items 6 thru -1 of argv
    else
        set flatVals to {}
    end if

    if application "Microsoft Excel" is not running then
        error "Microsoft Excel is not running" number 1001
    end if

    tell application "Microsoft Excel"
        set fullNames to full name of every workbook
        set wbNames to name of every workbook
        if fullNames is missing value then set fullNames to {}
        if wbNames is missing value then set wbNames to {}
    end tell
    set targetName to missing value
    repeat with i from 1 to (count of fullNames)
        if (item i of fullNames as text) is wbFullPath then
            set targetName to item i of wbNames
            exit repeat
        end if
    end repeat
    if targetName is missing value then
        error "Workbook not open: " & wbFullPath number 1002
    end if
    tell application "Microsoft Excel"
        set ws to worksheet sheetName of workbook targetName
        -- rangeName is pre-expanded by the Python caller to the full nRows x nCols
        -- block, so a plain matrix assignment fills every cell (a 1-cell range
        -- would otherwise capture only the top-left value).
        if nRows = 1 and nCols = 1 then
            set value of range rangeName of ws to my decodeVal(item 1 of flatVals)
        else if nRows = 1 then
            set theRow to {}
            repeat with c from 1 to nCols
                set end of theRow to my decodeVal(item c of flatVals)
            end repeat
            set value of range rangeName of ws to theRow
        else
            set matrix to {}
            repeat with r from 1 to nRows
                set theRow to {}
                repeat with c from 1 to nCols
                    set idx to (r - 1) * nCols + c
                    set end of theRow to my decodeVal(item idx of flatVals)
                end repeat
                set end of matrix to theRow
            end repeat
            set value of range rangeName of ws to matrix
        end if
    end tell
    try
        display notification "Written to Excel — \xE2\x8C\x98S to persist" with title "excel-mcp"
    end try
    return "OK"
end run
"""

# Set a FORMULA in an open workbook (single cell, single AE).
# argv: wbFullPath, sheetName, cellRef, formulaStr
# Returns "OK" on success.
_SCRIPT_LIVE_FORMULA = """\
on run argv
    set wbFullPath to item 1 of argv
    set sheetName to item 2 of argv
    set cellRef to item 3 of argv
    set formulaStr to item 4 of argv

    if application "Microsoft Excel" is not running then
        error "Microsoft Excel is not running" number 1001
    end if

    tell application "Microsoft Excel"
        set fullNames to full name of every workbook
        set wbNames to name of every workbook
        if fullNames is missing value then set fullNames to {}
        if wbNames is missing value then set wbNames to {}
    end tell
    set targetName to missing value
    repeat with i from 1 to (count of fullNames)
        if (item i of fullNames as text) is wbFullPath then
            set targetName to item i of wbNames
            exit repeat
        end if
    end repeat
    if targetName is missing value then
        error "Workbook not open: " & wbFullPath number 1002
    end if
    tell application "Microsoft Excel"
        set ws to worksheet sheetName of workbook targetName
        set formula of range cellRef of ws to formulaStr
    end tell
    try
        display notification "Formula written to Excel — \xE2\x8C\x98S to persist" with title "excel-mcp"
    end try
    return "OK"
end run
"""


# ── Internal helpers ──────────────────────────────────────────────────────────

def _encode_value(v: Any) -> str:
    """Encode a Python value as a type-tagged argv string for AppleScript.

    Tags:
      N:<number>   — numeric (int or float)
      B:true/false — boolean
      E:           — None / empty
      S:<text>     — everything else (including formulas)
    """
    if v is None:
        return "E:"
    if isinstance(v, bool):
        return f"B:{'true' if v else 'false'}"
    if isinstance(v, (int, float)):
        return f"N:{v}"
    return f"S:{v}"


def _coerce_cell(s: str) -> int | float | str | None:
    """Coerce a tab-separated output token to a Python value."""
    if not s or s == "MISSING":
        return None
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return s


def _parse_tsv_output(text: str) -> list[list[Any]]:
    """Parse the TSV grid emitted by the live_read scripts.

    First line: ``"nrows\\tncols"``.  Subsequent lines: tab-separated values.
    """
    lines = [ln for ln in text.split("\n") if ln != ""]
    if not lines:
        return []
    header = lines[0].split("\t")
    if len(header) != 2:
        raise ValueError(f"Unexpected read output header: {lines[0]!r}")
    nrows, ncols = int(header[0]), int(header[1])
    result: list[list[Any]] = []
    for i in range(nrows):
        line = lines[i + 1] if (i + 1) < len(lines) else ""
        row = (line.split("\t") + ([""] * ncols))[:ncols]
        result.append([_coerce_cell(c) for c in row])
    return result


def _pack_set_argv(
    wb_path: str,
    sheet: str,
    range_: str,
    values: list[list[Any]],
) -> list[str]:
    """Build the full argv list for _SCRIPT_LIVE_SET."""
    nrows = len(values)
    ncols = len(values[0]) if values else 0
    flat = [_encode_value(v) for row in values for v in row]
    return [wb_path, sheet, range_, str(nrows), str(ncols), *flat]


def _run_osascript(
    script: str,
    argv: list[str],
    timeout: int = _DEFAULT_TIMEOUT,
) -> dict[str, Any]:
    """Run an AppleScript via osascript stdin transport.

    The script MUST start with ``on run argv … end run``.
    Arguments are passed after ``--`` so they arrive in AppleScript ``argv``
    without any string interpolation into the script source.

    Returns:
        {ok, stdout, stderr, returncode, timed_out}
    """
    try:
        proc = subprocess.run(
            [_OSASCRIPT, "-", *argv],
            input=script,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return {
            "ok": False,
            "stdout": "",
            "stderr": f"osascript timed out after {timeout}s",
            "returncode": -1,
            "timed_out": True,
        }
    return {
        "ok": proc.returncode == 0,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
        "returncode": proc.returncode,
        "timed_out": False,
    }


def _classify_error(
    stderr: str,
    returncode: int,
    timed_out: bool = False,
) -> str:
    """Classify an osascript failure into one of the error-type tokens."""
    if timed_out:
        return _ERR_TIMEOUT
    s = stderr.lower()
    if "-1743" in stderr or "not authorized" in s or "is not allowed" in s:
        return _ERR_TCC_DENIED
    if "1001" in stderr or "not running" in s or "isn't running" in s:
        return _ERR_NOT_RUNNING
    if "1002" in stderr or "not open" in s or "workbook not open" in s:
        return _ERR_NOT_OPEN
    return _ERR_GENERIC


def _run_with_edit_retry(
    script: str,
    argv: list[str],
    max_retries: int = _EDIT_RETRY_COUNT,
    retry_delay: float = _EDIT_RETRY_DELAY,
    timeout: int = _DEFAULT_TIMEOUT,
    *,
    _runner: Any = None,  # injectable for tests
) -> dict[str, Any]:
    """Run script with a brief retry loop to handle transient edit-mode failures.

    TCC / not-running / not-open errors fail immediately without retry.
    All other failures (generic AE errors, timeouts) are retried up to
    *max_retries* times with *retry_delay* seconds between attempts.

    After all retries are exhausted the result gains ``edit_mode_hint=True``
    and a ``detail`` suggesting the user press Enter/Esc.
    """
    runner = _runner if _runner is not None else _run_osascript
    last: dict[str, Any] = {}
    for attempt in range(max_retries):
        last = runner(script, argv, timeout=timeout)
        if last["ok"]:
            return last
        err_type = _classify_error(
            last["stderr"], last["returncode"], last.get("timed_out", False)
        )
        # Hard errors: fail immediately
        if err_type in (_ERR_TCC_DENIED, _ERR_NOT_RUNNING, _ERR_NOT_OPEN):
            return last
        # Transient: wait and retry
        if attempt < max_retries - 1:
            time.sleep(retry_delay)

    last["edit_mode_hint"] = True
    last["detail"] = (
        "Excel appears to be in cell-edit mode — press Enter or Esc to finish "
        "editing, then re-run."
    )
    return last


def _error_detail(result: dict[str, Any]) -> str:
    """Compose a human-readable error message from an osascript result."""
    if result.get("detail"):
        return result["detail"]
    err_type = _classify_error(
        result["stderr"], result["returncode"], result.get("timed_out", False)
    )
    if err_type == _ERR_TCC_DENIED:
        return (
            "Apple Events access denied (error -1743). "
            "Go to System Settings → Privacy & Security → Automation and enable "
            "access for the terminal / Claude Code process."
        )
    if err_type == _ERR_NOT_RUNNING:
        return "Microsoft Excel is not running. Open Excel and the target workbook first."
    if err_type == _ERR_NOT_OPEN:
        return (
            "Workbook is not open in Excel. "
            "Open it in Excel, or use a file-engine tool (excel_write) for disk-based edits."
        )
    if err_type == _ERR_TIMEOUT:
        return "osascript timed out — Excel may be in edit mode. Press Enter/Esc and retry."
    stderr = result.get("stderr", "").strip()
    return f"Apple Event error: {stderr or '(no detail)'}"


# ── Public API ────────────────────────────────────────────────────────────────

def check_automation() -> dict[str, Any]:
    """Probe macOS Automation (TCC) permissions by sending a benign Excel AE.

    Detects:
    - TCC denial (error -1743) → automation_ok=False, remediation hint.
    - Excel not running → automation_ok=True (TCC is fine), excel_running=False.
    - Excel running and responding → automation_ok=True, excel_running=True.
    - Non-macOS → automation_ok=False, structured error.

    Returns:
        {automation_ok: bool, excel_running: bool, detail: str}
    """
    if not _IS_MACOS:
        return {
            "automation_ok": False,
            "excel_running": False,
            "detail": _NOT_MACOS_DETAIL,
        }
    result = _run_osascript(_SCRIPT_CHECK_AUTOMATION, [], timeout=8)
    stdout = result["stdout"].strip()

    if result["ok"] and stdout.startswith("NOT_RUNNING"):
        return {
            "automation_ok": True,
            "excel_running": False,
            "detail": (
                "Automation permission OK. Microsoft Excel is not running — "
                "live-mode tools require an open workbook in Excel."
            ),
        }

    if result["ok"] and stdout.startswith("OK:"):
        xl_version = stdout[3:]
        return {
            "automation_ok": True,
            "excel_running": True,
            "detail": f"Automation OK. Microsoft Excel {xl_version} is running.",
        }

    # osascript failed
    err_type = _classify_error(
        result["stderr"], result["returncode"], result.get("timed_out", False)
    )
    if err_type == _ERR_TCC_DENIED:
        return {
            "automation_ok": False,
            "excel_running": False,
            "detail": (
                "Apple Events access denied (error -1743). "
                "Go to System Settings → Privacy & Security → Automation and enable "
                "access for the terminal / Claude Code process."
            ),
        }
    return {
        "automation_ok": False,
        "excel_running": False,
        "detail": _error_detail(result),
    }


def is_open(path: str | Path) -> dict[str, Any]:
    """Check whether *path* is currently open in Microsoft Excel.

    Matches by full resolved path first, then falls back to basename.
    On non-macOS returns is_open=False with a structured error.

    Returns:
        {path, is_open: bool, workbook_name: str|None, detail: str}
    """
    p = Path(path).resolve()
    if not _IS_MACOS:
        return {
            "path": str(p),
            "is_open": False,
            "workbook_name": None,
            "detail": _NOT_MACOS_DETAIL,
        }
    result = _run_osascript(_SCRIPT_LIST_WORKBOOKS, [], timeout=8)
    stdout = result["stdout"].strip()

    if not result["ok"]:
        err_type = _classify_error(
            result["stderr"], result["returncode"], result.get("timed_out", False)
        )
        if err_type == _ERR_TCC_DENIED:
            return {
                "path": str(p),
                "is_open": False,
                "workbook_name": None,
                "detail": _error_detail(result),
            }
        # Any other failure — can't determine; assume not open
        return {
            "path": str(p),
            "is_open": False,
            "workbook_name": None,
            "detail": _error_detail(result),
        }

    if stdout == "NOT_RUNNING":
        return {
            "path": str(p),
            "is_open": False,
            "workbook_name": None,
            "detail": "Microsoft Excel is not running.",
        }

    # Parse workbook list: full_path<TAB>name per line
    open_wbs: dict[str, str] = {}
    for line in stdout.split("\n"):
        line = line.strip()
        if "\t" in line:
            full_path, _, wb_name = line.partition("\t")
            open_wbs[full_path.strip()] = wb_name.strip()

    # Full-path match
    if str(p) in open_wbs:
        wb_name = open_wbs[str(p)]
        return {
            "path": str(p),
            "is_open": True,
            "workbook_name": wb_name,
            "detail": f"Open as '{wb_name}'.",
        }

    # Basename fallback (handles iCloud path vs real path divergence)
    basename = p.name
    for full_path, wb_name in open_wbs.items():
        if Path(full_path).name == basename:
            return {
                "path": str(p),
                "is_open": True,
                "workbook_name": wb_name,
                "detail": f"Open as '{wb_name}' (basename match — paths differ).",
            }

    return {
        "path": str(p),
        "is_open": False,
        "workbook_name": None,
        "detail": (
            f"Not open in Excel. "
            f"{len(open_wbs)} workbook(s) currently open in Excel."
        ),
    }


def live_read(
    path: str | Path,
    sheet: str,
    range_: str,
    mode: str = "values",
) -> dict[str, Any]:
    """Read live cell content from an open workbook (sees unsaved edits).

    Args:
        path:   Absolute path to the open workbook.
        sheet:  Sheet name.
        range_: A1-notation range.
        mode:   ``"values"`` (default) or ``"formulas"``.

    Returns:
        {path, sheet, range, data: list[list[Any]], source: 'live', mode}

    Raises:
        RuntimeError: On non-macOS or AppleScript failure.
    """
    if not _IS_MACOS:
        raise RuntimeError(_NOT_MACOS_DETAIL)
    p = Path(path).resolve()
    script = (
        _SCRIPT_LIVE_READ_FORMULAS if mode == "formulas" else _SCRIPT_LIVE_READ_VALUES
    )
    argv = [str(p), sheet, range_]

    result = _run_with_edit_retry(script, argv)

    if not result["ok"]:
        raise RuntimeError(_error_detail(result))

    data = _parse_tsv_output(result["stdout"])
    return {
        "path": str(p),
        "sheet": sheet,
        "range": range_,
        "data": data,
        "source": "live",
        "mode": mode,
    }


def live_set(
    path: str | Path,
    sheet: str,
    range_: str,
    values: list[list[Any]],
) -> dict[str, Any]:
    """Set cell values in an open workbook via ONE atomic Apple Event.

    Never calls Select or Activate. Verifies by read-back after write.
    Shows a macOS notification prompting ⌘S after write.

    Args:
        path:   Absolute path to the open workbook.
        sheet:  Sheet name.
        range_: Top-left cell or full range (A1 notation).
        values: Row-major 2-D list of values.

    Returns:
        {path, sheet, range, cells_written, verified, detail}

    Raises:
        RuntimeError: On non-macOS or AppleScript failure.
    """
    if not _IS_MACOS:
        raise RuntimeError(_NOT_MACOS_DETAIL)
    p = Path(path).resolve()
    nrows = len(values)
    ncols = len(values[0]) if values else 0

    # Expand a top-left cell ("S25") to the full nRows x nCols block ("S25:S40")
    # so the write covers every cell — assigning a matrix to a single-cell range
    # fills only the top-left (silent partial write). A full range passed in is
    # re-derived from its own top-left, so both call styles behave identically.
    from openpyxl.utils.cell import (
        coordinate_from_string,
        column_index_from_string,
        get_column_letter,
    )

    top_left = range_.split(":")[0]
    col_letter, row_num = coordinate_from_string(top_left)
    c0 = column_index_from_string(col_letter)
    full_range = (
        f"{top_left}:{get_column_letter(c0 + ncols - 1)}{row_num + nrows - 1}"
        if (nrows * ncols) > 1
        else top_left
    )

    argv = _pack_set_argv(str(p), sheet, full_range, values)
    cells_written = nrows * ncols

    # E10: a rapid sequential set can return ok=True yet not persist (Excel busy
    # mid-recalc). Strategy: write -> verify -> retry. E11: the read-back verify
    # is FORMULA-AWARE — a formula cell is verified when it COMPUTED (non-None,
    # non-error), not when the read-back value equals the formula string.
    def _verify_block() -> bool:
        try:
            rb = _run_osascript(
                _SCRIPT_LIVE_READ_VALUES, [str(p), sheet, full_range], timeout=8
            )
            if not rb["ok"]:
                return False
            rb_data = _parse_tsv_output(rb["stdout"])
            if len(rb_data) != nrows:
                return False
            for i in range(nrows):
                if len(rb_data[i]) != ncols:
                    return False
                for j in range(ncols):
                    inp = values[i][j]
                    got = rb_data[i][j]
                    if isinstance(inp, str) and inp.startswith("="):
                        if got is None or (
                            isinstance(got, str) and got.startswith("#")
                        ):
                            return False
                    elif not _values_match(got, inp):
                        return False
            return True
        except Exception:
            return False

    verified = False
    attempts = 0
    for attempts in range(1, _EDIT_RETRY_COUNT + 1):
        result = _run_with_edit_retry(_SCRIPT_LIVE_SET, argv)
        if not result["ok"]:
            raise RuntimeError(_error_detail(result))
        verified = _verify_block()
        if verified:
            break
        if attempts < _EDIT_RETRY_COUNT:
            time.sleep(_EDIT_RETRY_DELAY)

    detail = "Unsaved in Excel — ⌘S to persist."
    if not verified:
        detail += f" [WARNING: read-back unverified after {attempts} attempts]"
    return {
        "path": str(p),
        "sheet": sheet,
        "range": full_range,
        "cells_written": cells_written,
        "verified": verified,
        "attempts": attempts,
        "detail": detail,
    }


def live_formula(
    path: str | Path,
    sheet: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Set a formula in an open workbook via AppleScript (triggers instant recalc).

    Args:
        path:    Absolute path to the open workbook.
        sheet:   Sheet name.
        cell:    Target cell in A1 notation.
        formula: Must start with ``=``.

    Returns:
        {path, sheet, cell, formula, verified_value, detail}

    Raises:
        RuntimeError: On non-macOS or AppleScript failure.
    """
    if not _IS_MACOS:
        raise RuntimeError(_NOT_MACOS_DETAIL)
    if not formula.startswith("="):
        formula = f"={formula}"
    p = Path(path).resolve()
    argv = [str(p), sheet, cell, formula]

    result = _run_with_edit_retry(_SCRIPT_LIVE_FORMULA, argv)

    if not result["ok"]:
        raise RuntimeError(_error_detail(result))

    # Read back the computed value (best-effort)
    verified_value: Any = None
    try:
        rb = _run_osascript(
            _SCRIPT_LIVE_READ_VALUES, [str(p), sheet, cell], timeout=8
        )
        if rb["ok"]:
            rb_data = _parse_tsv_output(rb["stdout"])
            verified_value = rb_data[0][0] if rb_data and rb_data[0] else None
    except Exception:
        pass

    return {
        "path": str(p),
        "sheet": sheet,
        "cell": cell,
        "formula": formula,
        "verified_value": verified_value,
        "detail": "Unsaved in Excel — ⌘S to persist.",
    }


# ── Private helper ────────────────────────────────────────────────────────────

def _values_match(a: Any, b: Any, rtol: float = 1e-6) -> bool:
    """Loose equality check between a read-back value and an expected value."""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        denom = max(abs(float(b)), 1e-12)
        return abs(float(a) - float(b)) / denom < rtol
    return str(a) == str(b)


# ── Sheet management (live) — add / delete / move / rename ──────────────────
#
# Each script finds the open workbook by full path, performs the structural
# change, collects `name of every worksheet` into a LIST inside the Excel tell
# block, then builds the newline-joined output OUTSIDE the block — because
# inside `tell application "Microsoft Excel"`, `linefeed` resolves to an Excel
# dictionary term, not the AppleScript constant (the E8 class of bug).

_SHEET_FINDER = """\
    set wbFullPath to item 1 of argv
    if application "Microsoft Excel" is not running then
        error "Microsoft Excel is not running" number 1001
    end if
    tell application "Microsoft Excel"
        set fullNames to full name of every workbook
        set wbNames to name of every workbook
        if fullNames is missing value then set fullNames to {}
        if wbNames is missing value then set wbNames to {}
    end tell
    set targetName to missing value
    repeat with i from 1 to (count of fullNames)
        if (item i of fullNames as text) is wbFullPath then
            set targetName to item i of wbNames
            exit repeat
        end if
    end repeat
    if targetName is missing value then
        error "Workbook not open: " & wbFullPath number 1002
    end if
"""

_SHEET_LIST_TAIL = """\
    set lfCh to linefeed
    set outStr to ""
    repeat with nm in sheetList
        set outStr to outStr & (nm as text) & lfCh
    end repeat
    return outStr
end run
"""

_OP_ADD = """\
    set sheetName to item 2 of argv
    tell application "Microsoft Excel"
        set wb to workbook targetName
        if (exists worksheet sheetName of wb) then error "Sheet already exists: " & sheetName number 1003
        set newWs to make new worksheet at end of wb
        set name of newWs to sheetName
        set sheetList to name of every worksheet of wb
    end tell
"""

_OP_DELETE = """\
    set sheetName to item 2 of argv
    tell application "Microsoft Excel"
        set wb to workbook targetName
        if not (exists worksheet sheetName of wb) then error "Sheet not found: " & sheetName number 1004
        set display alerts to false
        delete worksheet sheetName of wb
        set display alerts to true
        set sheetList to name of every worksheet of wb
    end tell
"""

_OP_MOVE = """\
    set sheetName to item 2 of argv
    set posMode to item 3 of argv
    set anchorName to item 4 of argv
    tell application "Microsoft Excel"
        set wb to workbook targetName
        if posMode is "after" then
            move worksheet sheetName of wb to after worksheet anchorName of wb
        else
            move worksheet sheetName of wb to before worksheet anchorName of wb
        end if
        set sheetList to name of every worksheet of wb
    end tell
"""

_OP_RENAME = """\
    set oldName to item 2 of argv
    set newName to item 3 of argv
    tell application "Microsoft Excel"
        set wb to workbook targetName
        if not (exists worksheet oldName of wb) then error "Sheet not found: " & oldName number 1004
        set name of worksheet oldName of wb to newName
        set sheetList to name of every worksheet of wb
    end tell
"""

_SCRIPT_LIVE_ADD_SHEET = "on run argv\n" + _SHEET_FINDER + _OP_ADD + _SHEET_LIST_TAIL
_SCRIPT_LIVE_DELETE_SHEET = "on run argv\n" + _SHEET_FINDER + _OP_DELETE + _SHEET_LIST_TAIL
_SCRIPT_LIVE_MOVE_SHEET = "on run argv\n" + _SHEET_FINDER + _OP_MOVE + _SHEET_LIST_TAIL
_SCRIPT_LIVE_RENAME_SHEET = "on run argv\n" + _SHEET_FINDER + _OP_RENAME + _SHEET_LIST_TAIL


def _run_sheet_op(script: str, argv: list[str]) -> list[str]:
    """Run a sheet-management AppleScript; return the resulting sheet-name list."""
    if not _IS_MACOS:
        raise RuntimeError(_NOT_MACOS_DETAIL)
    result = _run_with_edit_retry(script, argv)
    if not result["ok"]:
        raise RuntimeError(_error_detail(result))
    return [s for s in result["stdout"].splitlines() if s.strip()]


def live_add_sheet(path: str | Path, name: str) -> dict[str, Any]:
    """Add a worksheet (at end) to an open workbook via AppleScript."""
    p = Path(path).resolve()
    sheets = _run_sheet_op(_SCRIPT_LIVE_ADD_SHEET, [str(p), name])
    return {"path": str(p), "sheet": name, "sheets": sheets, "source": "live",
            "detail": f"Added '{name}' at end — ⌘S to persist."}


def live_delete_sheet(path: str | Path, name: str) -> dict[str, Any]:
    """Delete a worksheet from an open workbook (alerts suppressed)."""
    p = Path(path).resolve()
    sheets = _run_sheet_op(_SCRIPT_LIVE_DELETE_SHEET, [str(p), name])
    return {"path": str(p), "sheet": name, "sheets": sheets, "source": "live",
            "detail": f"Deleted '{name}' — ⌘S to persist."}


def live_move_sheet(
    path: str | Path, name: str, before: str | None = None, after: str | None = None
) -> dict[str, Any]:
    """Move a worksheet before/after an anchor sheet in an open workbook."""
    p = Path(path).resolve()
    if after:
        argv = [str(p), name, "after", after]
    elif before:
        argv = [str(p), name, "before", before]
    else:
        raise ValueError("live_move_sheet requires before= or after=")
    sheets = _run_sheet_op(_SCRIPT_LIVE_MOVE_SHEET, argv)
    return {"path": str(p), "sheet": name, "sheets": sheets, "source": "live",
            "detail": f"Moved '{name}' — ⌘S to persist."}


def live_rename_sheet(path: str | Path, old_name: str, new_name: str) -> dict[str, Any]:
    """Rename a worksheet in an open workbook."""
    p = Path(path).resolve()
    sheets = _run_sheet_op(_SCRIPT_LIVE_RENAME_SHEET, [str(p), old_name, new_name])
    return {"path": str(p), "old_name": old_name, "new_name": new_name,
            "sheets": sheets, "source": "live",
            "detail": f"Renamed '{old_name}' → '{new_name}' — ⌘S to persist."}
