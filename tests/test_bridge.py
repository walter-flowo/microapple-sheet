"""test_bridge.py — unit tests for the AppleScript live bridge.

These tests do NOT require Microsoft Excel to be running.  All osascript
calls are intercepted by a mock runner injected via bridge._run_with_edit_retry's
``_runner`` parameter or by patching bridge._run_osascript.

Coverage:
  1. _encode_value round-trips (Python → tag → AppleScript argv representation).
  2. _coerce_cell (AppleScript TSV output → Python values).
  3. _pack_set_argv: flat argv encoding for 1×1, 1×3, 2×3 arrays.
  4. _parse_tsv_output: reconstruct 2D data from mock AppleScript output.
  5. _classify_error: TCC denial / not-running / not-open / timeout / generic.
  6. _run_with_edit_retry — mock runner:
       a. Succeeds on first attempt → no sleep.
       b. Fails with transient error, succeeds on 3rd attempt → retries.
       c. Persistent transient error → fail-fast with edit_mode_hint=True.
       d. TCC error → fail immediately, no retry.
  7. is_open parsing with mocked _run_osascript:
       a. NOT_RUNNING stdout → is_open=False.
       b. Workbook in list → is_open=True, workbook_name set.
       c. No match → is_open=False.
  8. check_automation with mocked runner:
       a. NOT_RUNNING → automation_ok=True, excel_running=False.
       b. OK:<version> → automation_ok=True, excel_running=True.
       c. TCC error → automation_ok=False.
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, call, patch

_HERE = Path(__file__).resolve().parent
_ROOT = _HERE.parent
_SRC = _ROOT / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

import pytest

from microapple_sheet.bridge import (
    _ERR_GENERIC,
    _ERR_NOT_OPEN,
    _ERR_NOT_RUNNING,
    _ERR_TCC_DENIED,
    _ERR_TIMEOUT,
    _classify_error,
    _coerce_cell,
    _encode_value,
    _pack_set_argv,
    _parse_tsv_output,
    _run_with_edit_retry,
    _values_match,
    check_automation,
    is_open,
)


# ── 1. _encode_value ──────────────────────────────────────────────────────────

class TestEncodeValue:
    def test_none(self) -> None:
        assert _encode_value(None) == "E:"

    def test_int(self) -> None:
        enc = _encode_value(42)
        assert enc == "N:42"

    def test_float(self) -> None:
        enc = _encode_value(3.14)
        assert enc.startswith("N:")
        assert "3.14" in enc

    def test_bool_true(self) -> None:
        assert _encode_value(True) == "B:true"

    def test_bool_false(self) -> None:
        assert _encode_value(False) == "B:false"

    def test_string(self) -> None:
        assert _encode_value("hello") == "S:hello"

    def test_formula_still_string_tagged(self) -> None:
        # Formulas passed as values are S-tagged; live_formula handles "=" directly
        enc = _encode_value("=SUM(A1:A10)")
        assert enc == "S:=SUM(A1:A10)"

    def test_zero(self) -> None:
        assert _encode_value(0) == "N:0"

    def test_empty_string(self) -> None:
        # Empty string → S: tag (not E:, which is for None)
        enc = _encode_value("")
        assert enc == "S:"


# ── 2. _coerce_cell ───────────────────────────────────────────────────────────

class TestCoerceCell:
    def test_empty_string_is_none(self) -> None:
        assert _coerce_cell("") is None

    def test_missing_token_is_none(self) -> None:
        assert _coerce_cell("MISSING") is None

    def test_integer_string(self) -> None:
        assert _coerce_cell("42") == 42
        assert isinstance(_coerce_cell("42"), int)

    def test_float_string(self) -> None:
        assert _coerce_cell("3.14") == pytest.approx(3.14)
        assert isinstance(_coerce_cell("3.14"), float)

    def test_plain_text(self) -> None:
        assert _coerce_cell("hello") == "hello"

    def test_negative_int(self) -> None:
        assert _coerce_cell("-5") == -5

    def test_scientific_notation(self) -> None:
        v = _coerce_cell("1.5E+3")
        assert v == pytest.approx(1500.0)

    def test_formula_string_preserved(self) -> None:
        # Formulas should come back as strings (read in values mode gives the computed value,
        # in formulas mode the formula itself — both handled as plain strings here)
        assert _coerce_cell("=SUM(A1)") == "=SUM(A1)"


# ── 3. _pack_set_argv ─────────────────────────────────────────────────────────

class TestPackSetArgv:
    def test_1x1_scalar(self) -> None:
        argv = _pack_set_argv("/path/wb.xlsx", "Sheet1", "A1", [[42]])
        assert argv[0] == "/path/wb.xlsx"
        assert argv[1] == "Sheet1"
        assert argv[2] == "A1"
        assert argv[3] == "1"   # nrows
        assert argv[4] == "1"   # ncols
        assert argv[5] == "N:42"
        assert len(argv) == 6

    def test_1x3_row(self) -> None:
        argv = _pack_set_argv("/p/wb.xlsx", "S", "A1", [[10, 20, 30]])
        assert argv[3] == "1"   # nrows
        assert argv[4] == "3"   # ncols
        assert argv[5] == "N:10"
        assert argv[6] == "N:20"
        assert argv[7] == "N:30"
        assert len(argv) == 8

    def test_2x3_matrix_row_major(self) -> None:
        values = [[1, 2, 3], [4, 5, 6]]
        argv = _pack_set_argv("/p/wb.xlsx", "S", "A1:C2", values)
        assert argv[3] == "2"
        assert argv[4] == "3"
        # Row-major flat: 1,2,3,4,5,6
        encoded_vals = argv[5:]
        assert encoded_vals == ["N:1", "N:2", "N:3", "N:4", "N:5", "N:6"]

    def test_mixed_types(self) -> None:
        argv = _pack_set_argv("/p/wb.xlsx", "S", "A1:C1", [["hello", None, True]])
        assert argv[5] == "S:hello"
        assert argv[6] == "E:"
        assert argv[7] == "B:true"

    def test_path_prefix_args(self) -> None:
        """First 5 args: path, sheet, range, nrows, ncols."""
        argv = _pack_set_argv("/some/path/model.xlsx", "Inputs", "B4:D4", [[1, 2, 3]])
        assert argv[:5] == ["/some/path/model.xlsx", "Inputs", "B4:D4", "1", "3"]


# ── 4. _parse_tsv_output ──────────────────────────────────────────────────────

class TestParseTsvOutput:
    def test_single_cell(self) -> None:
        text = "1\t1\n42\n"
        data = _parse_tsv_output(text)
        assert data == [[42]]

    def test_single_row(self) -> None:
        text = "1\t3\n10\t20\t30\n"
        data = _parse_tsv_output(text)
        assert data == [[10, 20, 30]]

    def test_multi_row(self) -> None:
        text = "2\t3\n1\t2\t3\n4\t5\t6\n"
        data = _parse_tsv_output(text)
        assert data == [[1, 2, 3], [4, 5, 6]]

    def test_missing_values(self) -> None:
        text = "1\t3\nMISSING\t7\tMISSING\n"
        data = _parse_tsv_output(text)
        assert data[0][0] is None
        assert data[0][1] == 7
        assert data[0][2] is None

    def test_formula_strings(self) -> None:
        text = "1\t2\n=SUM(A1:A10)\t=B1*2\n"
        data = _parse_tsv_output(text)
        assert data[0][0] == "=SUM(A1:A10)"
        assert data[0][1] == "=B1*2"

    def test_empty_output_returns_empty_list(self) -> None:
        data = _parse_tsv_output("")
        assert data == []

    def test_float_preserved(self) -> None:
        text = "1\t1\n3.14159\n"
        data = _parse_tsv_output(text)
        assert data[0][0] == pytest.approx(3.14159)


# ── 5. _classify_error ────────────────────────────────────────────────────────

class TestClassifyError:
    def test_tcc_denied_by_error_code(self) -> None:
        assert _classify_error("error -1743", 1) == _ERR_TCC_DENIED

    def test_tcc_denied_by_text(self) -> None:
        assert _classify_error("Not authorized to send Apple events", 1) == _ERR_TCC_DENIED

    def test_tcc_denied_is_not_allowed(self) -> None:
        assert _classify_error("osascript is not allowed to send keystrokes", 1) == _ERR_TCC_DENIED

    def test_not_running_by_code(self) -> None:
        assert _classify_error("error 1001 Microsoft Excel is not running", 1) == _ERR_NOT_RUNNING

    def test_not_running_by_text(self) -> None:
        assert _classify_error("application isn't running", 1) == _ERR_NOT_RUNNING

    def test_not_open_by_code(self) -> None:
        assert _classify_error("error 1002 Workbook not open", 1) == _ERR_NOT_OPEN

    def test_timeout(self) -> None:
        assert _classify_error("", -1, timed_out=True) == _ERR_TIMEOUT

    def test_generic(self) -> None:
        assert _classify_error("some random error", 1) == _ERR_GENERIC

    def test_generic_returncode_nonzero(self) -> None:
        assert _classify_error("", 1) == _ERR_GENERIC


# ── 6. _run_with_edit_retry ──────────────────────────────────────────────────

def _ok_result(stdout: str = "OK") -> dict[str, Any]:
    return {"ok": True, "stdout": stdout, "stderr": "", "returncode": 0, "timed_out": False}


def _fail_result(stderr: str = "some error", rc: int = 1) -> dict[str, Any]:
    return {"ok": False, "stdout": "", "stderr": stderr, "returncode": rc, "timed_out": False}


def _timeout_result() -> dict[str, Any]:
    return {"ok": False, "stdout": "", "stderr": "timed out", "returncode": -1, "timed_out": True}


class TestEditRetry:
    def test_success_on_first_attempt(self) -> None:
        """First call succeeds → runner called once, no sleep."""
        runner = MagicMock(return_value=_ok_result("OK"))
        result = _run_with_edit_retry(
            "script", ["a", "b"],
            max_retries=3, retry_delay=0.0, timeout=5,
            _runner=runner,
        )
        assert result["ok"] is True
        assert runner.call_count == 1

    def test_retries_on_transient_then_succeeds(self) -> None:
        """Fails twice with generic error, succeeds on 3rd attempt."""
        runner = MagicMock(side_effect=[
            _fail_result("generic AE error"),
            _fail_result("generic AE error"),
            _ok_result("OK"),
        ])
        result = _run_with_edit_retry(
            "script", [],
            max_retries=3, retry_delay=0.0, timeout=5,
            _runner=runner,
        )
        assert result["ok"] is True
        assert runner.call_count == 3

    def test_exhausts_retries_sets_edit_mode_hint(self) -> None:
        """Persistent generic errors → edit_mode_hint=True after max retries."""
        runner = MagicMock(return_value=_fail_result("generic AE error"))
        result = _run_with_edit_retry(
            "script", [],
            max_retries=3, retry_delay=0.0, timeout=5,
            _runner=runner,
        )
        assert result["ok"] is False
        assert result.get("edit_mode_hint") is True
        assert "Enter" in result.get("detail", "") or "Esc" in result.get("detail", "")
        assert runner.call_count == 3

    def test_tcc_error_fails_immediately_no_retry(self) -> None:
        """TCC denial → fail immediately without any retry sleep."""
        runner = MagicMock(return_value=_fail_result("error -1743"))
        result = _run_with_edit_retry(
            "script", [],
            max_retries=3, retry_delay=0.0, timeout=5,
            _runner=runner,
        )
        assert result["ok"] is False
        assert runner.call_count == 1  # no retry
        assert result.get("edit_mode_hint") is None  # NOT edit mode

    def test_not_running_fails_immediately(self) -> None:
        runner = MagicMock(return_value=_fail_result("error 1001 not running"))
        result = _run_with_edit_retry(
            "script", [],
            max_retries=3, retry_delay=0.0, timeout=5,
            _runner=runner,
        )
        assert runner.call_count == 1

    def test_timeout_triggers_retry(self) -> None:
        """Timeout is a transient error → retried."""
        runner = MagicMock(side_effect=[
            _timeout_result(),
            _ok_result("OK"),
        ])
        result = _run_with_edit_retry(
            "script", [],
            max_retries=3, retry_delay=0.0, timeout=5,
            _runner=runner,
        )
        assert result["ok"] is True
        assert runner.call_count == 2


# ── 7. is_open (mocked _run_osascript) ────────────────────────────────────────

class TestIsOpen:
    """Unit tests for is_open() — all osascript calls are mocked.

    _IS_MACOS is patched to True so these tests run on Linux too.
    """

    def test_not_running_returns_false(self) -> None:
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _ok_result("NOT_RUNNING")
            result = is_open("/some/path/model.xlsx")
        assert result["is_open"] is False
        assert result["workbook_name"] is None
        assert "not running" in result["detail"].lower()

    def test_workbook_in_list_full_path_match(self) -> None:
        wb_path = "/Users/example/Documents/model.xlsx"
        workbook_list = f"{wb_path}\tmodel.xlsx\n"
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _ok_result(workbook_list)
            result = is_open(wb_path)
        assert result["is_open"] is True
        assert result["workbook_name"] == "model.xlsx"

    def test_workbook_basename_fallback(self) -> None:
        """iCloud path may differ; basename match should still work."""
        real_path = "/Users/example/Library/Mobile Documents/com~apple~CloudDocs/model.xlsx"
        excel_path = "/private/var/folders/model.xlsx"  # Excel sees a different full path
        workbook_list = f"{excel_path}\tmodel.xlsx\n"
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _ok_result(workbook_list)
            # Use basename-only path so full path won't match but basename will
            result = is_open("/some/other/path/model.xlsx")
        assert result["is_open"] is True
        assert result["workbook_name"] == "model.xlsx"
        assert "basename" in result["detail"].lower()

    def test_workbook_not_found(self) -> None:
        workbook_list = "/some/other/file.xlsx\tother.xlsx\n"
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _ok_result(workbook_list)
            result = is_open("/target/model.xlsx")
        assert result["is_open"] is False
        assert result["workbook_name"] is None

    def test_tcc_denied(self) -> None:
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _fail_result("error -1743")
            result = is_open("/some/path/model.xlsx")
        assert result["is_open"] is False
        assert "-1743" in result["detail"] or "denied" in result["detail"].lower()

    def test_not_macos_returns_false(self) -> None:
        """On non-macOS is_open() returns is_open=False with a structured message."""
        with patch("microapple_sheet.bridge._IS_MACOS", False):
            result = is_open("/some/path/model.xlsx")
        assert result["is_open"] is False
        assert "macOS" in result["detail"] or "live bridge" in result["detail"].lower()


# ── 8. check_automation (mocked _run_osascript) ───────────────────────────────

class TestCheckAutomation:
    """Unit tests for check_automation() — all osascript calls are mocked.

    _IS_MACOS is patched to True so these tests run on Linux too.
    """

    def test_not_running(self) -> None:
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _ok_result("NOT_RUNNING")
            result = check_automation()
        assert result["automation_ok"] is True
        assert result["excel_running"] is False

    def test_excel_running(self) -> None:
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _ok_result("OK:16.82.3")
            result = check_automation()
        assert result["automation_ok"] is True
        assert result["excel_running"] is True
        assert "16.82.3" in result["detail"]

    def test_tcc_denied(self) -> None:
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _fail_result("error -1743 not authorized")
            result = check_automation()
        assert result["automation_ok"] is False
        assert result["excel_running"] is False
        assert "System Settings" in result["detail"] or "-1743" in result["detail"]

    def test_generic_error(self) -> None:
        with patch("microapple_sheet.bridge._IS_MACOS", True), \
             patch("microapple_sheet.bridge._run_osascript") as mock_run:
            mock_run.return_value = _fail_result("some weird error", rc=1)
            result = check_automation()
        assert result["automation_ok"] is False

    def test_not_macos_returns_structured_error(self) -> None:
        """On non-macOS check_automation() returns automation_ok=False with a clear message."""
        with patch("microapple_sheet.bridge._IS_MACOS", False):
            result = check_automation()
        assert result["automation_ok"] is False
        assert result["excel_running"] is False
        assert "macOS" in result["detail"] or "live bridge" in result["detail"].lower()


# ── 9. _values_match ──────────────────────────────────────────────────────────

class TestValuesMatch:
    def test_equal_ints(self) -> None:
        from microapple_sheet.bridge import _values_match
        assert _values_match(42, 42) is True

    def test_float_tolerance(self) -> None:
        from microapple_sheet.bridge import _values_match
        assert _values_match(100.0, 100.000001) is True

    def test_string_match(self) -> None:
        from microapple_sheet.bridge import _values_match
        assert _values_match("hello", "hello") is True

    def test_mismatch(self) -> None:
        from microapple_sheet.bridge import _values_match
        assert _values_match(1, 2) is False

    def test_none_vs_none(self) -> None:
        from microapple_sheet.bridge import _values_match
        assert _values_match(None, None) is True

    def test_none_vs_value(self) -> None:
        from microapple_sheet.bridge import _values_match
        assert _values_match(None, 1) is False
