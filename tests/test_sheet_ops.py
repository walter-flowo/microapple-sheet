"""Sheet-management ops (add / delete / move / rename) — file-engine tests.

Covers the openpyxl path (no Excel required). The live AppleScript path is
integration-tested manually against an open workbook.
"""
from __future__ import annotations

import os

import pytest

from microapple_sheet import ops_openpyxl as ops


def _wb(tmp_path, sheets):
    p = os.path.join(str(tmp_path), "t.xlsx")
    ops.create(p, sheets)
    return p


def test_add_sheet(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    assert ops.add_sheet(p, "C")["sheets"] == ["A", "B", "C"]


def test_add_sheet_at_index(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    assert ops.add_sheet(p, "C", index=0)["sheets"] == ["C", "A", "B"]


def test_add_duplicate_raises(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    with pytest.raises(ValueError):
        ops.add_sheet(p, "B")


def test_rename_sheet(tmp_path):
    p = _wb(tmp_path, ["A", "B", "C"])
    r = ops.rename_sheet(p, "B", "B2")
    assert "B2" in r["sheets"] and "B" not in r["sheets"]


def test_rename_to_existing_raises(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    with pytest.raises(ValueError):
        ops.rename_sheet(p, "A", "B")


def test_rename_missing_raises(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    with pytest.raises(ValueError):
        ops.rename_sheet(p, "ZZ", "Y")


def test_move_sheet_before(tmp_path):
    p = _wb(tmp_path, ["A", "B", "C"])
    assert ops.move_sheet(p, "C", before="A")["sheets"] == ["C", "A", "B"]


def test_move_sheet_after(tmp_path):
    p = _wb(tmp_path, ["A", "B", "C"])
    assert ops.move_sheet(p, "A", after="C")["sheets"] == ["B", "C", "A"]


def test_move_missing_anchor_raises(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    with pytest.raises(ValueError):
        ops.move_sheet(p, "A", before="ZZ")


def test_delete_sheet(tmp_path):
    p = _wb(tmp_path, ["A", "B", "C"])
    assert ops.delete_sheet(p, "B")["sheets"] == ["A", "C"]


def test_delete_last_raises(tmp_path):
    p = _wb(tmp_path, ["A"])
    with pytest.raises(ValueError):
        ops.delete_sheet(p, "A")


def test_delete_missing_raises(tmp_path):
    p = _wb(tmp_path, ["A", "B"])
    with pytest.raises(ValueError):
        ops.delete_sheet(p, "ZZ")
