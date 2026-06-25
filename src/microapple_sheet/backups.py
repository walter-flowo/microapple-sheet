"""backups.py — timestamped workbook snapshots and restoration.

Phase 1a — IMPLEMENTED.

Every file-engine write must call snapshot(path) BEFORE touching the file.
Snapshots are stored in .backup/ adjacent to the workbook and pruned automatically
to keep the most recent N copies (default 10).
"""
from __future__ import annotations

import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

#: Default number of snapshots to keep per workbook
DEFAULT_KEEP = 10


def _backup_dir(path: Path) -> Path:
    """Return the .backup/ directory adjacent to *path*."""
    return path.parent / ".backup"


def _backup_stem(path: Path) -> str:
    """Return the stem used for backup filenames (workbook stem without extension)."""
    return path.stem


def _ts_now() -> str:
    """UTC ISO-8601 timestamp safe for use in filenames (no colons)."""
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def snapshot(path: str | Path) -> str:
    """Create a timestamped backup of *path* in the .backup/ subdirectory.

    Uses shutil.copy2 to preserve mtime metadata.
    Creates .backup/ directory if it doesn't exist.
    Calls prune() after the snapshot to respect DEFAULT_KEEP.

    Args:
        path: Absolute path to the workbook to back up. Must exist.

    Returns:
        Absolute path to the backup copy (str).

    Raises:
        FileNotFoundError: If *path* does not exist.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Cannot snapshot — file not found: {p}")

    bdir = _backup_dir(p)
    bdir.mkdir(parents=True, exist_ok=True)

    ts = _ts_now()
    backup_name = f"{_backup_stem(p)}.{ts}{p.suffix}"
    backup_path = bdir / backup_name

    shutil.copy2(p, backup_path)
    prune(p, keep=DEFAULT_KEEP)

    return str(backup_path)


def restore(path: str | Path, backup: str | Path) -> dict[str, Any]:
    """Replace workbook at *path* with the contents of *backup*.

    Args:
        path:   Absolute path to the workbook to restore (will be overwritten).
        backup: Absolute path to the backup copy to restore from.

    Returns:
        {path: str, restored_from: str, ok: bool}

    Raises:
        FileNotFoundError: If *backup* does not exist.
    """
    p = Path(path)
    b = Path(backup)

    if not b.exists():
        raise FileNotFoundError(f"Backup file not found: {b}")

    shutil.copy2(b, p)
    return {"path": str(p), "restored_from": str(b), "ok": True}


def list_backups(path: str | Path) -> list[str]:
    """Return absolute paths to all backups of *path*, newest first.

    Returns an empty list if no backups exist.
    """
    p = Path(path)
    bdir = _backup_dir(p)
    stem = _backup_stem(p)
    suffix = p.suffix

    if not bdir.exists():
        return []

    # Match files like <stem>.<timestamp><suffix>
    pattern = f"{stem}.*{suffix}"
    candidates = sorted(bdir.glob(pattern), key=lambda f: f.stat().st_mtime, reverse=True)
    return [str(c) for c in candidates]


def prune(path: str | Path, keep: int = DEFAULT_KEEP) -> int:
    """Delete oldest snapshots of *path* beyond the most recent *keep* copies.

    Args:
        path: Absolute path to the workbook whose backups are being pruned.
        keep: Number of most-recent backups to retain.

    Returns:
        Number of backup files deleted.
    """
    all_backups = list_backups(path)
    to_delete = all_backups[keep:]
    deleted = 0
    for bp in to_delete:
        try:
            Path(bp).unlink()
            deleted += 1
        except OSError:
            pass
    return deleted
