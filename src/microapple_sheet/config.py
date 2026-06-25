"""config.py — load, validate, and persist config.toml for excel-mcp.

Phase 1 (real implementation, not a stub).

Responsibilities:
  - get(key)                       Read a dot-path value from the parsed TOML.
  - set(key, value)                Write a value back to config.toml, preserving the
                                   literal TOML single-quoted string for preview.marker.
  - validate_marker(s)             Safety gate for the blast-radius pending marker:
                                     REJECT: control chars / newlines / length > 16
                                     WARN:   leading = + - @ (formula injection signal)
                                     OK:     everything else
"""
from __future__ import annotations

import re
import tomllib
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Config file location — sibling to this package at the project root
# ---------------------------------------------------------------------------

_HERE = Path(__file__).resolve()
_CONFIG_PATH: Path = _HERE.parents[2] / "config.toml"  # Excel_MCP/config.toml


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _load_raw() -> str:
    """Return the raw text of config.toml, raising FileNotFoundError if absent."""
    return _CONFIG_PATH.read_text(encoding="utf-8")


def _parse(raw: str) -> dict[str, Any]:
    return tomllib.loads(raw)


def _resolve_key(data: dict[str, Any], dotkey: str) -> tuple[dict[str, Any], str]:
    """Walk dot-separated key into nested dicts; return (leaf_dict, leaf_key).

    Raises KeyError if an intermediate key doesn't exist.
    """
    parts = dotkey.split(".")
    node = data
    for part in parts[:-1]:
        node = node[part]
    return node, parts[-1]


def _format_value_for_toml(value: Any, key: str) -> str:
    """Serialise a Python value to its TOML literal representation.

    Special case: preview.marker is written as a TOML literal string (single-quoted)
    so backslashes are preserved verbatim.
    """
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return repr(value)
    if isinstance(value, str):
        # Use TOML literal strings for marker (backslash-safe)
        if key == "marker":
            # Literal strings cannot contain single-quote; fall back to basic string if so
            if "'" not in value:
                return f"'{value}'"
        # Standard TOML basic string — escape required chars
        escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "\\r")
        return f'"{escaped}"'
    raise TypeError(f"Cannot serialise {type(value).__name__!r} to TOML inline value")


def _replace_value_in_raw(raw: str, dotkey: str, value: Any) -> str:
    """Replace the value of a key in raw TOML text, preserving all other formatting.

    Only handles simple top-level or [section] keys, not deeply nested tables.
    Replaces only the first matching line under the correct section.
    """
    parts = dotkey.split(".", 1)
    if len(parts) == 2:
        section, leaf = parts
    else:
        section, leaf = None, parts[0]

    new_val_str = _format_value_for_toml(value, leaf)

    lines = raw.splitlines(keepends=True)
    in_section = section is None  # True when key is at root level
    out: list[str] = []
    replaced = False

    for line in lines:
        # Track section headers [foo]
        section_match = re.match(r"^\s*\[([^\]]+)\]\s*(?:#.*)?$", line)
        if section_match:
            in_section = section_match.group(1).strip() == section if section else False
            out.append(line)
            continue

        if in_section and not replaced:
            # Match assignment lines: key = value  (with optional inline comment)
            key_match = re.match(
                rf"^(\s*{re.escape(leaf)}\s*=\s*)(.*?)(\s*(?:#.*)?)$",
                line,
            )
            if key_match:
                indent_eq = key_match.group(1)
                tail = key_match.group(3)  # trailing comment
                out.append(f"{indent_eq}{new_val_str}{tail}\n")
                replaced = True
                continue

        out.append(line)

    if not replaced:
        raise KeyError(f"Key {dotkey!r} not found in config.toml — cannot set")

    return "".join(out)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get(key: str) -> Any:
    """Read a dot-path value from config.toml.

    Args:
        key: Dot-separated key path (e.g. 'preview.marker').

    Returns:
        The value at that path.

    Raises:
        FileNotFoundError: config.toml does not exist.
        KeyError: Key path not found.
    """
    raw = _load_raw()
    data = _parse(raw)
    node, leaf = _resolve_key(data, key)
    return node[leaf]


def set(key: str, value: Any) -> None:  # noqa: A001
    """Persist a value to config.toml, preserving the literal marker string.

    Args:
        key:   Dot-separated key path (e.g. 'preview.enabled').
        value: New value.

    Raises:
        FileNotFoundError: config.toml does not exist.
        KeyError: Key path not found (we don't create new keys — edit config.toml directly).
        TypeError: Value cannot be serialised to TOML.
    """
    raw = _load_raw()
    new_raw = _replace_value_in_raw(raw, key, value)
    _CONFIG_PATH.write_text(new_raw, encoding="utf-8")


# ---------------------------------------------------------------------------
# Marker safety (pure — unit-testable with no file I/O)
# ---------------------------------------------------------------------------

_MAX_MARKER_LEN = 16
_FORMULA_INJECTION_CHARS = frozenset("=+-@")
# Control characters: U+0000–U+001F and U+007F, plus newline/CR explicitly
_CONTROL_CHAR_RE = re.compile(r"[\x00-\x1f\x7f]")


def validate_marker(s: str) -> tuple[bool, str | None]:
    """Validate a blast-radius pending-marker string.

    Two-level outcome:
      - (False, reason)     REJECT — marker must not be used.
      - (True,  warning)    ACCEPT with warning — safe but notable.
      - (True,  None)       ACCEPT — clean.

    Rules:
      REJECT if s contains any control character (incl. newline, tab, CR).
      REJECT if len(s) > MAX_MARKER_LEN (16 chars).
      WARN   if s starts with '=', '+', '-', or '@' (CSV/Excel injection signal;
              neutralised by forcing text format on the cell, but worth flagging).
      ACCEPT otherwise.

    Args:
        s: The candidate marker string.

    Returns:
        (accepted: bool, message: str | None)
    """
    # Reject: control characters
    if _CONTROL_CHAR_RE.search(s):
        return (
            False,
            "Marker contains control characters (including newline/tab/CR) — rejected. "
            "These cannot be written into a cell and would corrupt the blast-radius indicator.",
        )

    # Reject: too long
    if len(s) > _MAX_MARKER_LEN:
        return (
            False,
            f"Marker is {len(s)} characters — exceeds the {_MAX_MARKER_LEN}-char limit. "
            "Use a shorter marker or a symbol-font glyph.",
        )

    # Warn: leading formula-injection character
    if s and s[0] in _FORMULA_INJECTION_CHARS:
        return (
            True,
            f"Marker starts with {s[0]!r} — a leading '= + - @' can trigger formula "
            "evaluation in some spreadsheet tools. The cell is force-set to text format "
            "so the marker is safe here, but consider a neutral marker (e.g. '~\\.\\\\\\\\o' "
            "or '⏳') to avoid ambiguity.",
        )

    return (True, None)
