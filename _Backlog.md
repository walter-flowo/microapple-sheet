# Excel MCP — Build Backlog

MVP shipped 2026-06-25 (Phases 0–3: engine routing/Rule-0, file ops, AppleScript live bridge, audit_hardcoded + auto-detect routing + config). 100 tests; validated in the MRH two-pane A2A pilot. Plan: `~/.claude/plans/lazy-tumbling-puzzle.md`.

## Deferred (post-MVP)
- **B001 — Blast-radius preview animation.** ⏳ pending-glyph markers (empty cells) + fill+border highlight (overwrite cells), drain, ⏳-emoji colour sampling (Pillow), content-aware + config (`preview.marker`/`overwrite_fill`/...). MVP shipped textual dry-run only.
- **B002 — XML surgery for dynamic-array/LAMBDA workbooks** (file-mode cache-preserving edits). Classifier detects + refuses today; the **live bridge is the only safe path** for these (LO recalc → #N/A cascade; openpyxl drops caches).
- **B003 — Change-aware preview server** (`watcher.py` + `preview.py` SSE + MCP resources) for side-by-side on closed files.
- **B004 — Graph co-authoring backend** (`graph.py`) — cloud / remote / multi-editor.
- **B005 — `style.py` FW house-style module** (palette / `put()` / lead-sheet) + full charts / pivots / conditional-formatting breadth.

## Operational notes
- After any code change, **`/mcp` reconnect (or restart)** to expose new tools — the running server holds the old code until respawn. Registration is **PYTHONPATH-hardened** (not editable-`.pth` dependent), so a broken editable install won't re-break "Failed to connect".
- `audit_hardcoded`: **strict-range** mode for gating result/total ranges; **whole-sheet** mode is data-row vs totals-row aware (won't false-flag input columns that carry a SUM total).
- Live bridge: never `Select`/`Activate`; address by sheet-qualified reference; one atomic Apple Event per write; read-back verify.

## Distil-later
- A2A two-pane file-protocol (registry / bus.jsonl / locks + incremental gating) → candidate Framework; relates to `FW002_Astro_Pane_Protocol`.
