# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-06-25

### Added

- **Rule-0-safe engine routing** — `classify_workbook()` inspects every workbook
  before any write and picks the correct engine:
  - `openpyxl` for new / formula-free workbooks (safe load-modify-save)
  - `LibreOffice headless` for computed workbooks (two-step openpyxl edit + LO
    recalc; never drops cached `<v>` formula values)
  - Refuses automated edits on dynamic-array / LAMBDA / XLOOKUP workbooks (Rule
    0b: LO recalc cascades to `#N/A` on these; live bridge is the only safe path)
  - `excel_live` when the file is open in Excel (routes all writes through the
    AppleScript bridge, never dual-writing)

- **File operations** — `excel_create`, `excel_open`, `excel_info`, `excel_recalc`,
  `excel_convert` (PDF / CSV / XLSX via LibreOffice)

- **Read operations** — `excel_read` (values / formulas / both modes),
  `excel_read_table` (header-keyed records), `excel_list_sheets`, `excel_list_names`

- **Write operations** — `excel_write`, `excel_set_cell`, `excel_write_formula`,
  `excel_write_linked_formula` (Cell.coordinate-safe live references),
  `excel_format`, `excel_define_name`

- **AppleScript live bridge** (macOS + Microsoft Excel only) — `excel_live_read`,
  `excel_live_set`, `excel_live_formula`, `excel_is_open`, `excel_check_automation`;
  single atomic Apple Event per write; injection-safe argv transport; edit-mode
  retry loop with TCC / not-running / not-open error taxonomy

- **`excel_audit_hardcoded`** — column-peer heuristic lints workbooks for numeric
  literals that look calculated; totals-row aware; strict-range mode for gating
  specific output areas

- **Auto-detect routing** in `excel_write`, `excel_set_cell`, `excel_write_formula`
  — probes `is_open()` and routes live / file automatically

- **Config tools** — `excel_config_get` / `excel_config_set` for runtime
  configuration of preview markers, drain thresholds, and colour/font keys

- **Cross-platform LibreOffice discovery** — resolution order: env var
  `MICROAPPLE_SHEET_SOFFICE` → `shutil.which` → well-known platform paths;
  file-engine tools work on Linux/macOS; live bridge is macOS-only with clean
  structured errors on other platforms

- **Backup system** — timestamped `.backup/` snapshot adjacent to every workbook
  before any write; `DEFAULT_KEEP=10` pruning

### Validated

- Rule-0 proof: 320 → naive openpyxl = 20 (97% loss) → safe two-step = 320
  (100% restored); 100 tests green on macOS
