# Contributing to MicroApple Sheet

## Dev setup

```bash
git clone https://github.com/walter-flowo/microapple-sheet
cd microapple-sheet
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Run tests

```bash
python -m pytest tests/ -q
```

Tests are split by platform requirement:

- **All platforms** — engine classification, openpyxl file ops, bridge unit tests
  (osascript mocked), audit/formula validation (102 tests at MVP)
- **LibreOffice required** — Rule-0 regression tests (`test_rule0_regression.py`)
  and the computed-workbook write test; skipped automatically when `soffice` is
  not on PATH and not at a known platform path
- **macOS + Microsoft Excel required** — live-bridge integration tests; use
  `@pytest.mark.skipif(sys.platform != "darwin", ...)` for any new tests in
  this category

On Ubuntu CI, install LibreOffice before running tests:

```bash
sudo apt-get update && sudo apt-get install -y libreoffice-calc
```

## Pull requests

- Keep `excel_*` tool names unchanged — these are part of the public MCP API
- Add tests for new tools or engine paths
- Update `CHANGELOG.md` under `[Unreleased]` for any user-visible change
- Run `pytest -q` locally before pushing; CI must be green on both
  `ubuntu-latest` and `macos-latest`
