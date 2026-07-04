# Token Usage Inspector

A local web-based dashboard for analyzing LLM token usage from billing `.xlsx` files.
Supports arbitrary billing spreadsheet layouts via heuristic column detection.

![version](https://img.shields.io/badge/version-0.1.0-blue)

## Quick Start

```bash
# Install + run with uv (recommended)
uv sync
uv run token-inspector

# Or with pip
pip install flask openpyxl numpy
python -X utf8 app/server.py
```

Open **http://127.0.0.1:5000** — drop a billing `.xlsx` or load files from the `data/` directory.

## Features

- **Upload & Parse**: Drag-and-drop billing `.xlsx` files, or batch-load from `data/`
- **Auto-detect Columns**: Heuristic detection of 11 billing columns in both Chinese and English
- **Token Extraction**: Parses input, output, and cached-input token counts from usage descriptions
- **Price Estimation**: Least-squares linear regression (numpy) to estimate per-token prices per model
- **Interactive Charts**: Bar, stacked bar, doughnut, and timeline charts via Chart.js — with toggleable decompositions (by model, by token type)
- **Filterable Table**: Sort and filter all billing records; view truncated API keys with full key on hover

## Dashboard Sections

| Section | Description |
|---------|-------------|
| Summary Cards | Total tokens, cost, API keys, models, records |
| Cost by Resource | Total spend per resource, toggle to decompose by model or token type |
| Tokens by Resource | Stacked input/output/cached token bars per resource |
| Timeline | Cost and token trends over months (line charts) |
| Cost by Model | Doughnut chart of spend share per model |
| Tokens by Model | Stacked token bars per model |
| Model Breakdown by Resource | Drill into each resource's per-model token mix and estimated cost |
| Prices Table | Estimated ¥/1M tokens per model, derived via least-squares regression |
| Record Details | Sortable, filterable table of all billing rows |

## Data Directory

Place billing `.xlsx` files (or `.csv` files that are actually `.xlsx` format) in `data/`.
The app scans this directory and shows available files for batch loading.

Files in `data/` are git-ignored — not tracked in version control.

## Running Tests

```bash
# Unit tests (no server needed)
python -X utf8 -m unittest app.test_parser -v

# Integration tests (starts/stops server as subprocess)
python -X utf8 app/test_frontend.py
```

## Validate a Billing File

```bash
python -X utf8 app/validate.py data/your_file.xlsx
```

Prints a validation report with column detection results, token breakdowns, cost totals,
price estimates, and warnings.

## Encoding Note

On Windows, Python may default stdout to GBK. Always run with `-X utf8` or set `PYTHONUTF8=1`:

```bash
# Windows PowerShell
$env:PYTHONUTF8 = "1"

# Linux/macOS
export PYTHONUTF8=1
```

## Tech Stack

- **Backend**: Python 3.10+ / Flask
- **Frontend**: Vanilla HTML/CSS/JS + Chart.js v4 (CDN)
- **Parsing**: openpyxl, numpy
- **No build step**, no JavaScript framework

## License

MIT