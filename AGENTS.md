# Token Usage Inspector

A local web-based dashboard for analyzing LLM token usage from billing xlsx files.
Supports arbitrary billing spreadsheet layouts via heuristic column detection.

## Quick Start

```bash
# Install + run with uv (recommended)
uv sync
uv run token-inspector

# Or with pip
pip install flask openpyxl numpy
python -X utf8 app/server.py
```

Open http://127.0.0.1:5000 — drop a billing `.xlsx` or click "Load Sample Data".

## Project Layout

```
app/
  server.py          # Flask server (upload, sample, static serving)
  parser.py          # XLSX parser, aggregator, price estimator
  test_parser.py     # Unit tests (run: python -m unittest app.test_parser -v)
  test_frontend.py   # API integration tests
  static/
    index.html       # Dashboard HTML
    app.js           # Chart.js charts, table, model breakdown
    style.css        # Design system
data/
  .gitkeep           # Placeholder — billing files go here (not tracked)
```

## Architecture

- **Backend**: Python/Flask — parses xlsx via openpyxl, returns JSON
- **Frontend**: Vanilla HTML/CSS/JS + Chart.js v4 CDN — no build step
- **Parser**: Heuristic column detection (Chinese + English keywords), token type extraction via regex, least-squares price estimation via numpy
- **Tests**: `unittest` + integration test harness — no external test deps required

## Column Detection

The parser auto-detects 11 billing columns using keyword heuristics in both Chinese and English. Falls back gracefully if columns are missing. Supports these canonical names:

`date`, `resource_name`, `resource_id`, `billing_method`, `resource_type`, `model`, `usage_desc`, `site`, `transaction_type`, `service_fee`, `cost`

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/` | Dashboard HTML |
| `GET` | `/api/sample` | Parse bundled sample file |
| `GET` | `/api/data-files` | List `.xlsx` files in `data/` |
| `POST` | `/api/upload` | Upload and parse single xlsx (multipart `file`) |
| `POST` | `/api/upload-multiple` | Upload multiple xlsx files (multipart `files`) |
| `POST` | `/api/merge-data-files` | Load and merge files from `data/` (JSON `{"files": [...]}`) |

## Running Tests

**IMPORTANT — Never run `python app/server.py` directly in a foreground shell.**
It blocks and stalls the agent. Always use the test harness which manages the server lifecycle.

```bash
# Unit tests (no server needed)
python -X utf8 -m unittest app.test_parser -v

# Integration tests (starts/stops server as subprocess — safe, cross-platform)
# This is the recommended way to test the full app.
python -X utf8 app/test_frontend.py
```

### Safe server testing (for manual debugging)

```bash
# Windows — start in background, then kill when done
Start-Process -NoNewWindow python -ArgumentList "-X","utf8","app/server.py"
# ... test via curl or browser ...
Stop-Process -Name python -Force

# Linux/macOS — start in background, then kill when done
python -X utf8 app/server.py &
SERVER_PID=$!
# ... test via curl or browser ...
kill $SERVER_PID
```

## Encoding Note

On Windows, Python may default stdout to GBK. Always run with `-X utf8` or set `PYTHONUTF8=1`:
```bash
# Windows
$env:PYTHONUTF8 = "1"

# Linux/macOS
export PYTHONUTF8=1
```

## Data Directory

Place billing `.xlsx` files in `data/`. The app will list them and allow batch loading. Files in `data/` are git-ignored (not tracked). The sample file is no longer bundled — use the upload flow instead.
