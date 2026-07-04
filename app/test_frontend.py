"""
Non-blocking frontend testing script for the BLSC token inspect server.

Starts the Flask server, runs API tests, saves a JSON report, and cleans up.
Usage: python -X utf8 app/test_frontend.py
"""

import atexit
import json
import os
import signal
import subprocess
import sys
import time
import uuid
import tempfile
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent.parent
SERVER_SCRIPT = PROJECT_ROOT / "app" / "server.py"
SAMPLE_XLSX = PROJECT_ROOT / "data" / "资源ID账单_2026-07-01_2026-07-04_1783146280975_按资源汇总.csv"
REPORT_PATH = Path(tempfile.gettempdir()) / "opencode" / "frontend_test_report.json"

SERVER_URL = "http://127.0.0.1:5000"
POLL_URL = f"{SERVER_URL}/api/sample"
POLL_TIMEOUT = 30  # seconds to wait for server readiness
POLL_INTERVAL = 1.0

# ---------------------------------------------------------------------------
# Globals for subprocess management
# ---------------------------------------------------------------------------
_server_proc: subprocess.Popen | None = None


# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------
def _kill_server() -> None:
    """Terminate the server subprocess; called by atexit and on error."""
    global _server_proc
    proc = _server_proc
    if proc is None or proc.poll() is not None:
        return
    print("[cleanup] Terminating server (pid=%d) ..." % proc.pid, file=sys.stderr)
    try:
        proc.terminate()           # SIGTERM / CTRL_BREAK_EVENT on Windows
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            print("[cleanup] Server did not exit gracefully, force-killing ...",
                  file=sys.stderr)
            proc.kill()            # SIGKILL / TerminateProcess on Windows
            proc.wait(timeout=5)
    except Exception as exc:
        print("[cleanup] Error during server shutdown: %s" % exc, file=sys.stderr)
    _server_proc = None


atexit.register(_kill_server)


# ---------------------------------------------------------------------------
# HTTP helpers (stdlib only — no `requests`)
# ---------------------------------------------------------------------------
def _http_get(url: str, timeout: float = 10) -> tuple[int, bytes]:
    """Perform a GET request and return (status_code, body)."""
    req = Request(url, method="GET")
    try:
        resp = urlopen(req, timeout=timeout)
        return resp.status, resp.read()
    except HTTPError as e:
        return e.code, e.read()
    except URLError:
        raise


def _http_post_file(
    url: str, filepath: Path, timeout: float = 30
) -> tuple[int, bytes]:
    """POST a multipart/form-data request with a single file field named 'file'."""
    boundary = uuid.uuid4().hex
    filename = filepath.name
    file_bytes = filepath.read_bytes()

    # Build multipart body
    header = (
        b"--%b\r\n"
        b'Content-Disposition: form-data; name="file"; filename="%b"\r\n'
        b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    ) % (boundary.encode(), filename.encode())

    footer = b"\r\n--%b--\r\n" % boundary.encode()

    body = header + file_bytes + footer

    req = Request(url, data=body, method="POST")
    req.add_header(
        "Content-Type",
        "multipart/form-data; boundary=%s" % boundary,
    )

    try:
        resp = urlopen(req, timeout=timeout)
        return resp.status, resp.read()
    except HTTPError as e:
        return e.code, e.read()
    except URLError:
        raise


def _json_or_none(body: bytes) -> dict | None:
    """Try to decode bytes as JSON; return None on failure."""
    try:
        return json.loads(body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None


# ---------------------------------------------------------------------------
# Server lifecycle
# ---------------------------------------------------------------------------
def start_server() -> subprocess.Popen:
    """Start the Flask server as a background subprocess."""
    global _server_proc
    print("[server] Starting server: python %s" % SERVER_SCRIPT, file=sys.stderr)
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    _server_proc = subprocess.Popen(
        [sys.executable, "-X", "utf8", str(SERVER_SCRIPT)],
        cwd=str(PROJECT_ROOT),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        env=env,
        # On Windows, CREATE_NEW_PROCESS_GROUP enables Ctrl+Break signalling.
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0,
    )
    return _server_proc


def wait_for_server(poll_url: str = POLL_URL, timeout: float = POLL_TIMEOUT) -> bool:
    """Poll the server until it responds or timeout is reached."""
    deadline = time.monotonic() + timeout
    last_error = ""
    while time.monotonic() < deadline:
        try:
            code, body = _http_get(poll_url, timeout=POLL_INTERVAL)
            if code == 200:
                print("[server] Server is ready.", file=sys.stderr)
                return True
            last_error = "HTTP %d" % code
        except Exception as exc:
            last_error = str(exc)
        time.sleep(POLL_INTERVAL)
    print("[server] Server did not become ready within %.0fs (last: %s)" % (
        timeout, last_error), file=sys.stderr)
    return False


def stop_server() -> None:
    """Explicitly stop the server (called directly; atexit also handles it)."""
    _kill_server()


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
def test_index() -> dict:
    """GET / should return HTML with status 200."""
    url = f"{SERVER_URL}/"
    try:
        code, body = _http_get(url)
        is_html = any(
            tag in body.decode("utf-8", errors="replace").lower()
            for tag in ("<!doctype html", "<html", "<head")
        )
        passed = code == 200 and is_html
        return {
            "test": "GET /",
            "passed": passed,
            "status": code,
            "detail": "returns HTML" if is_html else "response is not HTML",
            "body_preview": body[:200].decode("utf-8", errors="replace"),
        }
    except Exception as e:
        return {"test": "GET /", "passed": False, "status": None, "detail": str(e)}


SAMPLE_TOP_KEYS = {"meta", "summary", "records", "by_key", "by_resource_name", "by_model", "timeline"}


def test_api_sample() -> dict:
    """GET /api/sample should return valid JSON with all expected keys."""
    try:
        code, body = _http_get(POLL_URL)
        data = _json_or_none(body)
        if data is None:
            return {
                "test": "GET /api/sample",
                "passed": False,
                "status": code,
                "detail": "response is not valid JSON",
            }
        returned_keys = set(data.keys())
        missing = SAMPLE_TOP_KEYS - returned_keys
        passed = code == 200 and not missing
        detail = (
            "all expected keys present"
            if passed
            else "missing keys: %s" % ", ".join(sorted(missing))
        )
        return {
            "test": "GET /api/sample",
            "passed": passed,
            "status": code,
            "detail": detail,
            "returned_keys": sorted(returned_keys),
        }
    except Exception as e:
        return {"test": "GET /api/sample", "passed": False, "status": None, "detail": str(e)}


def test_api_upload() -> dict:
    """POST /api/upload with the sample xlsx file should return valid JSON."""
    if not SAMPLE_XLSX.exists():
        return {
            "test": "POST /api/upload",
            "passed": False,
            "status": None,
            "detail": "sample file not found: %s" % SAMPLE_XLSX,
        }
    try:
        code, body = _http_post_file(f"{SERVER_URL}/api/upload", SAMPLE_XLSX)
        data = _json_or_none(body)
        if data is None:
            return {
                "test": "POST /api/upload",
                "passed": False,
                "status": code,
                "detail": "response is not valid JSON",
            }
        returned_keys = set(data.keys())
        missing = SAMPLE_TOP_KEYS - returned_keys
        passed = code == 200 and not missing
        detail = (
            "all expected keys present"
            if passed
            else "missing keys: %s" % ", ".join(sorted(missing))
        )
        return {
            "test": "POST /api/upload",
            "passed": passed,
            "status": code,
            "detail": detail,
            "returned_keys": sorted(returned_keys),
        }
    except Exception as e:
        return {"test": "POST /api/upload", "passed": False, "status": None, "detail": str(e)}


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------
def run_tests() -> list[dict]:
    """Run all API tests and return a list of result dicts."""
    tests = [test_index(), test_api_sample(), test_api_upload()]
    return tests


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    print("=" * 60, file=sys.stderr)
    print("  BLSC Token Inspect — Frontend Test Suite", file=sys.stderr)
    print("=" * 60, file=sys.stderr)

    # 1. Start server
    proc = start_server()
    if proc.poll() is not None:
        stderr_output = proc.stderr.read().decode("utf-8", errors="replace") if proc.stderr else ""
        print("[FATAL] Server process exited immediately:\n%s" % stderr_output,
              file=sys.stderr)
        return 1

    # 2. Wait for readiness
    ready = wait_for_server()
    if not ready:
        stop_server()
        return 1

    # 3. Run tests
    print("[tests] Running API tests ...", file=sys.stderr)
    results = run_tests()

    # 4. Stop server
    stop_server()

    # 5. Build report
    passed_count = sum(1 for r in results if r["passed"])
    total_count = len(results)

    report = {
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "summary": {
            "total": total_count,
            "passed": passed_count,
            "failed": total_count - passed_count,
        },
        "results": results,
    }

    # 6. Save report
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text(
        json.dumps(report, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print("[report] Saved to %s" % REPORT_PATH, file=sys.stderr)

    # 7. Print summary to stdout
    print()
    print("=" * 60)
    print("  Frontend Test Results")
    print("=" * 60)
    print("  Timestamp : %s" % report["timestamp"])
    print("  Total     : %d" % total_count)
    print("  Passed    : %d" % passed_count)
    print("  Failed    : %d" % (total_count - passed_count))
    print()
    for r in results:
        status = "PASS" if r["passed"] else "FAIL"
        print("  [%s] %s" % (status, r["test"]))
        print("        %s" % r["detail"])
    print()
    print("  Report file: %s" % REPORT_PATH)
    print("=" * 60)

    return 0 if passed_count == total_count else 1


if __name__ == "__main__":
    sys.exit(main())