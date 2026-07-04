"""
Flask server for token usage inspector.
"""
import datetime
import json
import tempfile
import os
import urllib.request
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, Response

from parser import parse_xlsx, aggregate_records

app = Flask(__name__, static_folder="static", static_url_path="")

STATIC_DIR = Path(__file__).resolve().parent / "static"
SAMPLE_PATH = Path(__file__).resolve().parent.parent / "data" / "资源ID账单_2026-07-01_2026-07-04_1783146280975_按资源汇总.csv"
DATA_DIR = Path(__file__).resolve().parent.parent / "data"
CHART_CDN_URL = "https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/upload", methods=["POST"])
def api_upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    suffix = Path(f.filename).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp_path = tmp.name
        f.save(tmp.name)

    try:
        result = parse_xlsx(tmp_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    finally:
        os.unlink(tmp_path)

    return jsonify(result)


@app.route("/api/sample")
def api_sample():
    if not SAMPLE_PATH.exists():
        return jsonify({"error": "Sample file not found"}), 404
    try:
        result = parse_xlsx(SAMPLE_PATH)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify(result)


@app.route("/api/data-files")
def api_data_files():
    """Scan data/ directory and return list of .xlsx files."""
    if not DATA_DIR.exists():
        return jsonify({"files": []})
    files = []
    for f in sorted(DATA_DIR.iterdir()):
        if f.suffix.lower() in (".xlsx", ".xls", ".csv"):
            files.append({"name": f.name, "size": f.stat().st_size})
    return jsonify({"files": files})


@app.route("/api/merge-data-files", methods=["POST"])
def api_merge_data_files():
    """Accept a JSON list of filenames from data/, parse each, merge, re-aggregate."""
    body = request.get_json(silent=True) or {}
    filenames = body.get("files", [])
    if not filenames or not isinstance(filenames, list):
        return jsonify({"error": "Provide a JSON body with a 'files' array"}), 400

    all_records = []
    errors = []

    for name in filenames:
        fpath = DATA_DIR / name
        if not fpath.exists() or fpath.suffix.lower() not in (".xlsx", ".xls", ".csv"):
            errors.append({"file": name, "error": "File not found or unsupported type"})
            continue
        try:
            result = parse_xlsx(fpath)
            if "records" in result:
                all_records.extend(result["records"])
        except Exception as e:
            errors.append({"file": name, "error": str(e)})

    if not all_records:
        return jsonify({"error": "No records could be parsed from any file", "errors": errors}), 400

    merged = aggregate_records(all_records, meta={
        "filenames": filenames,
        "api_keys": [],
        "resource_names": [],
        "models": [],
        "dates": [],
    })

    if errors:
        merged["meta"]["parse_errors"] = errors

    return jsonify(merged)


@app.route("/api/upload-multiple", methods=["POST"])
def api_upload_multiple():
    """Accept multiple billing files, parse each, merge records, and re-aggregate."""
    if "files" not in request.files:
        return jsonify({"error": "No files provided"}), 400

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files provided"}), 400

    all_records = []
    filenames = []
    errors = []

    for f in files:
        if f.filename == "":
            continue
        suffix = Path(f.filename).suffix or ".xlsx"
        filenames.append(f.filename)
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp_path = tmp.name
            f.save(tmp.name)
        try:
            result = parse_xlsx(tmp_path)
            if "records" in result:
                all_records.extend(result["records"])
        except Exception as e:
            errors.append({"file": f.filename, "error": str(e)})
        finally:
            os.unlink(tmp_path)

    if not all_records:
        return jsonify({"error": "No records could be parsed from any file", "errors": errors}), 400

    merged = aggregate_records(all_records, meta={
        "filenames": filenames,
        "api_keys": [],
        "resource_names": [],
        "models": [],
        "dates": [],
    })

    if errors:
        merged["meta"]["parse_errors"] = errors

    return jsonify(merged)


@app.route("/api/export-html", methods=["POST"])
def api_export_html():
    """Accept dashboard data JSON and return a self-contained HTML report for download."""
    body = request.get_json(silent=True)
    if not body:
        return jsonify({"error": "Invalid or missing JSON body"}), 400

    # Read static files
    index_html = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    style_css = (STATIC_DIR / "style.css").read_text(encoding="utf-8")
    app_js = (STATIC_DIR / "app.js").read_text(encoding="utf-8")

    # Chart.js — download and cache locally
    chart_js_path = STATIC_DIR / "chart.umd.min.js"
    if not chart_js_path.exists():
        try:
            urllib.request.urlretrieve(CHART_CDN_URL, chart_js_path)
        except Exception as e:
            return jsonify({
                "error": "Failed to download Chart.js. Please check your internet connection and try again. "
                         f"({e})"
            }), 500
    chart_js = chart_js_path.read_text(encoding="utf-8")

    # Serialize data as JSON — escape </script> to be safe inside the script tag
    data_json = json.dumps(body, ensure_ascii=False, default=str).replace("</script>", "<\\/script>")

    # Build the renderReportFromData function to inject at the end of the IIFE
    expose_code = """
window.renderReportFromData = function(data) {
    state.data = data;
    state.filteredRecords = [...data.records];
    renderAll();
};
"""

    # Bootstrap snippet that runs when the page loads standalone
    bootstrap = """
(function() {
    var dataEl = document.getElementById('report-data');
    if (!dataEl) return;
    var data;
    try { data = JSON.parse(dataEl.textContent); } catch (e) { return; }
    document.addEventListener('DOMContentLoaded', function() {
        if (typeof window.renderReportFromData === 'function') {
            window.renderReportFromData(data);
        }
    });
})();
"""

    # Insert renderReportFromData inside the IIFE (before the closing })
    inline_app_js = app_js.rstrip()
    if inline_app_js.endswith("})();"):
        inline_app_js = inline_app_js[:-5] + expose_code + "\n})();"
    else:
        inline_app_js = inline_app_js + "\n" + expose_code
    inline_app_js += "\n" + bootstrap

    # Build final HTML: replace external resources with inline content
    html = index_html

    # Replace external CSS link with inline <style>
    html = html.replace(
        '<link rel="stylesheet" href="style.css">',
        "<style>" + style_css + "</style>"
    )

    # Replace Chart.js CDN script with inline script
    html = html.replace(
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>',
        "<script>" + chart_js + "</script>"
    )

    # Replace app.js script with data block + inline script
    html = html.replace(
        '<script src="app.js"></script>',
        '<script id="report-data" type="application/json">' + data_json +
        '</script>\n    <script>' + inline_app_js + '</script>'
    )

    # Generate filename with today's date
    today = datetime.date.today().strftime("%Y-%m-%d")
    filename = "token-usage-report-{}.html".format(today)

    return Response(
        html,
        mimetype="text/html",
        headers={"Content-Disposition": 'attachment; filename="{}"'.format(filename)}
    )


def main():
    """Entry point for `uv run token-inspector`."""
    import glob as _glob
    static_files = _glob.glob(str(Path(__file__).parent / "static" / "**" / "*"), recursive=True)
    app.run(host="0.0.0.0", port=5000, debug=True, extra_files=static_files)


if __name__ == "__main__":
    main()
