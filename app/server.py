"""
Flask server for token usage inspector.
"""
import tempfile
import os
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory

from parser import parse_xlsx, aggregate_records

app = Flask(__name__, static_folder="static", static_url_path="")

SAMPLE_PATH = Path(__file__).resolve().parent.parent / "data" / "资源ID账单_2026-07-01_2026-07-04_1783146280975_按资源汇总.csv"
DATA_DIR = Path(__file__).resolve().parent.parent / "data"


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


def main():
    """Entry point for `uv run token-inspector`."""
    import glob as _glob
    static_files = _glob.glob(str(Path(__file__).parent / "static" / "**" / "*"), recursive=True)
    app.run(host="0.0.0.0", port=5000, debug=True, extra_files=static_files)


if __name__ == "__main__":
    main()
