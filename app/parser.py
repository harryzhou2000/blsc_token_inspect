"""
XLSX billing data parser — general-purpose column detection and normalization.
"""
import re
import json
from pathlib import Path
from typing import Any

import openpyxl


# Column detection heuristics: (canonical_name, [keyword_aliases])
COLUMN_PATTERNS = [
    ("date",              ["时间", "日期", "date", "time", "month"]),
    ("resource_name",     ["资源名称", "name", "resource_name"]),
    ("resource_id",       ["资源id", "resourceid", "resource_id", "apikey", "api_key", "key", "token"]),
    ("billing_method",    ["计费方式", "billing_method"]),
    ("resource_type",     ["资源类型", "type", "resource_type"]),
    ("model",             ["模型", "model"]),
    ("usage_desc",        ["配置描述", "用量描述", "用量", "usage", "消耗", "描述"]),
    ("site",              ["站点", "site", "region", "区域", "地域"]),
    ("transaction_type",  ["交易类型", "transaction_type"]),
    ("service_fee",       ["服务费", "service_fee"]),
    ("cost",              ["费用", "cost", "金额", "price"]),
]

# Token type patterns in usage descriptions (order matters: specific first)
TOKEN_TYPE_PATTERNS = [
    (re.compile(r"(缓存输入|缓存命中|cache[_\s]?hit|cached)[:\s：]*([\d,]+)\s*tokens?", re.IGNORECASE), "cache_hit"),
    (re.compile(r"(输入|input)[:\s：]*([\d,]+)\s*tokens?", re.IGNORECASE), "input"),
    (re.compile(r"(输出|output)[:\s：]*([\d,]+)\s*tokens?", re.IGNORECASE), "output"),
]


def _parse_float(val) -> float:
    """Parse a value to float, handling strings with commas."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def estimate_model_prices(records: list[dict]) -> dict[str, dict[str, float | None]]:
    """Estimate per-token prices for each model using least-squares linear regression.

    For each model, solves the overdetermined system:
        input_i * p_in + output_i * p_out + cache_i * p_cache = cost_i
    using numpy least-squares (minimizing ||Ax - b||^2).
    Requires at least 2 rows per model. Falls back to simple averaging for single-type rows
    if the regression produces implausible (negative or zero) results.
    """
    import numpy as np

    models = set(r["model"] for r in records if r["model"])
    prices: dict[str, dict[str, float | None]] = {}

    for model in models:
        model_recs = [r for r in records if r["model"] == model]
        prices[model] = {"input": None, "output": None, "cache_hit": None}

        # Build matrix A (n×3) and vector b (n×1)
        rows_A = []
        rows_b = []
        for r in model_recs:
            if r["tokens_total"] == 0 and r["cost"] == 0:
                continue
            rows_A.append([float(r["tokens_input"]), float(r["tokens_output"]), float(r["tokens_cache_hit"])])
            rows_b.append(float(r["cost"]))

        if len(rows_A) < 2:
            continue

        A = np.array(rows_A, dtype=np.float64)
        b = np.array(rows_b, dtype=np.float64)

        try:
            x, residuals, rank, sv = np.linalg.lstsq(A, b, rcond=None)
        except np.linalg.LinAlgError:
            continue

        # Only accept positive price estimates
        if x[0] > 0:
            prices[model]["input"] = float(x[0])
        if x[1] > 0:
            prices[model]["output"] = float(x[1])
        if x[2] > 0:
            prices[model]["cache_hit"] = float(x[2])

        # Fallback: if a price type couldn't be estimated via regression but
        # there are single-type rows, use the simple average.
        for token_type, col_idx in [("input", 0), ("output", 1), ("cache_hit", 2)]:
            if prices[model][token_type] is not None:
                continue
            token_key = f"tokens_{token_type}"
            single_recs = [
                r for r in model_recs
                if r[token_key] > 0
                and all(r[f"tokens_{other}"] == 0 for other in ["input", "output", "cache_hit"] if other != token_type)
            ]
            if single_recs:
                total_tokens = sum(r[token_key] for r in single_recs)
                total_cost = sum(r["cost"] for r in single_recs)
                if total_tokens > 0:
                    prices[model][token_type] = total_cost / total_tokens

    return prices


def detect_columns(headers: list[str]) -> dict[str, int]:
    """Map canonical column names to 0-based indices using heuristics."""
    mapping: dict[str, int] = {}
    normalized_headers = [h.strip().lower() if h else "" for h in headers]

    for canonical, aliases in COLUMN_PATTERNS:
        for idx, h in enumerate(normalized_headers):
            if canonical in mapping:
                break
            for alias in aliases:
                if alias in h:
                    mapping[canonical] = idx
                    break

    return mapping


def parse_usage_desc(text: str) -> list[dict[str, Any]]:
    """Parse a usage description cell into token type + count entries.
    
    Returns at most one entry — the first matching pattern (ordered by specificity).
    """
    if not text or not isinstance(text, str):
        return []

    for pattern, token_type in TOKEN_TYPE_PATTERNS:
        m = pattern.search(text)
        if m:
            count_str = m.group(2).replace(",", "")
            try:
                count = int(count_str)
            except ValueError:
                continue
            return [{"type": token_type, "tokens": count}]

    return []


def _ensure_xlsx(filepath: Path) -> Path:
    """If the file is an xlsx with a non-xlsx extension, copy to a temp .xlsx file."""
    if filepath.suffix.lower() in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
        return filepath
    # Check magic bytes — xlsx files are ZIP archives starting with PK
    with open(filepath, 'rb') as f:
        header = f.read(4)
    if header[:2] != b'PK':
        raise ValueError(f"File is not a valid xlsx: {filepath.name}")
    # Copy to temp .xlsx
    import tempfile, shutil
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    shutil.copyfile(filepath, tmp.name)
    return Path(tmp.name)


def parse_xlsx(filepath: str | Path) -> dict[str, Any]:
    """Parse an xlsx billing file and return normalized JSON-serializable data."""
    filepath = Path(filepath)
    actual_path = _ensure_xlsx(filepath)
    try:
        wb = openpyxl.load_workbook(actual_path)
    except Exception:
        if actual_path != filepath:
            actual_path.unlink(missing_ok=True)
        raise
    try:
        ws = wb.active

        raw_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row))

        def cell_val(cell):
            return cell.value

        rows = [[cell_val(c) for c in row] for row in raw_rows]

        if not rows:
            wb.close()
            return {"error": "Empty spreadsheet", "records": []}

        # Detect headers from first row
        headers = [str(c) if c is not None else "" for c in rows[0]]
        col_map = detect_columns(headers)

        records = []
        for row in rows[1:]:
            if all(c is None for c in row):
                continue
            values = [str(c) if c is not None else "" for c in row]

            usage_desc = values[col_map["usage_desc"]] if "usage_desc" in col_map else ""
            token_entries = parse_usage_desc(usage_desc)

            record = {
                "date": values[col_map["date"]] if "date" in col_map else "",
                "resource_name": values[col_map["resource_name"]] if "resource_name" in col_map else "",
                "resource_type": values[col_map["resource_type"]] if "resource_type" in col_map else "",
                "resource_id": values[col_map["resource_id"]] if "resource_id" in col_map else "",
                "billing_method": values[col_map["billing_method"]] if "billing_method" in col_map else "",
                "model": values[col_map["model"]] if "model" in col_map else "",
                "usage_desc": usage_desc,
                "site": values[col_map["site"]] if "site" in col_map else "",
                "transaction_type": values[col_map["transaction_type"]] if "transaction_type" in col_map else "",
                "service_fee": _parse_float(values[col_map["service_fee"]]) if "service_fee" in col_map else 0.0,
                "cost": _parse_float(values[col_map["cost"]]) if "cost" in col_map else 0.0,
                "tokens": token_entries,
            }
            record["tokens_input"] = sum(e["tokens"] for e in token_entries if e["type"] == "input")
            record["tokens_output"] = sum(e["tokens"] for e in token_entries if e["type"] == "output")
            record["tokens_cache_hit"] = sum(e["tokens"] for e in token_entries if e["type"] == "cache_hit")
            record["tokens_total"] = record["tokens_input"] + record["tokens_output"] + record["tokens_cache_hit"]
            records.append(record)

        wb.close()

        # Build aggregations via shared function
        return aggregate_records(records, meta={
            "filename": filepath.name,
            "column_map": {k: headers[v] for k, v in col_map.items()},
        })
    finally:
        if actual_path != filepath:
            actual_path.unlink(missing_ok=True)


def aggregate_records(records: list[dict], meta: dict | None = None) -> dict[str, Any]:
    """Aggregate a list of record dicts into summary, by_key, by_resource_name, by_model, timeline.

    This is the same logic used inside parse_xlsx, extracted for reuse when merging
    records from multiple files.

    Args:
        records: List of record dicts (as produced by parse_xlsx per-row logic).
        meta: Optional extra metadata to include in the top-level result.

    Returns:
        A dict with keys: meta, summary, records, by_key, by_resource_name, by_model,
        timeline, prices.
    """
    api_keys = sorted(set(r["resource_id"] for r in records if r["resource_id"]))
    models = sorted(set(r["model"] for r in records if r["model"]))
    dates = sorted(set(r["date"] for r in records if r["date"]))
    resource_names = sorted(set(r["resource_name"] for r in records if r["resource_name"]))

    # By API key
    by_key = {}
    for key in api_keys:
        kr = [r for r in records if r["resource_id"] == key]
        by_key[key] = {
            "resource_type": kr[0]["resource_type"] if kr else "",
            "record_count": len(kr),
            "tokens_input": sum(r["tokens_input"] for r in kr),
            "tokens_output": sum(r["tokens_output"] for r in kr),
            "tokens_cache_hit": sum(r["tokens_cache_hit"] for r in kr),
            "tokens_total": sum(r["tokens_total"] for r in kr),
            "cost": sum(r["cost"] for r in kr),
            "models": sorted(set(r["model"] for r in kr)),
        }

    # By resource_name (user-friendly labels)
    by_resource_name = {}
    for name in resource_names:
        nr = [r for r in records if r["resource_name"] == name]
        by_resource_name[name] = {
            "record_count": len(nr),
            "tokens_input": sum(r["tokens_input"] for r in nr),
            "tokens_output": sum(r["tokens_output"] for r in nr),
            "tokens_cache_hit": sum(r["tokens_cache_hit"] for r in nr),
            "tokens_total": sum(r["tokens_total"] for r in nr),
            "cost": sum(r["cost"] for r in nr),
            "models": sorted(set(r["model"] for r in nr)),
            "api_keys": sorted(set(r["resource_id"] for r in nr)),
        }

    # By model
    by_model = {}
    for model in models:
        mr = [r for r in records if r["model"] == model]
        by_model[model] = {
            "record_count": len(mr),
            "tokens_input": sum(r["tokens_input"] for r in mr),
            "tokens_output": sum(r["tokens_output"] for r in mr),
            "tokens_cache_hit": sum(r["tokens_cache_hit"] for r in mr),
            "tokens_total": sum(r["tokens_total"] for r in mr),
            "cost": sum(r["cost"] for r in mr),
        }

    # Timeline by date
    timeline = {}
    for date in dates:
        dr = [r for r in records if r["date"] == date]
        entry = {
            "tokens_input": sum(r["tokens_input"] for r in dr),
            "tokens_output": sum(r["tokens_output"] for r in dr),
            "tokens_cache_hit": sum(r["tokens_cache_hit"] for r in dr),
            "tokens_total": sum(r["tokens_total"] for r in dr),
            "cost": sum(r["cost"] for r in dr),
            "by_key": {},
        }
        for key in api_keys:
            kdr = [r for r in dr if r["resource_id"] == key]
            if kdr:
                entry["by_key"][key] = {
                    "tokens_input": sum(r["tokens_input"] for r in kdr),
                    "tokens_output": sum(r["tokens_output"] for r in kdr),
                    "tokens_cache_hit": sum(r["tokens_cache_hit"] for r in kdr),
                    "tokens_total": sum(r["tokens_total"] for r in kdr),
                    "cost": sum(r["cost"] for r in kdr),
                }
        timeline[date] = entry

    summary = {
        "total_records": len(records),
        "api_key_count": len(api_keys),
        "model_count": len(models),
        "date_count": len(dates),
        "tokens_input": sum(r["tokens_input"] for r in records),
        "tokens_output": sum(r["tokens_output"] for r in records),
        "tokens_cache_hit": sum(r["tokens_cache_hit"] for r in records),
        "tokens_total": sum(r["tokens_total"] for r in records),
        "cost": sum(r["cost"] for r in records),
    }

    result = {
        "meta": meta or {},
        "summary": summary,
        "records": records,
        "by_key": by_key,
        "by_resource_name": by_resource_name,
        "by_model": by_model,
        "timeline": timeline,
        "prices": estimate_model_prices(records),
    }
    # Always include these in meta for downstream consumers
    if meta:
        result["meta"].update({
            "api_keys": api_keys,
            "resource_names": resource_names,
            "models": models,
            "dates": dates,
        })
    else:
        result["meta"] = {
            "api_keys": api_keys,
            "resource_names": resource_names,
            "models": models,
            "dates": dates,
        }

    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python parser.py <file.xlsx>")
        sys.exit(1)

    result = parse_xlsx(sys.argv[1])
    print(json.dumps(result, ensure_ascii=False, indent=2))
