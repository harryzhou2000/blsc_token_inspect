"""
Validate a billing xlsx file. Checks structure, columns, and data quality.
Usage: python -X utf8 app/validate.py <file.xlsx>
"""
import json
import sys
from pathlib import Path

# Ensure app/ is on sys.path so that 'from parser import parse_xlsx' works
_APP_DIR = Path(__file__).resolve().parent
if str(_APP_DIR) not in sys.path:
    sys.path.insert(0, str(_APP_DIR))

from parser import parse_xlsx


def validate(filepath: str | Path) -> int:
    """Run validation and return exit code (0=valid, 1=warnings, 2=error)."""
    filepath = Path(filepath)

    if not filepath.exists():
        print(f"Error: file not found: {filepath}")
        return 2

    try:
        result = parse_xlsx(str(filepath))
    except Exception as e:
        print(f"=== Validation Report ===")
        print(f"File: {filepath.name}")
        print(f"Error: {e}")
        print(f"Result: ERROR")
        return 2

    if "error" in result:
        print(f"=== Validation Report ===")
        print(f"File: {filepath.name}")
        print(f"Error: {result['error']}")
        print(f"Result: ERROR")
        return 2

    records = result.get("records", [])
    summary = result.get("summary", {})
    meta = result.get("meta", {})
    prices = result.get("prices", {})

    warnings: list[str] = []

    # --- Header ---
    print(f"=== Validation Report ===")
    print(f"File: {filepath.name}")
    data_rows = len(records)
    print(f"Rows: {data_rows + 1} (1 header + {data_rows} data)")

    # --- Column detection ---
    col_map = meta.get("column_map", {})
    all_canonical = {
        "date", "resource_name", "resource_id", "billing_method",
        "resource_type", "model", "usage_desc", "site",
        "transaction_type", "service_fee", "cost",
    }
    found_cols = set(col_map.keys())
    missing_cols = all_canonical - found_cols
    print(f"Columns detected: {len(found_cols)}/{len(all_canonical)}")
    print()
    print("Canonical columns:")
    print(f"  [FOUND] {', '.join(sorted(found_cols))}")
    print(f"  [MISS] {', '.join(sorted(missing_cols))}")
    print()

    # --- Summary ---
    print("Summary:")
    print(f"  Records: {summary.get('total_records', 0)}")
    print(f"  API Keys: {summary.get('api_key_count', 0)}")
    print(f"  Models: {summary.get('model_count', 0)}")
    print(f"  Dates: {summary.get('date_count', 0)}")
    print()

    # --- Token breakdown ---
    tin = summary.get("tokens_input", 0)
    tout = summary.get("tokens_output", 0)
    tcache = summary.get("tokens_cache_hit", 0)
    ttotal = summary.get("tokens_total", 0)

    print("Tokens:")
    print(f"  Input:   {tin:>12,}")
    print(f"  Output:  {tout:>12,}")
    print(f"  Cache:   {tcache:>12,}")
    print(f"  Total:   {ttotal:>12,}")
    print()

    # --- Cost ---
    cost = summary.get("cost", 0)
    print(f"Cost: ¥{cost:,.2f}")
    print()

    # --- Price estimates ---
    if prices:
        print("Price Estimates (¥/1M tokens):")
        for model in sorted(prices.keys()):
            p = prices[model]
            inp = f"¥{p['input']:.2f}" if p.get('input') is not None else "N/A"
            out = f"¥{p['output']:.2f}" if p.get('output') is not None else "N/A"
            print(f"  {model:20s}  in={inp:>8s}  out={out:>8s}")
        print()

    # --- Warnings ---
    # Rows with 0 cost but non-zero tokens
    zero_cost_token_rows = [
        r for r in records
        if r.get("tokens_total", 0) > 0 and r.get("cost", 0) == 0
    ]
    if zero_cost_token_rows:
        warnings.append(
            f"Rows with 0 cost but non-zero tokens: {len(zero_cost_token_rows)}"
        )

    # Rows with 0 tokens but non-zero cost
    zero_token_cost_rows = [
        r for r in records
        if r.get("tokens_total", 0) == 0 and r.get("cost", 0) > 0
    ]
    if zero_token_cost_rows:
        warnings.append(
            f"Rows with 0 tokens but non-zero cost: {len(zero_token_cost_rows)}"
        )

    # Duplicate rows (same resource_id, model, date, usage_desc, cost)
    seen: set[tuple] = set()
    duplicates = 0
    for r in records:
        key = (
            r.get("resource_id", ""),
            r.get("model", ""),
            r.get("date", ""),
            r.get("usage_desc", ""),
            r.get("cost", 0),
        )
        if key in seen:
            duplicates += 1
        else:
            seen.add(key)
    if duplicates:
        warnings.append(f"Duplicate rows (exact match): {duplicates}")

    print(f"Warnings: {len(warnings)}")
    for w in warnings:
        print(f"  - {w}")
    print()

    # --- Result ---
    if warnings:
        print("Result: VALID (with warnings)")
        return 1
    else:
        print("Result: VALID")
        return 0


def main():
    if len(sys.argv) < 2:
        print("Usage: python -X utf8 app/validate.py <file.xlsx>")
        sys.exit(2)

    exit_code = validate(sys.argv[1])
    sys.exit(exit_code)


if __name__ == "__main__":
    main()