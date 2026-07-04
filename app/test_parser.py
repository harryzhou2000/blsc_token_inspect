"""
Tests for the billing XLSX parser using Python's built-in unittest.
"""
import unittest
import sys
from pathlib import Path

# Ensure app/ is on sys.path so that 'from parser import parse_xlsx' works
_APP_DIR = Path(__file__).resolve().parent
if str(_APP_DIR) not in sys.path:
    sys.path.insert(0, str(_APP_DIR))

from parser import parse_xlsx

# Sample data path: app/test_parser.py -> project root -> data/...
PROJECT_ROOT = _APP_DIR.parent
SAMPLE_FILE = PROJECT_ROOT / "data" / "资源ID账单_2026-07-01_2026-07-04_1783146280975_按资源汇总.csv"
CSV_EXT_FILE = PROJECT_ROOT / "data" / "资源ID账单_2026-07-01_2026-07-04_1783146280975_按资源汇总.csv"


class TestParser(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        """Parse the sample file once for all tests."""
        cls.result = parse_xlsx(str(SAMPLE_FILE))

    # ------------------------------------------------------------------
    # test_parse_sample_returns_correct_structure
    # ------------------------------------------------------------------
    def test_parse_sample_returns_correct_structure(self):
        result = self.result
        self.assertIsInstance(result, dict)
        for key in ("meta", "summary", "records", "by_key",
                     "by_resource_name", "by_model", "timeline"):
            with self.subTest(key=key):
                self.assertIn(key, result)

    # ------------------------------------------------------------------
    # test_summary_values
    # ------------------------------------------------------------------
    def test_summary_values(self):
        summary = self.result["summary"]
        self.assertEqual(summary["total_records"], 33)
        self.assertEqual(summary["api_key_count"], 4)
        self.assertEqual(summary["model_count"], 4)
        self.assertAlmostEqual(summary["cost"], 600.5, delta=0.1)
        self.assertGreater(summary["tokens_total"], 0)

    # ------------------------------------------------------------------
    # test_records_have_required_fields
    # ------------------------------------------------------------------
    def test_records_have_required_fields(self):
        required = ("date", "resource_name", "resource_id", "model",
                    "tokens_input", "tokens_output", "tokens_cache_hit",
                    "tokens_total", "cost")
        for i, record in enumerate(self.result["records"]):
            with self.subTest(record_index=i):
                for field in required:
                    self.assertIn(field, record,
                                  f"Record {i} missing field '{field}'")

    # ------------------------------------------------------------------
    # test_by_resource_name_aggregates
    # ------------------------------------------------------------------
    def test_by_resource_name_aggregates(self):
        by_name = self.result["by_resource_name"]
        self.assertGreater(len(by_name), 0)

        # Check expected resource names exist
        expected_names = {"wsy-main", "harry-opencode", "blsc-ai"}
        found = expected_names & set(by_name.keys())
        self.assertGreaterEqual(len(found), 1,
                                msg="Expected at least one of "
                                    "wsy-main/harry-opencode/blsc-ai")

        for name, entry in by_name.items():
            with self.subTest(resource_name=name):
                self.assertIn("api_keys", entry,
                              f"by_resource_name['{name}'] missing 'api_keys'")
                self.assertIn("models", entry,
                              f"by_resource_name['{name}'] missing 'models'")
                self.assertIsInstance(entry["api_keys"], list)
                self.assertIsInstance(entry["models"], list)
                self.assertGreater(len(entry["api_keys"]), 0)
                self.assertGreater(len(entry["models"]), 0)

    # ------------------------------------------------------------------
    # test_timeline_structure
    # ------------------------------------------------------------------
    def test_timeline_structure(self):
        timeline = self.result["timeline"]
        self.assertGreater(len(timeline), 0)

        for date_str, entry in timeline.items():
            with self.subTest(date=date_str):
                self.assertIn("by_key", entry)
                self.assertIsInstance(entry["by_key"], dict)
                # Each date-level by_key should have at least one entry
                self.assertGreater(len(entry["by_key"]), 0)

    # ------------------------------------------------------------------
    # test_invalid_file_raises_error
    # ------------------------------------------------------------------
    def test_invalid_file_raises_error(self):
        bad_path = str(PROJECT_ROOT / "data" / "nonexistent_file.xlsx")
        with self.assertRaises(FileNotFoundError):
            parse_xlsx(bad_path)

    # ------------------------------------------------------------------
    # test_column_detection
    # ------------------------------------------------------------------
    def test_column_detection(self):
        col_map = self.result["meta"]["column_map"]
        expected_canonical = {
            "date", "resource_name", "resource_id", "model",
            "usage_desc", "cost",
        }
        for canon in expected_canonical:
            with self.subTest(canonical=canon):
                self.assertIn(canon, col_map,
                              f"column_map missing canonical '{canon}'")

    # ------------------------------------------------------------------
    # test_token_parsing
    # ------------------------------------------------------------------
    def test_token_parsing(self):
        for i, record in enumerate(self.result["records"]):
            tokens = record.get("tokens", [])
            if not tokens:
                continue
            with self.subTest(record_index=i):
                for token_entry in tokens:
                    self.assertIn("type", token_entry)
                    self.assertIn("tokens", token_entry)
                    self.assertIn(token_entry["type"],
                                  ("input", "output", "cache_hit"))

    # ------------------------------------------------------------------
    # test_csv_extension_xlsx_parses
    # ------------------------------------------------------------------
    def test_csv_extension_xlsx_parses(self):
        """Parse a CSV-named file that is actually an xlsx."""
        if not CSV_EXT_FILE.exists():
            self.skipTest(f"CSV test file not found: {CSV_EXT_FILE}")
        result = parse_xlsx(str(CSV_EXT_FILE))
        self.assertIsInstance(result, dict)
        self.assertNotIn("error", result)
        self.assertIn("records", result)
        self.assertIn("summary", result)
        # Verify we got data
        records = result["records"]
        self.assertGreater(len(records), 0, "Expected at least 1 data record")
        summary = result["summary"]
        self.assertGreater(summary["total_records"], 0)
        # The file should have comparable structure: at least one api key, model, date
        self.assertGreaterEqual(summary["api_key_count"], 1)
        self.assertGreaterEqual(summary["model_count"], 1)
        self.assertGreaterEqual(summary["date_count"], 1)
        # Verify token fields exist
        self.assertGreater(summary["tokens_total"], 0)


if __name__ == "__main__":
    unittest.main()