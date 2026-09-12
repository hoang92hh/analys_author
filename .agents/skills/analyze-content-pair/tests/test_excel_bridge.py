from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

SCRIPT = Path(__file__).parents[1] / "scripts" / "excel_bridge.py"
SPEC = importlib.util.spec_from_file_location("excel_bridge", SCRIPT)
bridge = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules[SPEC.name] = bridge
SPEC.loader.exec_module(bridge)


class ExcelBridgeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        (self.root / "reference").mkdir()
        (self.root / "criteria").mkdir()
        (self.root / "reference/a.txt").write_text("A", encoding="utf-8")
        (self.root / "reference/b.txt").write_text("B", encoding="utf-8")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Criteria"
        sheet.append(["ID", "Nhóm", "Tiêu chí", "requirement", "Result"])
        sheet.append(["C002", "G2", "Second", "Analyze second", None])
        sheet.append(["C001", "G1", "First", "Analyze first", None])
        notes = workbook.create_sheet("Notes")
        notes["A1"], notes["B1"] = "Keep me", "=1+1"
        workbook.save(self.root / "criteria/source.xlsx")
        workbook.close()
        config = self.root / "config.yaml"
        config.write_text(
            "\n".join([
                "content_a: reference/a.txt", "content_b: reference/b.txt",
                "criteria: criteria/source.xlsx", "criteria_payload: output/payload.json",
                "analysis_results: output/results.json", "output: output/result_analysis.xlsx",
                "result_column: Result",
            ]), encoding="utf-8"
        )
        self.paths = bridge.read_config(config, self.root)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _write_valid_results(self, payload):
        document = {"schema_version": "1.0", "results": [
            {"criterion_key": item["criterion_key"], "criterion_id": item["criterion_id"],
             "result": f"Result for {item['criterion_id']}"}
            for item in payload["criteria"]
        ]}
        self.paths.analysis_results.parent.mkdir(parents=True, exist_ok=True)
        self.paths.analysis_results.write_text(json.dumps(document), encoding="utf-8")

    def test_prepare_preserves_row_order_and_builds_keys(self):
        payload = bridge.prepare(self.paths)
        self.assertEqual([x["criterion_id"] for x in payload["criteria"]], ["C002", "C001"])
        self.assertEqual([x["criterion_key"] for x in payload["criteria"]],
                         ["Criteria!2:C002", "Criteria!3:C001"])
        self.assertFalse(self.paths.output.exists())

    def test_finalize_creates_compact_workbook_in_source_order(self):
        payload = bridge.prepare(self.paths)
        self._write_valid_results(payload)
        self.assertEqual(bridge.finalize(self.paths), 2)
        source = load_workbook(self.paths.criteria, data_only=False)
        output = load_workbook(self.paths.output, data_only=False)
        try:
            self.assertEqual(output.sheetnames, ["Criteria"])
            rows = list(output["Criteria"].iter_rows(values_only=True))
            self.assertEqual(rows, [
                ("ID", "Nhóm", "Tiêu chí", "Requirement", "Result"),
                ("C002", "G2", "Second", "Analyze second", "Result for C002"),
                ("C001", "G1", "First", "Analyze first", "Result for C001"),
            ])
            self.assertIsNone(source["Criteria"]["E2"].value)
            self.assertEqual(source["Notes"]["A1"].value, "Keep me")
        finally:
            source.close()
            output.close()

    def test_empty_result_is_filtered_from_output(self):
        payload = bridge.prepare(self.paths)
        document = {"schema_version": "1.0", "results": [
            {"criterion_key": payload["criteria"][0]["criterion_key"],
             "criterion_id": "C002", "result": "Reusable DNA"},
            {"criterion_key": payload["criteria"][1]["criterion_key"],
             "criterion_id": "C001", "result": "   "},
        ]}
        self.paths.analysis_results.parent.mkdir(parents=True, exist_ok=True)
        self.paths.analysis_results.write_text(json.dumps(document), encoding="utf-8")
        self.assertEqual(bridge.finalize(self.paths), 1)
        output = load_workbook(self.paths.output, data_only=False)
        try:
            rows = list(output.active.iter_rows(values_only=True))
            self.assertEqual(rows, [
                ("ID", "Nhóm", "Tiêu chí", "Requirement", "Result"),
                ("C002", "G2", "Second", "Analyze second", "Reusable DNA"),
            ])
        finally:
            output.close()

    def test_missing_criterion_object_is_rejected_without_output(self):
        payload = bridge.prepare(self.paths)
        item = payload["criteria"][0]
        self.paths.analysis_results.parent.mkdir(parents=True, exist_ok=True)
        self.paths.analysis_results.write_text(json.dumps({"schema_version": "1.0", "results": [{
            "criterion_key": item["criterion_key"], "criterion_id": item["criterion_id"],
            "result": "Only one"
        }]}), encoding="utf-8")
        with self.assertRaises(bridge.BridgeError):
            bridge.finalize(self.paths)
        self.assertFalse(self.paths.output.exists())

    def test_stale_payload_is_rejected(self):
        bridge.prepare(self.paths)
        payload = json.loads(self.paths.criteria_payload.read_text(encoding="utf-8"))
        payload["criteria"][0]["row"] = 99
        self.paths.criteria_payload.write_text(json.dumps(payload), encoding="utf-8")
        self.paths.analysis_results.write_text(
            json.dumps({"schema_version": "1.0", "results": []}), encoding="utf-8")
        with self.assertRaisesRegex(bridge.BridgeError, "stale"):
            bridge.finalize(self.paths)
        self.assertFalse(self.paths.output.exists())


if __name__ == "__main__":
    unittest.main()
