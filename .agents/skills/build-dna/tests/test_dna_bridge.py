"""Synthetic behavior observations test transport/invariants, not learned author DNA."""
import copy
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "dna_bridge.py"
spec = importlib.util.spec_from_file_location("dna_bridge", SCRIPT)
bridge = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = bridge
spec.loader.exec_module(bridge)


class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="dna-test-")
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.paths = bridge.Paths(root / "result_analysis.xlsx", root / "grouped.json",
                                  root / "decisions.json", root / "AUTHOR_DNA.xlsx")
        observations = {
            "Part1": "A has measurements and a label. B compresses measurements and explains the label.",
            "Part2": "A has measurements, but no label. B retains measurements.",
            "Part10": "No measurements or label in A.",
        }
        workbook = Workbook()
        workbook.remove(workbook.active)
        for name in ("Part10", "Part2", "Part1"):
            sheet = workbook.create_sheet(name)
            sheet.append(bridge.INPUT_HEADERS)
            ids = ("C460", "C100", "C999") if name == "Part2" else ("C100", "C460", "C999")
            for criterion_id in ids:
                result = observations[name] if criterion_id == "C100" else (
                    "A has dialogue. B turns explanations into narration." if criterion_id == "C460" else None)
                sheet.append([criterion_id, "Nhóm thử", "Tiêu chí " + criterion_id, "Đối chiếu A→B", result])
        workbook.create_sheet("Notes").append(["Ignored"])
        workbook.save(self.paths.input)
        workbook.close()
        self.original = self.paths.input.read_bytes()
        self.payload = bridge.prepare(self.paths)

    def pattern(self, pattern_id="C100.P01"):
        return {
            "pattern_id": pattern_id, "observation": "Nén measurements khi có material.",
            "status": "keep", "confidence": "High",
            "frequency": {"numerator": 1, "denominator": 2, "ratio": 0.5}, "level": "COMMON",
            "eligible_parts": ["Part1", "Part2"], "assessed_eligible_parts": ["Part1", "Part2"],
            "supporting_parts": ["Part1"], "unknown_occurrence_parts": [],
            "part_assessments": [
                {"part_id": "Part1", "eligibility": "eligible", "occurrence": "present", "reason": "Có measurements và B nén."},
                {"part_id": "Part2", "eligibility": "eligible", "occurrence": "absent", "reason": "Có measurements nhưng B giữ."},
                {"part_id": "Part10", "eligibility": "ineligible", "occurrence": "unknown", "reason": "Không có measurements."},
            ],
            "evidence": [
                {"part_id": "Part1", "result_excerpt": "B compresses measurements", "role": "support"},
                {"part_id": "Part2", "result_excerpt": "B retains measurements.", "role": "counterexample"},
                {"part_id": "Part10", "result_excerpt": "No measurements or label in A.", "role": "ineligible"},
            ],
            "condition": "", "action": "Thường nén thông số phụ thành mô tả quan hệ.",
            "filter_reason": "", "redundant_with": [], "conflicts_with": [], "reconcile_note": "",
        }

    def singleton(self, pattern_id="C100.P02"):
        p = self.pattern(pattern_id)
        p.update(observation="Phát triển giải nghĩa từ label nguồn.",
                 frequency={"numerator": 1, "denominator": 1, "ratio": 1}, level="POSSIBLE",
                 eligible_parts=["Part1"], assessed_eligible_parts=["Part1"],
                 condition="Source có label cần giải nghĩa.", action="Có thể diễn giải ý nghĩa label từ điểm neo source.")
        p["part_assessments"][0]["reason"] = "Source có label và B giải nghĩa."
        p["part_assessments"][1].update(eligibility="ineligible", occurrence="unknown", reason="Source không có label.")
        p["evidence"][0]["result_excerpt"] = "B compresses measurements and explains the label."
        p["evidence"][1].update(result_excerpt="no label", role="ineligible")
        return p

    def unassessed(self):
        p = self.pattern("C999.P01")
        p.update(status="drop", confidence="Low", observation="Chưa có observation.",
                 frequency={"numerator": 0, "denominator": 0, "ratio": None}, level="UNASSESSED",
                 eligible_parts=[], assessed_eligible_parts=[], supporting_parts=[],
                 unknown_occurrence_parts=[], evidence=[], action="", filter_reason="Result trống.")
        for a in p["part_assessments"]:
            a.update(eligibility="unknown", occurrence="unknown", reason="Result trống.")
        return p

    def document(self):
        criteria = [
            {"criterion_id": c["criterion_id"], "synthesis": "Đã xét các observations.",
             "patterns": [], "dna_result": None} for c in self.payload["criteria"]
        ]
        criteria[0].update(patterns=[self.pattern()], dna_result={
            "pattern_ids": ["C100.P01"], "action": "Thường nén thông số phụ thành mô tả quan hệ."})
        criteria[2]["patterns"] = [self.unassessed()]
        return {"schema_version": "3.0", "payload_digest": self.payload["payload_digest"],
                "metadata": {"corpus_scope": "Corpus thử về hiện vật.", "limitations": "",
                             "reconciliation_completed": True, "merge_completed": True},
                "criteria": criteria}

    def merge(self, item, **fields):
        survivors = [p for p in item["patterns"] if p["status"] == "keep"]
        item["dna_result"] = {"pattern_ids": [p["pattern_id"] for p in survivors],
                              "action": "Thường nén thông số phụ; có thể giải nghĩa label nếu source có label.",
                              **fields} if survivors else None

    def save(self, document):
        bridge.write_json(self.paths.dna_results, document)

    def edit_input(self, edit):
        workbook = load_workbook(self.paths.input)
        try:
            edit(workbook)
            workbook.save(self.paths.input)
        finally:
            workbook.close()

    def test_grouping_preserves_blanks_coordinates_input_and_numeric_part_order(self):
        self.assertEqual(self.payload["parts"], ["Part1", "Part2", "Part10"])
        self.assertEqual(self.payload["ignored_sheets"], ["Notes"])
        self.assertEqual(self.payload["criteria"][0]["observations"][1]["result_cell"], "Part2!E3")
        self.assertEqual([o["result"] for o in self.payload["criteria"][2]["observations"]], ["", "", ""])
        self.assertEqual(self.paths.input.read_bytes(), self.original)

    def test_multiple_patterns_merge_to_one_row_four_columns_no_extra_sheet(self):
        doc = self.document()
        item = doc["criteria"][0]
        item["patterns"].append(self.singleton())
        self.merge(item, condition_variation="Giải nghĩa chỉ khi source có label.")
        item["dna_result"]["action"] = "=Thường nén; có thể giải nghĩa label khi source có label."
        self.save(doc)
        self.assertEqual(bridge.finalize(self.paths), 1)
        workbook = load_workbook(self.paths.output)
        try:
            self.assertEqual(workbook.sheetnames, ["AUTHOR_DNA"])
            sheet = workbook.active
            self.assertEqual((sheet.max_row, sheet.max_column), (2, 4))
            self.assertEqual([c.value for c in sheet[1]], ["ID", "Tiêu chí", "DNA Rule", "Condition/Variation"])
            self.assertEqual(sheet["A2"].value, "C100")
            self.assertEqual(sheet["C2"].data_type, "s")
            self.assertTrue(sheet["C2"].value.startswith("ACTION\n="))
            self.assertNotIn("C100.P", sheet["C2"].value)
            self.assertEqual(sheet.freeze_panes, "C2")
        finally:
            workbook.close()
        self.assertEqual(self.paths.input.read_bytes(), self.original)

    def test_single_observation_possible_with_material_trigger(self):
        doc = self.document()
        doc["criteria"][0]["patterns"] = [self.singleton("C100.P01")]
        self.merge(doc["criteria"][0], condition_variation="Source có label.")
        self.assertEqual(len(bridge.validate_results(self.payload, doc)), 1)
        doc["criteria"][0]["patterns"][0]["level"] = "CORE"
        with self.assertRaisesRegex(bridge.BridgeError, "singleton"):
            bridge.validate_results(self.payload, doc)

    def test_singleton_without_source_condition_rejected(self):
        doc = self.document()
        p = self.singleton("C100.P01")
        p["condition"] = ""
        doc["criteria"][0]["patterns"] = [p]
        self.merge(doc["criteria"][0])
        with self.assertRaisesRegex(bridge.BridgeError, "condition"):
            bridge.validate_results(self.payload, doc)

    def test_unknown_occurrence_excluded_from_denominator_not_absent(self):
        payload = copy.deepcopy(self.payload)
        payload["criteria"][0]["observations"][1]["result"] = "A has measurements and a label. Label treatment is not described."
        doc = self.document()
        p = self.singleton("C100.P01")
        p.update(eligible_parts=["Part1", "Part2"], unknown_occurrence_parts=["Part2"],
                 limitations="Part2 có label nhưng không biết cách xử lý.")
        p["part_assessments"][1].update(eligibility="eligible", reason="Có label nhưng occurrence chưa biết.")
        p["evidence"][1].update(result_excerpt="A has measurements and a label.", role="eligible")
        doc["criteria"][0]["patterns"] = [p]
        self.merge(doc["criteria"][0], condition_variation="Có label mới áp dụng.")
        self.assertEqual(len(bridge.validate_results(payload, doc)), 1)
        p.update(frequency={"numerator": 1, "denominator": 2, "ratio": 0.5}, level="COMMON")
        with self.assertRaisesRegex(bridge.BridgeError, "assessed eligible"):
            bridge.validate_results(payload, doc)

    def test_unknown_eligibility_requires_limitation_not_ineligible_inference(self):
        doc = self.document()
        p = doc["criteria"][0]["patterns"][0]
        p["part_assessments"][2].update(eligibility="unknown", reason="Chưa biết material.")
        p["evidence"].pop()
        with self.assertRaisesRegex(bridge.BridgeError, "limitations"):
            bridge.validate_results(self.payload, doc)
        p["limitations"] = "Eligibility Part10 chưa xác định."
        self.assertEqual(len(bridge.validate_results(self.payload, doc)), 1)

    def test_all_eligible_occurrences_unknown_has_zero_denominator(self):
        doc = self.document()
        p = doc["criteria"][0]["patterns"][0]
        p.update(status="drop", confidence="Low", filter_reason="Occurrence chưa biết.",
                 assessed_eligible_parts=[], supporting_parts=[], unknown_occurrence_parts=["Part1", "Part2"],
                 frequency={"numerator": 0, "denominator": 0, "ratio": None}, level="UNASSESSED")
        for a in p["part_assessments"][:2]:
            a["occurrence"] = "unknown"
        p["evidence"][0].update(role="eligible", result_excerpt="A has measurements")
        p["evidence"][1].update(role="eligible", result_excerpt="A has measurements")
        self.merge(doc["criteria"][0])
        self.assertEqual(bridge.validate_results(self.payload, doc), [])

    def test_frequency_bands_and_one_support_on_three_assessed_parts(self):
        for n, d, c, level in [(3,3,"","CORE"),(4,5,"","DEFAULT"),(2,3,"","COMMON"),
                               (1,3,"","POSSIBLE"),(3,10,"","POSSIBLE"),(1,4,"","INSTANCE_SPECIFIC"),
                               (1,4,"material trigger","POSSIBLE"),(1,1,"trigger","POSSIBLE"),(0,0,"","UNASSESSED")]:
            with self.subTest(n=n,d=d):
                self.assertEqual(bridge.expected_level(n,d,c),level)
        payload = copy.deepcopy(self.payload)
        payload["criteria"][0]["observations"][2]["result"] = "A has measurements. B retains measurements."
        doc = self.document()
        p = doc["criteria"][0]["patterns"][0]
        p.update(eligible_parts=self.payload["parts"], assessed_eligible_parts=self.payload["parts"],
                 frequency={"numerator":1,"denominator":3,"ratio":1/3},level="POSSIBLE")
        p["part_assessments"][2].update(eligibility="eligible",occurrence="absent",reason="Có measurements nhưng không nén.")
        p["evidence"][2].update(result_excerpt="B retains measurements.",role="counterexample")
        self.assertEqual(len(bridge.validate_results(payload,doc)),1)

    def test_frequency_and_part_list_tampering_rejected(self):
        for mode in ("ratio","bool","count","order","duplicate","assessed","support"):
            doc = self.document()
            p = doc["criteria"][0]["patterns"][0]
            if mode=="ratio": p["frequency"]["ratio"]=1
            elif mode=="bool": p["frequency"]["numerator"]=True
            elif mode=="count": p["frequency"]["denominator"]=3
            elif mode=="order": p["eligible_parts"].reverse()
            elif mode=="duplicate": p["supporting_parts"].append("Part1")
            elif mode=="assessed": p["assessed_eligible_parts"].pop()
            else: p["supporting_parts"]=["Part2"]
            with self.subTest(mode=mode),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_absence_and_ineligibility_need_explicit_evidence(self):
        for index in (1,2):
            doc=self.document()
            doc["criteria"][0]["patterns"][0]["evidence"].pop(index)
            with self.subTest(index=index),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_low_confidence_and_unresolved_kept_pattern_rejected(self):
        for mode in ("Low","unresolved"):
            doc=self.document()
            p=doc["criteria"][0]["patterns"][0]
            if mode=="Low": p["confidence"]="Low"
            else: p["evidence"].append({"part_id":"Part1","result_excerpt":"A has measurements","role":"unresolved"})
            with self.subTest(mode=mode),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_drop_one_pattern_preserves_criterion_but_all_dropped_omits_it(self):
        doc=self.document()
        p=self.singleton()
        p.update(status="drop",confidence="Low",filter_reason="Chưa đủ căn cứ.")
        doc["criteria"][0]["patterns"].append(p)
        self.assertEqual(len(bridge.validate_results(self.payload,doc)),1)
        doc["criteria"][0]["patterns"][0].update(status="drop",confidence="Low",filter_reason="Chưa đủ căn cứ.")
        self.merge(doc["criteria"][0])
        self.save(doc)
        self.assertEqual(bridge.finalize(self.paths),0)
        workbook=load_workbook(self.paths.output)
        try:
            self.assertEqual(workbook.active.max_row,1)
        finally:
            workbook.close()

    def test_merge_must_cover_exactly_all_kept_patterns_and_conditions(self):
        for mode in ("missing","dropped","duplicate","condition","null","empty_action"):
            doc=self.document()
            item=doc["criteria"][0]
            item["patterns"].append(self.singleton())
            self.merge(item,condition_variation="Source có label.")
            if mode=="missing": item["dna_result"]["pattern_ids"].pop()
            elif mode=="dropped": item["patterns"][1].update(status="drop",confidence="Low",filter_reason="Loại.")
            elif mode=="duplicate": item["dna_result"]["pattern_ids"].append("C100.P01")
            elif mode=="condition": item["dna_result"].pop("condition_variation")
            elif mode=="null": item["dna_result"]=None
            else: item["dna_result"]["action"]=""
            with self.subTest(mode=mode),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_optional_sections_need_kept_pattern_basis_no_empty_labels(self):
        doc=self.document()
        merged=doc["criteria"][0]["dna_result"]
        merged.update(when="",selection="",validation="",condition_variation="")
        kept=bridge.validate_results(self.payload,doc)
        rule=bridge.output_rows(self.payload,kept)[1][2]
        self.assertEqual(rule,"ACTION\n"+merged["action"])
        for field in ("when","selection","validation"):
            bad=copy.deepcopy(doc)
            bad["criteria"][0]["dna_result"][field]="Unfounded instruction"
            with self.subTest(field=field),self.assertRaisesRegex(bridge.BridgeError,"basis|condition"):
                bridge.validate_results(self.payload,bad)
        doc["criteria"][0]["patterns"][0].update(selection="Ưu tiên thông số chính.",validation="Kiểm tra còn quan hệ chính.")
        merged.update(selection="Ưu tiên thông số chính.",validation="Kiểm tra còn quan hệ chính.")
        self.assertIn("SELECTION\n",bridge.output_rows(self.payload,bridge.validate_results(self.payload,doc))[1][2])

    def test_workflow_mapping_rejected_at_pattern_and_merge_levels(self):
        for target in ("pattern","merged"):
            doc=self.document()
            item=doc["criteria"][0]
            (item["patterns"][0] if target=="pattern" else item["dna_result"])["target_step"]="STEP 3"
            with self.subTest(target=target),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_old_review_schemas_and_incomplete_reconcile_merge_rejected(self):
        for mode in ("1.0","2.0","reconciliation_completed","merge_completed"):
            doc=self.document()
            if mode in ("1.0","2.0"): doc["schema_version"]=mode
            else: doc["metadata"][mode]=False
            with self.subTest(mode=mode),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_cross_criterion_evidence_not_allowed_even_with_reconcile_link(self):
        doc=self.document()
        p=self.pattern("C460.P01")
        p["redundant_with"]=["C100.P01"]
        p["reconcile_note"]="Overlap."
        doc["criteria"][1]["patterns"]=[p]
        self.merge(doc["criteria"][1])
        with self.assertRaisesRegex(bridge.BridgeError,"same-criterion"):
            bridge.validate_results(self.payload,doc)

    def test_invalid_reconcile_references_and_cycles_rejected(self):
        for mode in ("unknown","self","cycle"):
            doc=self.document()
            p=doc["criteria"][0]["patterns"][0]
            p["reconcile_note"]="Overlap."
            if mode=="unknown": p["redundant_with"]=["C100.P99"]
            elif mode=="self": p["redundant_with"]=["C100.P01"]
            else:
                other=copy.deepcopy(p)
                other["pattern_id"]="C100.P02"
                p["redundant_with"]=["C100.P02"]
                other["redundant_with"]=["C100.P01"]
                doc["criteria"][0]["patterns"].append(other)
                self.merge(doc["criteria"][0])
            with self.subTest(mode=mode),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_conflicting_kept_patterns_need_conditions_and_notes(self):
        doc=self.document()
        item=doc["criteria"][0]
        p=item["patterns"][0]
        other=self.singleton()
        p.update(conflicts_with=["C100.P02"],reconcile_note="Different source dimensions.")
        other.update(conflicts_with=["C100.P01"],reconcile_note="Different source dimensions.")
        item["patterns"].append(other)
        self.merge(item,condition_variation="Label/measurements riêng.")
        with self.assertRaisesRegex(bridge.BridgeError,"Conflicting"):
            bridge.validate_results(self.payload,doc)
        p["condition"]="Source có measurements phụ."
        self.assertEqual(len(bridge.validate_results(self.payload,doc)),1)

    def test_all_criteria_and_part_assessments_covered_unique_ids(self):
        for mode in ("criterion","assessment","duplicate","prefix"):
            doc=self.document()
            if mode=="criterion": doc["criteria"].pop()
            elif mode=="assessment": doc["criteria"][0]["patterns"][0]["part_assessments"].pop()
            elif mode=="duplicate":
                doc["criteria"][0]["patterns"].append(copy.deepcopy(doc["criteria"][0]["patterns"][0]))
            else: doc["criteria"][0]["patterns"][0]["pattern_id"]="C460.P01"
            with self.subTest(mode=mode),self.assertRaises(bridge.BridgeError):
                bridge.validate_results(self.payload,doc)

    def test_replace_only_dna_sheet_preserves_other_sheets_and_repeat_filename(self):
        workbook=Workbook()
        notes=workbook.active
        notes.title="Notes"
        notes.append(["Giữ nguyên",12])
        notes["A1"].font=Font(bold=True,color="FF0000")
        notes["B1"].number_format="0.00"
        notes["A2"]="=AUTHOR_DNA!A2"
        notes["A3"].comment=Comment("Ghi chú","Reviewer")
        notes["A4"]="Link"
        notes["A4"].hyperlink="https://example.invalid"
        notes.merge_cells("C3:D3")
        notes.column_dimensions["A"].width=37
        notes.row_dimensions[1].height=35
        notes.freeze_panes="B2"
        chart=BarChart()
        chart.add_data(Reference(notes,min_col=2,min_row=1,max_row=1))
        notes.add_chart(chart,"F1")
        workbook.create_sheet("AUTHOR_DNA").append(["Old", "extra", "columns", "to", "remove"])
        hidden=workbook.create_sheet("Private")
        hidden["A1"]=123
        hidden.sheet_state="hidden"
        workbook.save(self.paths.output)
        workbook.close()
        existing=load_workbook(self.paths.output)
        before={name:bridge.sheet_snapshot(existing[name]) for name in ("Notes","Private")}
        existing.close()
        doc=self.document()
        self.save(doc)
        self.assertEqual(bridge.finalize(self.paths),1)
        self.assertEqual(bridge.finalize(self.paths),1)
        check=load_workbook(self.paths.output)
        try:
            self.assertEqual(check.sheetnames,["Notes","AUTHOR_DNA","Private"])
            self.assertEqual(check.active.title,"Notes")
            self.assertEqual(check["AUTHOR_DNA"].max_column,4)
            self.assertEqual(check["AUTHOR_DNA"].max_row,2)
            self.assertEqual({name:bridge.sheet_snapshot(check[name]) for name in before},before)
            self.assertEqual(check["Notes"]["A2"].value,"=AUTHOR_DNA!A2")
            self.assertEqual(len(check["Notes"]._charts),1)
        finally:
            check.close()
        self.assertEqual(list(self.paths.output.parent.glob("AUTHOR_DNA*.xlsx")),[self.paths.output])
        self.assertEqual(self.paths.input.read_bytes(),self.original)

    def test_existing_workbook_without_dna_sheet_gets_only_that_new_sheet(self):
        workbook=Workbook()
        workbook.active.title="Notes"
        workbook.active["A1"]="Keep"
        workbook.save(self.paths.output)
        workbook.close()
        self.save(self.document())
        bridge.finalize(self.paths)
        check=load_workbook(self.paths.output)
        try:
            self.assertEqual(check.sheetnames,["AUTHOR_DNA","Notes"])
            self.assertEqual(check["Notes"]["A1"].value,"Keep")
            self.assertEqual(check.active.title,"Notes")
        finally:
            check.close()

    def test_failure_and_concurrent_changes_preserve_existing_output(self):
        self.save(self.document())
        bridge.finalize(self.paths)
        original=self.paths.output.read_bytes()
        doc=self.document()
        doc["criteria"][0]["dna_result"]["action"]="x"*32767
        self.save(doc)
        with self.assertRaisesRegex(bridge.BridgeError,"cell limits"):
            bridge.finalize(self.paths)
        self.assertEqual(self.paths.output.read_bytes(),original)
        self.save(self.document())
        with patch.object(bridge.os,"replace",side_effect=PermissionError("Workbook busy")):
            with self.assertRaises(PermissionError):
                bridge.finalize(self.paths)
        self.assertEqual(self.paths.output.read_bytes(),original)
        actual_save=Workbook.save
        concurrent=original+b"concurrent edit"
        def save_then_edit(workbook,path):
            actual_save(workbook,path)
            self.paths.output.write_bytes(concurrent)
        with patch.object(Workbook,"save",save_then_edit):
            with self.assertRaisesRegex(bridge.BridgeError,"Output changed"):
                bridge.finalize(self.paths)
        self.assertEqual(self.paths.output.read_bytes(),concurrent)
        self.assertEqual(list(self.paths.output.parent.glob(".AUTHOR_DNA.*.xlsx")),[])

    def test_input_change_during_export_preserves_output(self):
        self.save(self.document())
        bridge.finalize(self.paths)
        old=self.paths.output.read_bytes()
        actual_save=Workbook.save
        def save_then_edit(workbook,path):
            actual_save(workbook,path)
            self.paths.input.write_bytes(self.original+b"changed")
        with patch.object(Workbook,"save",save_then_edit):
            with self.assertRaisesRegex(bridge.BridgeError,"Input changed"):
                bridge.finalize(self.paths)
        self.assertEqual(self.paths.output.read_bytes(),old)

    def test_stale_payload_and_wrong_digest_preserve_output(self):
        self.save(self.document())
        bridge.finalize(self.paths)
        old=self.paths.output.read_bytes()
        changed=copy.deepcopy(self.payload)
        changed["criteria"][0]["observations"][0]["result"]="Modified"
        bridge.write_json(self.paths.grouped_payload,changed)
        with self.assertRaisesRegex(bridge.BridgeError,"stale"):
            bridge.finalize(self.paths)
        bridge.write_json(self.paths.grouped_payload,self.payload)
        doc=self.document()
        doc["payload_digest"]="0"*64
        self.save(doc)
        with self.assertRaises(bridge.BridgeError):
            bridge.finalize(self.paths)
        self.assertEqual(self.paths.output.read_bytes(),old)

    def test_invalid_input_does_not_replace_prepared_payload(self):
        old=self.paths.grouped_payload.read_bytes()
        edits=[
            lambda w:w["Part2"].delete_rows(2),
            lambda w:setattr(w["Part2"]["A3"],"value","C460"),
            lambda w:setattr(w["Part2"]["D3"],"value","Different"),
            lambda w:setattr(w["Part2"]["E3"],"value","=1+1"),
            lambda w:setattr(w["Part2"]["E3"],"value",12),
            lambda w:setattr(w["Part2"]["A1"],"value","Wrong"),
            lambda w:setattr(w["Part2"],"title","Part02"),
        ]
        for edit in edits:
            self.paths.input.write_bytes(self.original)
            self.edit_input(edit)
            with self.assertRaises(bridge.BridgeError):
                bridge.prepare(self.paths)
            self.assertEqual(self.paths.grouped_payload.read_bytes(),old)

    def test_config_rejects_aliases_and_versioned_final_names(self):
        config=self.paths.input.parent/"config.yaml"
        text="input: result_analysis.xlsx\ngrouped_payload: grouped.json\ndna_results: decisions.json\noutput: AUTHOR_DNA.xlsx\n"
        config.write_text(text,encoding="utf-8")
        self.assertEqual(bridge.read_config(config,config.parent),self.paths)
        for output in ("result_analysis.xlsx","AUTHOR_DNA_v2.xlsx","AUTHOR_DNA_new.xlsx"):
            config.write_text(text.replace("output: AUTHOR_DNA.xlsx","output: "+output),encoding="utf-8")
            with self.subTest(output=output),self.assertRaises(bridge.BridgeError):
                bridge.read_config(config,config.parent)

    def test_duplicate_json_keys_rejected(self):
        self.paths.dna_results.write_text('{"criteria":[],"criteria":[]}',encoding="utf-8")
        with self.assertRaises(bridge.BridgeError):
            bridge.load_json(self.paths.dna_results)

    def test_review_schema_internal_references_resolve(self):
        schema=json.loads((SCRIPT.parents[1]/"schemas/dna_results.schema.json").read_text(encoding="utf-8"))
        def walk(value):
            if isinstance(value,dict):
                if "$ref" in value:
                    self.assertTrue(value["$ref"].startswith("#/"))
                    node=schema
                    for key in value["$ref"][2:].split("/"):
                        node=node[key.replace("~1","/").replace("~0","~")]
                    self.assertIsInstance(node,dict)
                for child in value.values():
                    walk(child)
            elif isinstance(value,list):
                for child in value:
                    walk(child)
        walk(schema)
        self.assertEqual(schema["properties"]["schema_version"]["const"],bridge.VERSION)


if __name__=="__main__":
    unittest.main()
