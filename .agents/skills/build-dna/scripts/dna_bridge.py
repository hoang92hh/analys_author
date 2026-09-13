#!/usr/bin/env python3
"""Group Part observations and export LLM-authored DNA; no semantic synthesis."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

VERSION = "3.0"
INPUT_HEADERS = ["ID", "Nhóm", "Tiêu chí", "Requirement", "Result"]
OUTPUT_HEADERS = ["ID", "Tiêu chí", "DNA Rule", "Condition/Variation"]
PATTERN_FIELDS = {"pattern_id", "observation", "status", "confidence", "frequency", "level",
                  "eligible_parts", "assessed_eligible_parts", "supporting_parts",
                  "unknown_occurrence_parts", "part_assessments", "evidence",
                  "condition", "action", "filter_reason", "redundant_with",
                  "conflicts_with", "reconcile_note"}
PATTERN_OPTIONAL = {"selection", "validation", "limitations"}
MERGED_FIELDS = {"pattern_ids", "action"}
MERGED_OPTIONAL = {"when", "selection", "validation", "condition_variation"}
RESULT_SHEET = "AUTHOR_DNA"
PART = re.compile(r"Part([1-9][0-9]*)\Z")
CRITERION = re.compile(r"C[0-9]+\Z")


class BridgeError(RuntimeError):
    pass


@dataclass(frozen=True)
class Paths:
    input: Path
    grouped_payload: Path
    dna_results: Path
    output: Path


def read_config(config: Path, root: Path) -> Paths:
    # Same flat key: value convention as Skill 1; no YAML dependency required.
    values = {}
    for number, raw in enumerate(config.read_text(encoding="utf-8-sig").splitlines(), 1):
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        if ":" not in raw:
            raise BridgeError(f"Invalid config line {number}")
        key, value = (piece.strip() for piece in raw.split(":", 1))
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        if key in values or not value:
            raise BridgeError(f"Duplicate key or empty value at config line {number}")
        values[key] = value
    if set(values) != {"input", "grouped_payload", "dna_results", "output"}:
        raise BridgeError("Config requires only input, grouped_payload, dna_results, output")
    paths = Paths(**{key: (root / value).resolve() for key, value in values.items()})
    validate_paths(paths)
    return paths


def validate_paths(paths: Paths) -> None:
    resolved = [path.resolve() for path in vars(paths).values()]
    if len(set(resolved)) != len(resolved):
        raise BridgeError("Input and artifact paths must all be distinct")
    existing = [path for path in resolved if path.exists()]
    for i, path in enumerate(existing):
        if any(os.path.samefile(path, other) for other in existing[i + 1:]):
            raise BridgeError("Input and artifact paths must not alias the same file")
    if paths.input.suffix.lower() != ".xlsx" or paths.output.suffix.lower() != ".xlsx":
        raise BridgeError("Input and output must be .xlsx files")
    if paths.output.name != "AUTHOR_DNA.xlsx":
        raise BridgeError("Final output must be AUTHOR_DNA.xlsx; replace its AUTHOR_DNA sheet, do not version the filename")
    if any(path.suffix.lower() != ".json" for path in (paths.grouped_payload, paths.dna_results)):
        raise BridgeError("Payload and decisions must be .json files")


def digest(payload: dict) -> str:
    body = {key: value for key, value in payload.items() if key != "payload_digest"}
    return hashlib.sha256(json.dumps(body, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode("utf-8")).hexdigest()


def extract(paths: Paths) -> dict:
    validate_paths(paths)
    source_hash = hashlib.sha256(paths.input.read_bytes()).hexdigest()
    workbook = load_workbook(paths.input, read_only=True, data_only=False)
    try:
        sheets = sorted((sheet for sheet in workbook if PART.fullmatch(sheet.title)),
                        key=lambda sheet: int(PART.fullmatch(sheet.title).group(1)))
        if not sheets:
            raise BridgeError("Input has no Part sheets (Part1, Part2, ...)")
        ignored = [sheet.title for sheet in workbook if not PART.fullmatch(sheet.title)]
        malformed = [name for name in ignored if name.lower().startswith("part")]
        if malformed:
            raise BridgeError(f"Invalid Part sheet names: {malformed}")
        grouped: dict[str, dict] = {}
        part_ids = []
        for sheet in sheets:
            part_ids.append(sheet.title)
            rows = sheet.iter_rows()
            headers = [cell.value for cell in next(rows, [])]
            if headers != INPUT_HEADERS:
                raise BridgeError(f"{sheet.title}: expected exactly {INPUT_HEADERS}")
            seen = set()
            for row in rows:
                values = [cell.value for cell in row]
                if all(value in (None, "") for value in values):
                    continue
                if any(cell.data_type in ("f", "e") for cell in row):
                    raise BridgeError(f"{sheet.title}!{row[0].row}: formulas/errors are not observations")
                if any(value is not None and not isinstance(value, str) for value in values):
                    raise BridgeError(f"{sheet.title}!{row[0].row}: criterion fields must be text")
                criterion_id, group, criterion, requirement, result = values
                if not isinstance(criterion_id, str) or not CRITERION.fullmatch(criterion_id):
                    raise BridgeError(f"{sheet.title}!{row[0].row}: invalid criterion ID")
                if criterion_id in seen:
                    raise BridgeError(f"{sheet.title}: duplicate ID {criterion_id}")
                if not criterion or not criterion.strip() or not requirement or not requirement.strip():
                    raise BridgeError(f"{sheet.title}: missing criterion name or Requirement for {criterion_id}")
                seen.add(criterion_id)
                metadata = {"criterion_id": criterion_id, "group": group,
                            "criterion": criterion, "requirement": requirement}
                if len(part_ids) == 1:
                    grouped[criterion_id] = {**metadata, "observations": []}
                elif criterion_id not in grouped:
                    raise BridgeError(f"{sheet.title}: unknown criterion {criterion_id}")
                elif any(grouped[criterion_id][key] != value for key, value in metadata.items()):
                    raise BridgeError(f"{sheet.title}: inconsistent metadata for {criterion_id}")
                grouped[criterion_id]["observations"].append({
                    "part_id": sheet.title, "row": row[0].row,
                    "result_cell": f"{sheet.title}!E{row[0].row}",
                    "result": result if result is not None else "",
                })
            if not seen:
                raise BridgeError(f"{sheet.title}: no criteria")
            if seen != set(grouped):
                raise BridgeError(f"{sheet.title}: missing criteria {sorted(set(grouped) - seen)}")
        payload = {"schema_version": VERSION,
                   "source": {"input": str(paths.input.resolve()), "sha256": source_hash},
                   "parts": part_ids, "ignored_sheets": ignored,
                   "criteria": list(grouped.values())}
        payload["payload_digest"] = digest(payload)
        return payload
    finally:
        workbook.close()


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(name, path)
    finally:
        Path(name).unlink(missing_ok=True)


def load_json(path: Path) -> Any:
    def unique(pairs):
        value = {}
        for key, item in pairs:
            if key in value:
                raise BridgeError(f"Duplicate JSON key {key!r} in {path}")
            value[key] = item
        return value
    return json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=unique)


def prepare(paths: Paths) -> dict:
    payload = extract(paths)
    write_json(paths.grouped_payload, payload)
    return payload



def require_object(value: Any, required: set, optional: set | None = None) -> None:
    if not isinstance(value, dict) or not required <= set(value) or set(value) - required - (optional or set()):
        raise BridgeError(f"Invalid fields: required {sorted(required)}, optional {sorted(optional or set())}")


def require_text(value: Any, label: str, nonempty: bool = False) -> None:
    if not isinstance(value, str) or (nonempty and not value.strip()):
        raise BridgeError(f"{label}: expected {'nonempty ' if nonempty else ''}text")


def part_list(value: Any, parts: list[str], label: str) -> list[str]:
    if not isinstance(value, list) or any(not isinstance(part, str) for part in value):
        raise BridgeError(f"{label}: expected Part array")
    if value != [part for part in parts if part in value]:
        raise BridgeError(f"{label}: unknown, duplicate or out-of-order Parts")
    return value


def expected_level(numerator: int, denominator: int, condition: str) -> str:
    if denominator == 0:
        return "UNASSESSED"
    if numerator == denominator == 1:
        return "POSSIBLE"
    if numerator == denominator:
        return "CORE"
    for threshold, level in ((80, "DEFAULT"), (50, "COMMON"), (30, "POSSIBLE")):
        if numerator * 100 >= threshold * denominator:
            return level
    return "POSSIBLE" if condition.strip() else "INSTANCE_SPECIFIC"


def validate_pattern(pattern: Any, criterion: dict, parts: list[str]) -> None:
    require_object(pattern, PATTERN_FIELDS, PATTERN_OPTIONAL)
    for field in {"pattern_id", "observation", "status", "confidence", "level",
                  "condition", "action", "filter_reason", "reconcile_note"} | (PATTERN_OPTIONAL & set(pattern)):
        require_text(pattern[field], field)
    label = pattern["pattern_id"]
    if not re.fullmatch(re.escape(criterion["criterion_id"]) + r"\.P[0-9]{2,}", label):
        raise BridgeError(f"{label}: pattern_id must retain its Criterion ID")
    require_text(pattern["observation"], label + " observation", True)
    if pattern["status"] not in ("keep", "drop") or pattern["confidence"] not in ("High", "Medium", "Low"):
        raise BridgeError(f"{label}: invalid status/confidence")
    assessments = pattern["part_assessments"]
    if not isinstance(assessments, list) or [
            item.get("part_id") if isinstance(item, dict) else None for item in assessments] != parts:
        raise BridgeError(f"{label}: assess every Part once in corpus order")
    for item in assessments:
        require_object(item, {"part_id", "eligibility", "occurrence", "reason"})
        require_text(item["reason"], label + " assessment reason", True)
        if item["eligibility"] not in ("eligible", "ineligible", "unknown") or item["occurrence"] not in ("present", "absent", "unknown"):
            raise BridgeError(f"{label}: invalid eligibility/occurrence")
        if item["eligibility"] != "eligible" and item["occurrence"] != "unknown":
            raise BridgeError(f"{label}: occurrence must be unknown outside established eligibility")
    expected = {
        "eligible_parts": [a["part_id"] for a in assessments if a["eligibility"] == "eligible"],
        "assessed_eligible_parts": [a["part_id"] for a in assessments if a["eligibility"] == "eligible" and a["occurrence"] != "unknown"],
        "supporting_parts": [a["part_id"] for a in assessments if a["occurrence"] == "present"],
        "unknown_occurrence_parts": [a["part_id"] for a in assessments if a["eligibility"] == "eligible" and a["occurrence"] == "unknown"],
    }
    for field, values in expected.items():
        if part_list(pattern[field], parts, label + " " + field) != values:
            raise BridgeError(f"{label}: {field} must match assessments")
    require_object(pattern["frequency"], {"numerator", "denominator", "ratio"})
    f = pattern["frequency"]
    n, d, ratio = f["numerator"], f["denominator"], f["ratio"]
    if type(n) is not int or type(d) is not int or (n, d) != (
            len(expected["supporting_parts"]), len(expected["assessed_eligible_parts"])):
        raise BridgeError(f"{label}: frequency must count supporting / assessed eligible Parts; exclude unknown")
    if d == 0:
        if ratio is not None:
            raise BridgeError(f"{label}: zero denominator requires ratio null")
    elif type(ratio) not in (int, float) or not math.isfinite(ratio) or not math.isclose(ratio, n / d, rel_tol=0, abs_tol=1e-6):
        raise BridgeError(f"{label}: incorrect frequency ratio")
    if pattern["level"] != expected_level(n, d, pattern["condition"]):
        raise BridgeError(f"{label}: incorrect frequency level or singleton override")
    observations = {a["part_id"]: a["result"] for a in criterion["observations"]}
    evidence = pattern["evidence"]
    if not isinstance(evidence, list):
        raise BridgeError(f"{label}: evidence must be an array")
    roles = {part: set() for part in parts}
    by_part = {a["part_id"]: a for a in assessments}
    for ev in evidence:
        require_object(ev, {"part_id", "result_excerpt", "role"})
        for key, value in ev.items():
            require_text(value, label + " evidence " + key, True)
        part, quote, role = ev["part_id"], ev["result_excerpt"], ev["role"]
        if part not in observations or quote not in observations[part]:
            raise BridgeError(f"{label}: quote not found in same-criterion Part Result")
        if role not in ("support", "eligible", "ineligible", "counterexample", "condition", "unresolved"):
            raise BridgeError(f"{label}: invalid evidence role")
        roles[part].add(role)
        if role == "support" and by_part[part]["occurrence"] != "present":
            raise BridgeError(f"{label}: support disagrees with occurrence")
        if role == "counterexample" and by_part[part]["occurrence"] != "absent":
            raise BridgeError(f"{label}: counterexample requires eligible absent")
        if role == "ineligible" and by_part[part]["eligibility"] != "ineligible":
            raise BridgeError(f"{label}: ineligible evidence disagrees with eligibility")
    for part, assessment in by_part.items():
        present_roles = roles[part]
        if assessment["eligibility"] == "eligible" and not present_roles & {"support", "eligible", "counterexample", "condition"}:
            raise BridgeError(f"{label}: eligible needs material evidence")
        if assessment["eligibility"] == "ineligible" and "ineligible" not in present_roles:
            raise BridgeError(f"{label}: ineligible needs explicit evidence; otherwise unknown")
        if assessment["occurrence"] == "present" and "support" not in present_roles:
            raise BridgeError(f"{label}: present needs support evidence")
        if assessment["occurrence"] == "absent" and "counterexample" not in present_roles:
            raise BridgeError(f"{label}: absent needs counterexample, not silence")
    for field in ("redundant_with", "conflicts_with"):
        refs = pattern[field]
        if not isinstance(refs, list) or any(not isinstance(ref, str) for ref in refs) or len(refs) != len(set(refs)) or label in refs:
            raise BridgeError(f"{label}: invalid {field}")
    if pattern["redundant_with"] or pattern["conflicts_with"]:
        require_text(pattern["reconcile_note"], label + " reconcile_note", True)
    if pattern["status"] == "drop":
        require_text(pattern["filter_reason"], label + " filter_reason", True)
        return
    require_text(pattern["action"], label + " action", True)
    if pattern["confidence"] == "Low" or n == 0 or pattern["filter_reason"] != "":
        raise BridgeError(f"{label}: kept pattern needs support, High/Medium, empty filter_reason")
    if any("unresolved" in values for values in roles.values()):
        raise BridgeError(f"{label}: unresolved internal conflict cannot be kept")
    if (n == d == 1 or n * 100 < 30 * d) and not pattern["condition"].strip():
        raise BridgeError(f"{label}: singleton/rare pattern requires a reusable source condition")
    if pattern["condition"].strip() and not any("condition" in values or "support" in values for values in roles.values()):
        raise BridgeError(f"{label}: condition needs local evidence")
    if any(a["eligibility"] == "unknown" or (
            a["eligibility"] == "eligible" and a["occurrence"] == "unknown") for a in assessments):
        require_text(pattern.get("limitations", ""), label + " unknown limitations", True)


def validate_results(payload: dict, document: Any) -> list[dict]:
    require_object(document, {"schema_version", "payload_digest", "metadata", "criteria"})
    if document["schema_version"] != VERSION or document["payload_digest"] != payload["payload_digest"]:
        raise BridgeError("Expected review schema 3.0 and current payload digest; prepare and synthesize again")
    metadata = document["metadata"]
    require_object(metadata, {"corpus_scope", "limitations", "reconciliation_completed", "merge_completed"})
    require_text(metadata["corpus_scope"], "corpus_scope", True)
    require_text(metadata["limitations"], "corpus limitations")
    if metadata["reconciliation_completed"] is not True or metadata["merge_completed"] is not True:
        raise BridgeError("Complete semantic RECONCILE and MERGE before export")
    criteria = document["criteria"]
    if not isinstance(criteria, list) or [
            item.get("criterion_id") if isinstance(item, dict) else None for item in criteria] != [
                item["criterion_id"] for item in payload["criteria"]]:
        raise BridgeError("Missing, duplicate, unknown or out-of-order criteria")
    patterns, kept = {}, []
    for item, source in zip(criteria, payload["criteria"]):
        require_object(item, {"criterion_id", "synthesis", "patterns", "dna_result"})
        require_text(item["synthesis"], item["criterion_id"] + " synthesis", True)
        if not isinstance(item["patterns"], list):
            raise BridgeError("patterns must be an array")
        for pattern in item["patterns"]:
            validate_pattern(pattern, source, payload["parts"])
            if pattern["pattern_id"] in patterns:
                raise BridgeError("Duplicate pattern_id")
            patterns[pattern["pattern_id"]] = pattern
        survivors = [p for p in item["patterns"] if p["status"] == "keep"]
        merged = item["dna_result"]
        if not survivors:
            if merged is not None:
                raise BridgeError(f"{item['criterion_id']}: no kept pattern; dna_result must be null")
            continue
        require_object(merged, MERGED_FIELDS, MERGED_OPTIONAL)
        if merged["pattern_ids"] != [p["pattern_id"] for p in survivors]:
            raise BridgeError(f"{item['criterion_id']}: merge must cover exactly all kept patterns in order")
        for field in {"action"} | (MERGED_OPTIONAL & set(merged)):
            require_text(merged[field], item["criterion_id"] + " merged " + field, field == "action")
        for field in ("selection", "validation"):
            if merged.get(field, "").strip() and not any(p.get(field, "").strip() for p in survivors):
                raise BridgeError(f"{item['criterion_id']}: merged {field} needs kept-pattern basis")
        has_condition = any(p["condition"].strip() for p in survivors)
        if merged.get("when", "").strip() and not has_condition:
            raise BridgeError(f"{item['criterion_id']}: merged WHEN needs a kept-pattern condition; no scope filler")
        if has_condition and not (merged.get("when", "").strip() or merged.get("condition_variation", "").strip()):
            raise BridgeError(f"{item['criterion_id']}: merge must retain kept-pattern conditions")
        kept.append(item)
    # References detect relationships only; never combine evidence/frequency.
    for pattern in patterns.values():
        for field in ("redundant_with", "conflicts_with"):
            if any(ref not in patterns for ref in pattern[field]):
                raise BridgeError(f"{pattern['pattern_id']}: unknown reconciliation reference")
        if pattern["status"] == "keep":
            for ref in pattern["conflicts_with"]:
                other = patterns[ref]
                if other["status"] == "keep" and (
                        not pattern["condition"].strip() or not other["condition"].strip() or not other["reconcile_note"].strip()):
                    raise BridgeError("Conflicting kept patterns need evidenced conditions and reconciliation notes")
    active, done = set(), set()

    def visit(pattern_id: str) -> None:
        if pattern_id in active:
            raise BridgeError("Cyclic redundant_with references")
        if pattern_id in done:
            return
        active.add(pattern_id)
        for ref in patterns[pattern_id]["redundant_with"]:
            visit(ref)
        active.remove(pattern_id)
        done.add(pattern_id)

    for pattern_id in patterns:
        visit(pattern_id)
    return kept


def output_rows(payload: dict, criteria: list[dict]) -> list[list]:
    names = {item["criterion_id"]: item["criterion"] for item in payload["criteria"]}
    rows = [OUTPUT_HEADERS.copy()]
    for item in criteria:
        merged = item["dna_result"]
        sections = []
        for field, label in (("when", "WHEN"), ("action", "ACTION"),
                             ("selection", "SELECTION"), ("validation", "VALIDATION")):
            if merged.get(field, "").strip():
                sections.append(label + "\n" + merged[field])
        row = [item["criterion_id"], names[item["criterion_id"]],
               "\n\n".join(sections), merged.get("condition_variation") or None]
        if any(isinstance(value, str) and (
                len(value) > 32767 or ILLEGAL_CHARACTERS_RE.search(value)) for value in row):
            raise BridgeError(f"{item['criterion_id']}: output exceeds Excel cell limits or contains illegal characters")
        rows.append(row)
    return rows


def sheet_snapshot(sheet) -> tuple:
    """Verify ordinary worksheet content/layout on unrelated sheets after save."""
    if not hasattr(sheet, "iter_rows"):
        return (sheet.title, sheet.sheet_state, type(sheet).__name__)
    cells = tuple((cell.coordinate, cell.value, cell.data_type, tuple(cell._style or ()),
                   cell.number_format,
                   (cell.comment.text, cell.comment.author) if cell.comment else None,
                   (cell.hyperlink.target, cell.hyperlink.location) if cell.hyperlink else None)
                  for row in sheet.iter_rows() for cell in row if cell.value is not None or cell.has_style or cell.comment or cell.hyperlink)
    columns = tuple((key, d.width, d.hidden, d.min, d.max, d.outlineLevel)
                    for key, d in sheet.column_dimensions.items())
    rows = tuple((key, d.height, d.hidden, d.outlineLevel)
                 for key, d in sheet.row_dimensions.items())
    return (sheet.title, sheet.sheet_state, cells, tuple(str(r) for r in sheet.merged_cells.ranges),
            sheet.freeze_panes, sheet.auto_filter.ref, columns, rows,
            len(sheet._charts), len(sheet._images))


def file_hash(path: Path) -> str | None:
    return hashlib.sha256(path.read_bytes()).hexdigest() if path.exists() else None


def finalize(paths: Paths) -> int:
    payload = load_json(paths.grouped_payload)
    if payload != extract(paths):
        raise BridgeError("Prepared payload is stale or modified; prepare and synthesize again")
    kept = validate_results(payload, load_json(paths.dna_results))
    rows = output_rows(payload, kept)
    paths.output.parent.mkdir(parents=True, exist_ok=True)
    original_hash = file_hash(paths.output)
    workbook = load_workbook(paths.output, data_only=False) if original_hash is not None else Workbook()
    fd, name = tempfile.mkstemp(prefix=".AUTHOR_DNA.", suffix=".xlsx", dir=paths.output.parent)
    os.close(fd)
    temporary = Path(name)
    try:
        if original_hash is None:
            workbook.remove(workbook.active)
        active_title = workbook.active.title if workbook.active is not None else RESULT_SHEET
        other_sheets = {s.title: sheet_snapshot(s) for s in workbook if s.title != RESULT_SHEET}
        if RESULT_SHEET in workbook.sheetnames:
            position = workbook.sheetnames.index(RESULT_SHEET)
            workbook.remove(workbook[RESULT_SHEET])
        else:
            position = 0
        sheet = workbook.create_sheet(RESULT_SHEET, position)
        if sheet.title != RESULT_SHEET:
            raise BridgeError("AUTHOR_DNA sheet name collides with another sheet")
        if active_title in workbook.sheetnames:
            workbook.active = workbook.sheetnames.index(active_title)
        for row in rows:
            sheet.append(row)
            for cell in sheet[sheet.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244062")
        for column, width in zip("ABCD", (12, 36, 95, 75)):
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "C2"
        sheet.auto_filter.ref = sheet.dimensions
        expected_names = workbook.sheetnames.copy()
        workbook.save(temporary)
        check = load_workbook(temporary, data_only=False)
        try:
            if check.sheetnames != expected_names or [
                    list(row) for row in check[RESULT_SHEET].iter_rows(values_only=True)] != rows:
                raise BridgeError("Saved DNA sheet failed four-column round-trip verification")
            if any(sheet_snapshot(check[name]) != snapshot for name, snapshot in other_sheets.items()):
                raise BridgeError("Saved workbook changed an unrelated sheet")
        finally:
            check.close()
        if file_hash(paths.input) != payload["source"]["sha256"]:
            raise BridgeError("Input changed during export; prepare and synthesize again")
        if file_hash(paths.output) != original_hash:
            raise BridgeError("Output changed during export; retry without replacing concurrent edits")
        os.replace(temporary, paths.output)
    finally:
        workbook.close()
        temporary.unlink(missing_ok=True)
    return len(kept)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("prepare", "finalize"))
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--project-root", type=Path, default=Path.cwd())
    args = parser.parse_args(argv)
    try:
        paths = read_config(args.config, args.project_root)
        if args.command == "prepare":
            payload = prepare(paths)
            print(f"Grouped {len(payload['parts'])} Parts x {len(payload['criteria'])} criteria: {paths.grouped_payload}")
            if payload["ignored_sheets"]:
                print(f"Ignored non-Part sheets: {payload['ignored_sheets']}")
        else:
            count = finalize(paths)
            print(f"Replaced AUTHOR_DNA sheet with {count} criterion results: {paths.output}")
            if count == 0:
                print("No DNA patterns had sufficient evidence; AUTHOR_DNA contains headers only.")
        return 0
    except (BridgeError, OSError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
