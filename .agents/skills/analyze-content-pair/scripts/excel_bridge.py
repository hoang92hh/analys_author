#!/usr/bin/env python3
"""Deterministic Excel I/O bridge for analyze-content-pair."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence

from openpyxl import load_workbook

SCHEMA_VERSION = "1.0"
REQUIRED_CONFIG = (
    "content_a", "content_b", "criteria", "criteria_payload",
    "analysis_results", "output", "result_column",
)


class BridgeError(RuntimeError):
    pass


@dataclass(frozen=True)
class Paths:
    project_root: Path
    content_a: Path
    content_b: Path
    criteria: Path
    criteria_payload: Path
    analysis_results: Path
    output: Path
    result_column: str


def read_config(config_path: Path, project_root: Path | None = None) -> Paths:
    values: Dict[str, str] = {}
    for number, raw in enumerate(
        config_path.resolve().read_text(encoding="utf-8-sig").splitlines(), 1
    ):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in raw:
            raise BridgeError(f"Config line {number} is not key: value")
        key, value = raw.split(":", 1)
        key, value = key.strip(), value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        if not key or key in values:
            raise BridgeError(f"Invalid or duplicate config key at line {number}")
        values[key] = value
    missing = [key for key in REQUIRED_CONFIG if not values.get(key)]
    if missing:
        raise BridgeError("Missing config values: " + ", ".join(missing))

    root = (project_root or Path.cwd()).resolve()

    def resolve(value: str) -> Path:
        path = Path(value)
        return (path if path.is_absolute() else root / path).resolve()

    return Paths(
        root, resolve(values["content_a"]), resolve(values["content_b"]),
        resolve(values["criteria"]), resolve(values["criteria_payload"]),
        resolve(values["analysis_results"]), resolve(values["output"]),
        values["result_column"],
    )


def _require_files(paths: Iterable[Path]) -> None:
    missing = [str(path) for path in paths if not path.is_file()]
    if missing:
        raise BridgeError("Missing input file(s): " + ", ".join(missing))


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _headers(sheet: Any) -> Dict[str, int]:
    headers: Dict[str, int] = {}
    for cell in sheet[1]:
        if cell.value is None:
            continue
        name = str(cell.value).strip()
        if name in headers:
            raise BridgeError(f"Duplicate header {name!r} in sheet {sheet.title!r}")
        headers[name] = cell.column
    return headers


def _relative(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return str(path)


def extract_criteria(paths: Paths) -> Dict[str, Any]:
    _require_files((paths.content_a, paths.content_b, paths.criteria))
    # Normal mode releases the ZIP handle reliably on Windows after close().
    # The prepare path never calls save(), so the source remains read-only in practice.
    workbook = load_workbook(paths.criteria, read_only=False, data_only=False)
    criteria: List[Dict[str, Any]] = []
    seen_ids = set()
    criteria_sheets = 0
    required = ("ID", "Nhóm", "Tiêu chí", "requirement")
    try:
        for sheet in workbook.worksheets:
            headers = _headers(sheet)
            present = [name in headers for name in required]
            if any(present) and not all(present):
                absent = [name for name in required if name not in headers]
                raise BridgeError(
                    f"Sheet {sheet.title!r} has partial criteria schema; missing: {', '.join(absent)}"
                )
            if not all(present):
                continue
            criteria_sheets += 1
            if paths.result_column not in headers:
                raise BridgeError(
                    f"Sheet {sheet.title!r} is missing result column {paths.result_column!r}"
                )
            for row in range(2, sheet.max_row + 1):
                raw_id = sheet.cell(row, headers["ID"]).value
                if raw_id is None or str(raw_id).strip() == "":
                    if any(sheet.cell(row, headers[name]).value not in (None, "") for name in required[1:]):
                        raise BridgeError(f"Criterion row {sheet.title}!{row} has no ID")
                    continue
                criterion_id = str(raw_id).strip()
                if criterion_id in seen_ids:
                    raise BridgeError(f"Duplicate criterion ID: {criterion_id}")
                seen_ids.add(criterion_id)
                criteria.append({
                    "criterion_key": f"{sheet.title}!{row}:{criterion_id}",
                    "criterion_id": criterion_id,
                    "sheet": sheet.title,
                    "row": row,
                    "group": sheet.cell(row, headers["Nhóm"]).value,
                    "criterion": sheet.cell(row, headers["Tiêu chí"]).value,
                    "requirement": sheet.cell(row, headers["requirement"]).value,
                })
    finally:
        workbook.close()
    if criteria_sheets != 1:
        raise BridgeError(f"Expected exactly one criteria sheet, found {criteria_sheets}")
    if not criteria:
        raise BridgeError("No criteria rows found")
    return {
        "schema_version": SCHEMA_VERSION,
        "source": {
            "content_a": _relative(paths.content_a, paths.project_root),
            "content_b": _relative(paths.content_b, paths.project_root),
            "criteria": _relative(paths.criteria, paths.project_root),
            "criteria_sha256": _sha256(paths.criteria),
            "result_column": paths.result_column,
        },
        "criteria": criteria,
    }


def _write_json_atomic(path: Path, data: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(name, path)
    except Exception:
        try:
            os.unlink(name)
        except FileNotFoundError:
            pass
        raise


def prepare(paths: Paths) -> Dict[str, Any]:
    payload = extract_criteria(paths)
    _write_json_atomic(paths.criteria_payload, payload)
    return payload


def _load_json(path: Path) -> Any:
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            return json.load(handle)
    except json.JSONDecodeError as error:
        raise BridgeError(f"Invalid JSON in {path}: {error}") from error


def validate_results(payload: Mapping[str, Any], document: Any) -> Dict[str, str]:
    if not isinstance(document, dict) or set(document) != {"schema_version", "results"}:
        raise BridgeError("Result document must contain only schema_version and results")
    if document["schema_version"] != SCHEMA_VERSION:
        raise BridgeError(f"Unsupported schema_version: {document['schema_version']!r}")
    if not isinstance(document["results"], list):
        raise BridgeError("results must be an array")
    expected = {item["criterion_key"]: item for item in payload["criteria"]}
    mapped: Dict[str, str] = {}
    errors: List[str] = []
    for index, item in enumerate(document["results"]):
        label = f"results[{index}]"
        if not isinstance(item, dict) or set(item) != {"criterion_key", "criterion_id", "result"}:
            errors.append(f"{label} has invalid fields")
            continue
        key = item["criterion_key"]
        if not isinstance(key, str) or not key:
            errors.append(f"{label}.criterion_key is invalid")
        elif key in mapped:
            errors.append(f"Duplicate result: {key}")
        elif key not in expected:
            errors.append(f"Unknown criterion_key: {key}")
        elif item["criterion_id"] != expected[key]["criterion_id"]:
            errors.append(f"criterion_id does not match {key}")
        elif not isinstance(item["result"], str):
            errors.append(f"Invalid result: {key}")
        else:
            mapped[key] = item["result"].strip()
    missing = [key for key in expected if key not in mapped]
    if missing:
        errors.append("Missing results: " + ", ".join(missing))
    if errors:
        raise BridgeError("; ".join(errors))
    return mapped


def finalize(paths: Paths) -> int:
    _require_files((paths.content_a, paths.content_b, paths.criteria, paths.criteria_payload, paths.analysis_results))
    payload = _load_json(paths.criteria_payload)
    if payload != extract_criteria(paths):
        raise BridgeError("criteria_payload is stale; run prepare again")
    mapped = validate_results(payload, _load_json(paths.analysis_results))
    if paths.output == paths.criteria:
        raise BridgeError("Output must not overwrite criteria workbook")
    paths.output.parent.mkdir(parents=True, exist_ok=True)

    selected = [
        item for item in payload["criteria"]
        if mapped[item["criterion_key"]]
    ]

    fd, name = tempfile.mkstemp(prefix=f".{paths.output.stem}.", suffix=".xlsx", dir=paths.output.parent)
    os.close(fd)
    temporary = Path(name)
    try:
        from openpyxl import Workbook

        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = payload["criteria"][0]["sheet"]
            sheet.append(["ID", "Nhóm", "Tiêu chí", "Requirement", paths.result_column])
            for item in selected:
                sheet.append([
                    item["criterion_id"], item["group"], item["criterion"],
                    item["requirement"], mapped[item["criterion_key"]],
                ])
            workbook.save(temporary)
        finally:
            workbook.close()

        output = load_workbook(temporary, data_only=False)
        try:
            if output.sheetnames != [payload["criteria"][0]["sheet"]]:
                raise BridgeError("Output verification failed: unexpected sheets")
            sheet = output.active
            expected_rows = [
                ["ID", "Nhóm", "Tiêu chí", "Requirement", paths.result_column],
                *[
                    [item["criterion_id"], item["group"], item["criterion"],
                     item["requirement"], mapped[item["criterion_key"]]]
                    for item in selected
                ],
            ]
            actual_rows = [list(row) for row in sheet.iter_rows(values_only=True)]
            if actual_rows != expected_rows:
                raise BridgeError("Output verification failed: data mismatch")
        finally:
            output.close()
        os.replace(temporary, paths.output)
    except Exception:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass
        raise
    return len(selected)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("prepare", "finalize"))
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--project-root", type=Path, default=Path.cwd())
    args = parser.parse_args(argv)
    try:
        paths = read_config(args.config, args.project_root)
        if args.command == "prepare":
            count = len(prepare(paths)["criteria"])
            print(f"Prepared {count} criteria: {paths.criteria_payload}")
        else:
            count = finalize(paths)
            print(f"Wrote {count} results: {paths.output}")
        return 0
    except (BridgeError, OSError, KeyError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
