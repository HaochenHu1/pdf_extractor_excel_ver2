#!/usr/bin/env python3
"""Evaluate merge-region quality from regression outputs.

This script computes micro-averaged precision/recall/F1 for merged-cell predictions
using exact region matching.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


@dataclass(frozen=True)
class MergeRegion:
    start_row: int
    end_row: int
    start_col: int
    end_col: int

    @classmethod
    def from_obj(cls, obj: dict[str, Any]) -> "MergeRegion":
        return cls(
            start_row=int(obj["start_row"]),
            end_row=int(obj["end_row"]),
            start_col=int(obj["start_col"]),
            end_col=int(obj["end_col"]),
        )


def _pick_first(record: dict[str, Any], keys: Iterable[str], default: Any) -> Any:
    for key in keys:
        if key in record:
            return record[key]
    return default


def _record_id(record: dict[str, Any], index: int) -> str:
    doc = str(_pick_first(record, ["doc_id", "pdf", "file", "document"], "unknown_doc"))
    page = _pick_first(record, ["page", "page_num"], "?")
    table = _pick_first(record, ["table_id", "table", "sheet_name"], index)
    return f"{doc}|page={page}|table={table}"


def _load_json_file(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if isinstance(payload, dict):
        if "records" in payload and isinstance(payload["records"], list):
            return [item for item in payload["records"] if isinstance(item, dict)]
        return [payload]
    return []


def _load_jsonl_file(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line:
            continue
        obj = json.loads(line)
        if isinstance(obj, dict):
            rows.append(obj)
    return rows


def load_records(path: Path) -> list[dict[str, Any]]:
    if path.is_file():
        if path.suffix.lower() == ".jsonl":
            return _load_jsonl_file(path)
        return _load_json_file(path)

    records: list[dict[str, Any]] = []
    for file_path in sorted(path.glob("*.json")):
        records.extend(_load_json_file(file_path))
    for file_path in sorted(path.glob("*.jsonl")):
        records.extend(_load_jsonl_file(file_path))
    return records


def _to_regions(raw_regions: Any) -> set[MergeRegion]:
    if not isinstance(raw_regions, list):
        return set()

    regions: set[MergeRegion] = set()
    for item in raw_regions:
        if not isinstance(item, dict):
            continue
        try:
            regions.add(MergeRegion.from_obj(item))
        except (KeyError, TypeError, ValueError):
            continue
    return regions


def evaluate(records: list[dict[str, Any]], verbose: bool = False) -> tuple[int, int, int]:
    true_positive = 0
    predicted_total = 0
    expected_total = 0

    for idx, record in enumerate(records, start=1):
        predicted_regions = _to_regions(
            _pick_first(record, ["predicted_merges", "pred_merges", "merges_pred", "merges"], [])
        )
        expected_regions = _to_regions(
            _pick_first(record, ["expected_merges", "golden_merges", "gold_merges", "merges_gold"], [])
        )

        tp = len(predicted_regions & expected_regions)
        true_positive += tp
        predicted_total += len(predicted_regions)
        expected_total += len(expected_regions)

        if verbose:
            rid = _record_id(record, idx)
            print(
                f"[{rid}] tp={tp} pred={len(predicted_regions)} gold={len(expected_regions)}"
            )

    return true_positive, predicted_total, expected_total


def ratio(numerator: int, denominator: int) -> float:
    if denominator == 0:
        return 0.0
    return numerator / denominator


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compute precision/recall/F1 for merged-cell predictions from regression outputs."
    )
    parser.add_argument(
        "input_path",
        type=Path,
        help="JSON/JSONL file or directory containing regression outputs.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print per-record counts.",
    )
    args = parser.parse_args()

    records = load_records(args.input_path)
    if not records:
        raise SystemExit(f"No evaluation records found in: {args.input_path}")

    tp, pred_total, gold_total = evaluate(records, verbose=args.verbose)
    precision = ratio(tp, pred_total)
    recall = ratio(tp, gold_total)
    f1 = ratio(2 * precision * recall, precision + recall) if (precision + recall) else 0.0

    print(f"records={len(records)}")
    print(f"true_positive={tp}")
    print(f"predicted_total={pred_total}")
    print(f"gold_total={gold_total}")
    print(f"precision={precision:.4f}")
    print(f"recall={recall:.4f}")
    print(f"f1={f1:.4f}")


if __name__ == "__main__":
    main()
