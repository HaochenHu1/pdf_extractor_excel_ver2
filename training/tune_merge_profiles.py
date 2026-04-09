#!/usr/bin/env python3
"""Offline tuner for merge-detection profiles using regression JSON/JSONL datasets."""

from __future__ import annotations

import argparse
import json
import random
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from eval_merge_quality import load_records


DEFAULT_PROFILE = {
    "min_confidence": 0.5,
    "method_bonus": {
        "empty_neighbor": 0.0,
        "geometry+empty_neighbor": 0.0,
    },
    "min_span": 2,
}


@dataclass
class CandidateResult:
    profile: dict[str, Any]
    true_positive: int
    predicted_total: int
    expected_total: int
    precision: float
    recall: float
    f1: float


def _safe_ratio(numerator: int, denominator: int) -> float:
    if denominator == 0:
        return 0.0
    return numerator / denominator


def _region_key(region: dict[str, Any]) -> tuple[int, int, int, int]:
    return (
        int(region["start_row"]),
        int(region["end_row"]),
        int(region["start_col"]),
        int(region["end_col"]),
    )


def _region_span(region: dict[str, Any]) -> int:
    row_span = int(region["end_row"]) - int(region["start_row"]) + 1
    col_span = int(region["end_col"]) - int(region["start_col"]) + 1
    return max(row_span, col_span)


def _filter_predicted_merges(record: dict[str, Any], profile: dict[str, Any]) -> set[tuple[int, int, int, int]]:
    raw_pred = record.get("predicted_merges") or record.get("pred_merges") or record.get("merges_pred") or record.get("merges") or []
    if not isinstance(raw_pred, list):
        return set()

    min_confidence = float(profile.get("min_confidence", 0.0))
    min_span = int(profile.get("min_span", 1))
    method_bonus = profile.get("method_bonus", {})
    if not isinstance(method_bonus, dict):
        method_bonus = {}

    kept: set[tuple[int, int, int, int]] = set()
    for region in raw_pred:
        if not isinstance(region, dict):
            continue
        try:
            key = _region_key(region)
        except (KeyError, TypeError, ValueError):
            continue

        span = _region_span(region)
        if span < min_span:
            continue

        confidence = float(region.get("confidence", 1.0))
        method = str(region.get("method", ""))
        adjusted = confidence + float(method_bonus.get(method, 0.0))
        if adjusted < min_confidence:
            continue

        kept.add(key)

    return kept


def _expected_merges(record: dict[str, Any]) -> set[tuple[int, int, int, int]]:
    raw_gold = record.get("expected_merges") or record.get("golden_merges") or record.get("gold_merges") or record.get("merges_gold") or []
    if not isinstance(raw_gold, list):
        return set()

    kept: set[tuple[int, int, int, int]] = set()
    for region in raw_gold:
        if not isinstance(region, dict):
            continue
        try:
            kept.add(_region_key(region))
        except (KeyError, TypeError, ValueError):
            continue
    return kept


def evaluate_profile(records: list[dict[str, Any]], profile: dict[str, Any]) -> CandidateResult:
    tp = 0
    pred_total = 0
    gold_total = 0

    for record in records:
        pred = _filter_predicted_merges(record, profile)
        gold = _expected_merges(record)
        tp += len(pred & gold)
        pred_total += len(pred)
        gold_total += len(gold)

    precision = _safe_ratio(tp, pred_total)
    recall = _safe_ratio(tp, gold_total)
    f1 = _safe_ratio(2 * precision * recall, precision + recall) if (precision + recall) else 0.0
    return CandidateResult(profile, tp, pred_total, gold_total, precision, recall, f1)


def grid_candidates() -> list[dict[str, Any]]:
    profiles: list[dict[str, Any]] = []
    for min_confidence in [0.3, 0.4, 0.5, 0.6, 0.7]:
        for min_span in [1, 2, 3]:
            for bonus_empty in [0.0, 0.05]:
                for bonus_geom in [0.0, 0.05, 0.1]:
                    profiles.append(
                        {
                            "min_confidence": min_confidence,
                            "min_span": min_span,
                            "method_bonus": {
                                "empty_neighbor": bonus_empty,
                                "geometry+empty_neighbor": bonus_geom,
                            },
                        }
                    )
    return profiles


def _load_baseline(path: Path | None) -> dict[str, Any]:
    if path is None:
        return dict(DEFAULT_PROFILE)
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise SystemExit(f"Baseline profile must be a JSON object: {path}")
    baseline = dict(DEFAULT_PROFILE)
    baseline.update(payload)
    return baseline


def _print_candidate_table(results: list[CandidateResult], topk: int) -> None:
    print("\nTop candidates")
    print("rank\tprecision\trecall\tf1\tmin_conf\tmin_span\tbonus_empty\tbonus_geom")
    for idx, result in enumerate(results[:topk], start=1):
        bonus = result.profile["method_bonus"]
        print(
            f"{idx}\t{result.precision:.4f}\t{result.recall:.4f}\t{result.f1:.4f}"
            f"\t{result.profile['min_confidence']:.2f}\t{result.profile['min_span']}"
            f"\t{bonus.get('empty_neighbor', 0.0):.2f}\t{bonus.get('geometry+empty_neighbor', 0.0):.2f}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Offline merge profile tuner (no runtime patching).")
    parser.add_argument("--input", required=True, type=Path, help="Evaluation dataset path (.json/.jsonl or directory).")
    parser.add_argument("--baseline-profile", type=Path, default=None, help="Optional baseline profile JSON.")
    parser.add_argument("--precision-floor", type=float, default=0.90, help="Minimum precision constraint.")
    parser.add_argument("--topk", type=int, default=10, help="Number of top candidates to print.")
    parser.add_argument("--out-json", type=Path, default=None, help="Optional output path for best profile JSON.")
    parser.add_argument("--seed", type=int, default=7, help="Random seed for reproducibility.")
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Write suggested profile artifact only (does NOT patch runtime constants).",
    )
    args = parser.parse_args()

    random.seed(args.seed)
    records = load_records(args.input)
    if not records:
        raise SystemExit(f"No records found in input: {args.input}")

    baseline_profile = _load_baseline(args.baseline_profile)
    baseline = evaluate_profile(records, baseline_profile)

    candidates = [evaluate_profile(records, profile) for profile in grid_candidates()]
    feasible = [row for row in candidates if row.precision >= args.precision_floor]
    ranked_pool = feasible if feasible else candidates
    ranked = sorted(
        ranked_pool,
        key=lambda row: (-row.f1, -row.precision, -row.recall, row.profile["min_confidence"], row.profile["min_span"]),
    )
    best = ranked[0]

    print(f"records={len(records)} precision_floor={args.precision_floor:.2f}")
    print(
        f"best precision={best.precision:.4f} recall={best.recall:.4f} f1={best.f1:.4f} "
        f"min_conf={best.profile['min_confidence']:.2f} min_span={best.profile['min_span']}"
    )
    _print_candidate_table(ranked, max(1, args.topk))

    print("\nBaseline comparison")
    print(f"baseline precision={baseline.precision:.4f} recall={baseline.recall:.4f} f1={baseline.f1:.4f}")
    print(
        "delta "
        f"precision={best.precision - baseline.precision:+.4f} "
        f"recall={best.recall - baseline.recall:+.4f} "
        f"f1={best.f1 - baseline.f1:+.4f}"
    )

    best_payload = {
        "profile": best.profile,
        "metrics": {
            "precision": round(best.precision, 6),
            "recall": round(best.recall, 6),
            "f1": round(best.f1, 6),
        },
        "baseline_metrics": {
            "precision": round(baseline.precision, 6),
            "recall": round(baseline.recall, 6),
            "f1": round(baseline.f1, 6),
        },
    }
    print("\nBest profile JSON")
    print(json.dumps(best_payload, ensure_ascii=False, indent=2, sort_keys=True))

    out_json = args.out_json
    if args.apply and out_json is None:
        out_json = Path("training/suggested_merge_profile.json")

    if out_json is not None:
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(best_payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(f"\nWrote suggested profile artifact: {out_json}")
        if args.apply:
            print("Apply mode only writes artifact; runtime constants were not modified.")


if __name__ == "__main__":
    main()
