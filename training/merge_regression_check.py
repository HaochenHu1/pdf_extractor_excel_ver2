from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path
import sys
from types import SimpleNamespace

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from pdf_table_extractor import detect_pdf_kind, extract_tables_for_pdf, infer_merged_regions


def build_args(mode: str, verbose: bool) -> SimpleNamespace:
    return SimpleNamespace(
        pages="all",
        mode=mode,
        prefer="both",
        accuracy_threshold=50.0,
        min_rows=2,
        min_cols=2,
        min_filled_ratio=0.15,
        ocr_lang="eng",
        ocr_lang_auto=False,
        borderless=False,
        img2table_min_confidence=50,
        verbose=verbose,
    )


def ensure_text_fixture(fixtures_dir: Path) -> Path:
    fixture_path = fixtures_dir / "fixture_text_table.pdf"
    if fixture_path.exists():
        return fixture_path

    import fitz  # local dependency already used by pipeline

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    lines = [
        "Group A,,Group B,",
        "Q1,Q2,Q3,Q4",
        "10,11,12,13",
        "14,15,16,17",
    ]
    y = 72
    for line in lines:
        page.insert_text((72, y), line, fontsize=11)
        y += 18
    doc.save(fixture_path)
    doc.close()
    return fixture_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Run local merge regression checks on fixture PDFs.")
    parser.add_argument("--fixtures-dir", type=Path, default=Path("training"))
    parser.add_argument("--mode", choices=["auto", "camelot", "pdfplumber", "img2table"], default="auto")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    fixtures = sorted(p for p in args.fixtures_dir.glob("*.pdf"))
    if not fixtures:
        print(f"No PDF fixtures found in {args.fixtures_dir}")
        return 1

    generated_text_pdf = ensure_text_fixture(args.fixtures_dir)
    text_pdf = next((p for p in [generated_text_pdf] + fixtures if detect_pdf_kind(p) == "text"), None)
    scanned_pdf = next((p for p in fixtures if detect_pdf_kind(p) == "scanned"), None)
    selected = [p for p in [text_pdf, scanned_pdf] if p is not None]
    if not selected:
        selected = fixtures[:2]

    print("fixture,kind,tables,total_merges,header_merges,body_merges,rejects,reject_reasons")
    for pdf in selected:
        kind = detect_pdf_kind(pdf)
        extract_args = build_args(args.mode, args.verbose)
        tables = extract_tables_for_pdf(pdf, extract_args)
        total_merges = 0
        header_merges = 0
        body_merges = 0
        rejects = 0
        reason_counter: Counter[str] = Counter()
        for table in tables:
            merge_plan = infer_merged_regions(table)
            total_merges += len(merge_plan["accepted"])
            header_merges += len(merge_plan["header_spans"])
            body_merges += len(merge_plan["body_spans"])
            rejects += len(merge_plan["rejected"])
            for item in merge_plan["rejected"]:
                for reason in item.get("reason_tags", [])[:3]:
                    reason_counter[str(reason)] += 1
        top_reasons = ";".join(f"{r}:{c}" for r, c in reason_counter.most_common(3))
        print(
            f"{pdf.name},{kind},{len(tables)},{total_merges},{header_merges},{body_merges},{rejects},{top_reasons}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
