<h1 align="center">PDF Table Extractor → Excel</h1>

<p align="center">
A production-minded CLI for extracting tabular data from PDF files and exporting clean, auditable Excel workbooks.
</p>



## Why this project exists

Most PDF table extraction tools fail in one of two ways: they work well on ideal text PDFs but collapse on scanned documents, or they recover data from scans but return low-quality structure.

This project takes a pragmatic approach:

- Use multiple extraction engines and pick the most useful result.
- Keep filtering and quality gates explicit.
- Emit a summary sheet so every output is inspectable.
- Support both single-file and batch workflows with consistent CLI behavior.



## Core capabilities

- Extract tables from **text PDFs** and **scanned/image PDFs**.
- Unified CLI with backend selection:
  - `auto` (default strategy)
  - `camelot`
  - `pdfplumber`
  - `img2table` (OCR-centric)
- Page targeting with flexible syntax (`all`, single pages, lists, ranges).
- Table quality controls:
  - minimum rows/columns
  - non-empty cell ratio
  - Camelot accuracy threshold
- OCR controls for scan-heavy or multilingual documents.
- Per-table Excel sheets plus a `_summary` worksheet with extraction metadata.
- Optional Excel styling policy (`--excel-style-mode basic|off`).



## Repository layout

- `pdf_table_extractor.py` — main CLI entry point
- `requirements.txt` — Python dependencies
- `training/` — evaluation/tuning helpers and sample assets



## Installation

### 1) Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2) Install OCR dependencies (needed for `img2table` workflows)

```bash
pip install img2table
```

Install **Tesseract OCR** at system level.

Windows example:

```bash
winget install --id UB-Mannheim.TesseractOCR
```

> If you do not use `img2table`, Tesseract is optional.



## Quick start

### Single PDF (default output naming)

```bash
python pdf_table_extractor.py input.pdf
```

Output:

```text
input_tables.xlsx
```

### Folder/batch mode

```bash
python pdf_table_extractor.py ./pdfs
```

Default batch output directory:

```text
./pdfs/extracted_tables/
```

### Explicit output targets

```bash
# Single file output path
python pdf_table_extractor.py input.pdf -o output.xlsx

# Output directory (single or batch mode)
python pdf_table_extractor.py input.pdf --output-dir ./out
python pdf_table_extractor.py ./pdfs --output-dir ./out
```

### Target selected pages only

```bash
python pdf_table_extractor.py input.pdf --pages 1-3,5
```

### Force extraction backend

```bash
python pdf_table_extractor.py input.pdf --mode auto
python pdf_table_extractor.py input.pdf --mode camelot
python pdf_table_extractor.py input.pdf --mode pdfplumber
python pdf_table_extractor.py input.pdf --mode img2table
```

### Verbose run

```bash
python pdf_table_extractor.py input.pdf --verbose
```



## CLI reference

```text
-h, --help
-o, --output
--output-dir
--pages
--mode {auto,camelot,pdfplumber,img2table}
--prefer {stream,lattice,both}
--min-rows
--min-cols
--min-filled-ratio
--accuracy-threshold
--ocr-lang
--ocr-lang-auto
--borderless
--img2table-min-confidence
--excel-style-mode {basic,off}
--verbose
```

Notable options:

- `--prefer {stream,lattice,both}`: Camelot extraction strategy preference.
- `--ocr-lang` / `--ocr-lang-auto`: OCR language tuning for scanned docs.
- `--img2table-min-confidence`: Lower for noisy scans if needed.
- `--excel-style-mode {basic,off}`: toggle post-processing styling behavior.



## Output contract

Generated workbook contents:

- `Table_001`, `Table_002`, ... : one sheet per accepted table.
- `_summary`: metadata for traceability (source page, engine, quality signals, shape).

### Deterministic merge-reject reason tokens

`_summary.merge_reject_top_reasons` is normalized for machine consumption:

### Stable merge reject reasons in `_summary`

- `_summary.merge_reject_top_reasons` now stores sanitized machine-friendly tokens only.
- Tokens are stripped of control/newline characters, normalized to `[a-z0-9_]`, capped to 40 chars, and limited to top 5 distinct reasons.
- Ordering is deterministic: count descending, then label ascending.

---



## Practical tuning guidance

1. **Text-native PDFs**: start with `--mode camelot`.
2. **Scanned PDFs**: try `--mode img2table` + proper OCR language.
3. If quality is weak:
   - narrow scope via `--pages`
   - switch backend (`--mode`)
   - relax/tighten filters (`--min-filled-ratio`, `--accuracy-threshold`)
   - lower `--img2table-min-confidence` on noisy scans

Example (Chinese + English OCR):

```bash
python pdf_table_extractor.py financial_report.pdf -o financial_tables.xlsx --ocr-lang "chi_sim+eng"
```



- Accuracy on heavily degraded Chinese scanned PDFs can still vary depending on image quality and OCR configuration.

---

## Phase 3: Golden Merge Expectations (JSON)

For merge-quality regression checks, create JSON/JSONL records with these fields:

```json
{
  "doc_id": "training1.pdf",
  "page": 1,
  "table_id": "Table_001",
  "predicted_merges": [
    {"start_row": 0, "end_row": 0, "start_col": 0, "end_col": 2}
  ],
  "expected_merges": [
    {"start_row": 0, "end_row": 0, "start_col": 0, "end_col": 2}
  ]
}
```

Minimal schema notes:

- `predicted_merges` and `expected_merges` are arrays of merge-region objects.
- Each merge region must include integer fields:
  - `start_row`, `end_row`, `start_col`, `end_col`
- Matching is **exact region equality** on those 4 coordinates.

Evaluate with:

```bash
python training/eval_merge_quality.py <path_to_json_or_jsonl_or_directory>
```

### Profile tuning (offline)

Use the offline tuner to search small merge-profile grids against your regression set:

```bash
python training/tune_merge_profiles.py --input <path_to_json_or_jsonl_or_directory> --topk 10 --precision-floor 0.90
```

`tune_merge_profiles.py` does **not** modify runtime constants by default. Even with `--apply`, it only writes a suggested profile artifact file.

---

## Phase 3 quality gate (quick check)

Use this reproducible command sequence before sharing results:

```bash
python -m py_compile pdf_table_extractor.py training/eval_merge_quality.py training/tune_merge_profiles.py
python training/eval_merge_quality.py /tmp/merge_eval_sample.json
python training/tune_merge_profiles.py --help | head -n 40
```

Interpretation:

- **PASS**: commands run successfully and produce metrics/help output.
- **WARN**: command runs but reports empty/missing input data (for example, no records found); fix the dataset path and rerun.
