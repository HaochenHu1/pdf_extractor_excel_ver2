# PDF Table Extractor to Excel

A (small) command-line Python tool that extracts tabular data from PDF documents and exports results to a structured Excel workbook.

This project is designed for practical workflows where PDFs can be either:
- **Text-based** (machine-readable), or
- **Scanned/image-based** (OCR required).

---

## Features

- Extracts tables from a PDF into a single `.xlsx` file.
- Supports multiple extraction backends:
  - `camelot` (strong for structured, text-based PDFs)
  - `pdfplumber` (reliable fallback for text PDFs)
  - `img2table` (useful for scanned/OCR workflows)
- Automatic backend selection with manual override.
- Page-range targeting (`all`, single pages, lists, and ranges).
- Quality filters for extracted tables (row/column thresholds, fill ratio, accuracy).
- Optional OCR tuning for scanned Chinese documents.
- Adds a `_summary` worksheet with extraction metadata.

---

## Project Structure

- `pdf_table_extractor.py` — main CLI script
- `requirements.txt` — Python dependencies
- `training/` — sample/training PDFs

---

## Installation

### 1) Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2) Optional: OCR support for scanned PDFs

```bash
pip install img2table
```

If you plan to use OCR via `img2table`, install **Tesseract OCR** on your system.

Example (Windows):

```bash
winget install --id UB-Mannheim.TesseractOCR
```

---

## Usage

### Basic (beginner mode)

```bash
python pdf_table_extractor.py input.pdf
```

Default output:

```text
input_tables.xlsx
```

You can also pass a folder to process all PDFs inside it:

```bash
python pdf_table_extractor.py ./pdfs
```

Default batch output folder:

```text
./pdfs/extracted_tables/
```

### Specify output file (single PDF)

```bash
python pdf_table_extractor.py input.pdf -o output.xlsx
```

### Specify output folder (single PDF or batch)

```bash
python pdf_table_extractor.py input.pdf --output-dir ./out
python pdf_table_extractor.py ./pdfs --output-dir ./out
```

### Process selected pages

```bash
python pdf_table_extractor.py input.pdf --pages 1-3,5
```

### Select extraction backend

```bash
python pdf_table_extractor.py input.pdf --mode auto
python pdf_table_extractor.py input.pdf --mode camelot
python pdf_table_extractor.py input.pdf --mode pdfplumber
python pdf_table_extractor.py input.pdf --mode img2table
```

### Verbose mode

```bash
python pdf_table_extractor.py input.pdf --verbose
```

---

## CLI Options (Reference)

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

---

## Output Format

The generated Excel workbook contains:

- `Table_001`, `Table_002`, ... — one worksheet per extracted table
- `_summary` — page number, extraction engine, score/quality indicators, and table shape metadata

### Stable merge reject reasons in `_summary`

- `_summary.merge_reject_top_reasons` now stores sanitized machine-friendly tokens only.
- Tokens are stripped of control/newline characters, normalized to `[a-z0-9_]`, capped to 40 chars, and limited to top 5 distinct reasons.
- Ordering is deterministic: count descending, then label ascending.

---

## Practical Notes

- For text-based PDFs, `camelot` is usually the best first choice.
- For scanned PDFs, `img2table` + Tesseract OCR is recommended.
- If extraction quality is low, try:
  - restricting pages with `--pages`
  - forcing a different backend with `--mode`
  - adjusting thresholds (`--min-filled-ratio`, `--accuracy-threshold`)
  - lowering `--img2table-min-confidence` for noisy scans

Example (multilingual OCR):

```bash
python pdf_table_extractor.py financial_report.pdf -o financial_tables.xlsx --ocr-lang "chi_sim+eng"
```

---

## Known Limitation

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
