# PDF Table Extractor to Excel

A command-line Python tool that extracts tabular data from PDF documents and exports results to a structured Excel workbook.

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

### Basic

```bash
python pdf_table_extractor.py input.pdf
```

Default output:

```text
input_tables.xlsx
```

### Specify output file

```bash
python pdf_table_extractor.py input.pdf -o output.xlsx
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
--verbose
```

---

## Output Format

The generated Excel workbook contains:

- `Table_001`, `Table_002`, ... — one worksheet per extracted table
- `_summary` — page number, extraction engine, score/quality indicators, and table shape metadata

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
