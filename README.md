#PDF Table Extractor to Excel

This package contains a local Python tool that takes a PDF as input and writes only the detected tables into an Excel workbook.

##What it does

- Reads a PDF file
- Detects whether the PDF is mainly text based or scanned
- Tries table extraction with free Python libraries
- Writes each extracted table to its own Excel sheet
- Adds a `_summary` sheet with page number, extraction engine, and basic metadata

##Why this design

Existing free tools already cover most of the hard work:

- `Camelot` is strong for text based PDFs with clear table structure
- `pdfplumber` is a good fallback when Camelot misses tables
- `img2table` is useful for scanned PDFs and OCR based workflows

So this script does not reimplement table recognition from scratch. It wraps the best free options into one command line tool.

##Files

- `pdf_table_extractor.py` : main script
- `requirements.txt` : Python dependencies
#======
#Step 1
#======
##Install

###Base install

```
pip install -r requirements.txt
```

###Optional OCR support for scanned PDFs

```
pip install img2table
```

If you want OCR with Tesseract, install Tesseract on your machine as well.

```
winget install --id UB-Mannheim.TesseractOCR
```


#======
#Step 2
#======

##Usage

###Basic

```
python pdf_table_extractor.py input.pdf
```

This writes:

```
input_tables.xlsx
```

###Choose output path

```
python pdf_table_extractor.py input.pdf -o output.xlsx
```

###Process only some pages

```
python pdf_table_extractor.py input.pdf --pages 1-3,5
```

###Force a specific tool

```
python pdf_table_extractor.py input.pdf --mode camelot
python pdf_table_extractor.py input.pdf --mode pdfplumber
python pdf_table_extractor.py input.pdf --mode img2table
```

###Verbose logging

```
python pdf_table_extractor.py input.pdf --verbose
```

##Notes

- For text based PDFs, `Camelot` usually performs best.
- For scanned PDFs, install `img2table` and OCR support.
- If a table is not detected well, try changing pages, forcing a backend, or lowering thresholds.

##Useful options

```
--ocr-lang eng
```

##Output structure

The Excel workbook contains:

- one sheet per extracted table, named `Table_001`, `Table_002`, etc.
- one `_summary` sheet listing page, engine, score, rows, and columns

##Example use

```
python pdf_table_extractor.py financial_report.pdf -o financial_tables.xlsx --ocr-lang "chi_sim+eng"
```

##Issues need to be fixed

Chinese scanned pdfs do not have an accurate output
