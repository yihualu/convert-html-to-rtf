# convert-html-to-rtf

Convert HTML content in CSV files to RTF while preserving rich formatting (bold, italic, underline, colors, etc.).

This repository includes two conversion backends:

- **Pandoc backend** (`src/pandoc/html_to_rtf_by_pandoc.py`)  
  Cross-platform, uses `pypandoc` + Pandoc.
- **Microsoft Word backend** (`src/word/html_to_rtf_by_word.py`)  
  Windows-only, uses Word COM automation via `pywin32`.

It also includes a small utility to normalize/fix RTF headers:

- `src/update_rtf_endoding.py`

## Repository structure

```text
src/
  pandoc/
    html_to_rtf_by_pandoc.py
  word/
    html_to_rtf_by_word.py
  update_rtf_endoding.py
```

## Requirements

## 1) Python

- Python 3.8+ recommended

## 2) Pandoc backend

- `pypandoc` Python package
- Pandoc installed and available on PATH

Install:

```bash
pip install pypandoc
```

Pandoc download: https://pandoc.org

## 3) Microsoft Word backend (Windows only)

- Microsoft Word installed
- `pywin32` Python package

Install:

```bash
pip install pywin32
```

## Input and output format

- Input: CSV file with HTML content in a specified column (default: column **3**, 1-based)
- Output: CSV file with generated RTF in a specified column (default: column **4**, 1-based)
- Default output file for both converters: `test_output.csv` in the corresponding script directory

## Usage

## Pandoc converter

```bash
python src/pandoc/html_to_rtf_by_pandoc.py INPUT.csv \
  --output OUTPUT.csv \
  --num-rows 1000 \
  --progress-every 100 \
  --html-col 3 \
  --rtf-col 4 \
  --encoding utf-8
```

Arguments:

- `input_csv` (required): input CSV path
- `--output`: output CSV path
- `--num-rows`: process only first N rows
- `--progress-every`: progress interval (default `100`)
- `--html-col`: HTML column number, 1-based (default `3`)
- `--rtf-col`: RTF output column number, 1-based (default `4`)
- `--encoding`: CSV encoding (default `utf-8`)

## Microsoft Word converter

```bash
python src/word/html_to_rtf_by_word.py INPUT.csv \
  --output OUTPUT.csv \
  --num-rows 1000 \
  --progress-every 100 \
  --threads 4 \
  --html-col 3 \
  --rtf-col 4
```

Arguments:

- `input_csv` (required): input CSV path
- `--output`: output CSV path
- `--num-rows`: process only first N rows
- `--progress-every`: progress interval (default `100`)
- `--threads`: number of concurrent Word instances (default `4`)
- `--html-col`: HTML column number, 1-based (default `3`)
- `--rtf-col`: RTF output column number, 1-based (default `4`)

## RTF header fix utility

`src/update_rtf_endoding.py` is a standalone script that rewrites RTF headers (including `deflang1036`) for CSV column D (index 3).

Edit the `input_file` and `output_file` variables in the script, then run:

```bash
python src/update_rtf_endoding.py
```

## Notes

- Pandoc converter validates Pandoc availability at startup and exits if missing.
- Word converter validates Microsoft Word availability at startup and exits if missing.
- Empty/blank HTML cells are converted to empty strings.
