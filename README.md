# Data Cleaner

A small Python project that cleans a messy Indonesian customer dataset with [pandas](https://pandas.pydata.org/), exports the result to a styled Excel workbook, and generates a PDF summary report with [fpdf2](https://py-pdf.github.io/fpdf2/).

## Features

- **Deduplication** — drops rows that are identical except for `id`
- **Date normalization** — parses 9 mixed formats (`DD/MM/YYYY`, `MM-DD-YYYY`, `DD Mon YYYY`, ISO, dotted, …) into `YYYY-MM-DD`; unparseable values become blank and are reported
- **Phone normalization** — strips non-digits and converts `+62…` / `8xx…` prefixes to Indonesian `08xx…` format
- **Text tidying** — names and cities to title case, emails to lowercase, all trimmed
- **Validation reporting** — invalid emails, failed dates, and per-column missing counts are printed, never silently dropped
- **Excel export** — bold header, freeze pane, autofilter, and proper Excel date formatting via openpyxl
- **PDF report** — summary metrics table plus a 10-row data sample via fpdf2

## Installation

Requires Python 3.10+.

```bash
pip install pandas openpyxl fpdf2
```

## Usage

```bash
# 1. Clean the raw CSV -> styled Excel workbook
python src/clean.py

# 2. Generate the PDF summary report
python src/report.py
```

Expected console output from `clean.py`:

```
Baris awal      : 25
Duplikat dibuang: 3
Baris akhir     : 22
Tanggal gagal parse (dikosongkan): 2
Email invalid   : ['dedi.kurniawan@email']
```

Artifacts produced:

| Script | Output |
|---|---|
| `src/clean.py` | `data/clean/customers_bersih.xlsx` — 22 clean rows, sheet `Customers` |
| `src/report.py` | `output/laporan_customers.pdf` — 1-page summary + 10-row sample |

## Project Structure

```
data-cleaner/
├── data/
│   ├── raw/
│   │   └── customers_kotor.csv   # raw input (do not edit)
│   └── clean/
│       └── customers_bersih.xlsx # cleaned output
├── output/
│   └── laporan_customers.pdf     # summary report
├── src/
│   ├── clean.py                  # cleaning pipeline
│   └── report.py                 # PDF report generator
└── README.md
```

## Input Data (`customers_kotor.csv`)

25 rows × 7 columns: `id, nama, email, no_hp, tanggal_daftar, kota, umur`.

Intentionally messy, mirroring real-world Indonesian customer data:

| Mess type | Examples |
|---|---|
| Full duplicates (different `id`) | id 1 = id 13, id 3 = id 4, id 21 = id 22 |
| Missing values | blank email (ids 5, 15, 24), name (6, 24), city (7, 15), age (8, 16, 21, 22) |
| Mixed date formats | `15/02/2024`, `03-15-2024`, `15 Jan 2024`, `2024/03/01`, `01.04.2024`, `2024.05.20`, `May 25 2024`, plus invalid `xxx` and `2024-13-45` |
| Inconsistent phone numbers | `+628…`, `62812-…`, spaces, parentheses, `telp:` prefix, letters (`0812-abc-…`), short numbers, missing leading `0`, blank |
| Misc | uppercase/spaced email (id 18), wrong-case/spaced cities (18, 19), invalid email (id 23) |

## Sample Output

Cleaned data (`customers_bersih.xlsx`, first rows):

| id | nama | email | no_hp | tanggal_daftar | kota | umur |
|---|---|---|---|---|---|---|
| 1 | Andi Pratama | andi.pratama@email.com | 081234567890 | 2024-01-15 | Jakarta | 25 |
| 2 | Siti Rahayu | siti.rahayu@email.com | 081234567891 | 2024-02-15 | Bandung | 30 |
| 3 | Budi Santoso | budi.santoso@email.com | 081234567892 | 2024-03-15 | Surabaya | 28 |
| 5 | Dewi Lestari | — | 081234567893 | 2024-01-15 | Medan | 22 |

PDF report (`laporan_customers.pdf`):

> **Laporan Data Cleaning - Customers**
>
> | Metric | Value |
> |---|---|
> | Total rows (raw) | 25 |
> | Duplicates removed | 3 |
> | Total rows (clean) | 22 |
> | Dates failed to parse | 2 |
> | Invalid emails | dedi.kurniawan@email |
>
> …followed by a 10-row sample of the cleaned data.

## Notes & Limitations

- Dates are interpreted **day-first (Indonesian convention)**; the US `MM-DD-YYYY` order is only a fallback, so `03-15-2024` still parses as March 15.
- Phone numbers that are too short to repair (e.g. `08127896`) are kept as-is rather than guessed.
- Flagged values (invalid emails, failed dates) are **reported, not deleted** — the reviewer decides what to do with them.
