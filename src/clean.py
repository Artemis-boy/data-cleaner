"""Data cleaning: customers_kotor.csv -> customers_bersih.xlsx.

Aturan:
- Duplikat: baris dengan semua field sama kecuali `id` dianggap duplikat.
- Tanggal -> YYYY-MM-DD (yang tak bisa diparse jadi kosong).
- No HP -> digit saja, prefix 62/8xx dinormalkan ke 08xx.
- Nama & kota -> trim + title case; email -> trim + lowercase.
- Field kosong (string berisi spasi saja) -> kosong (NaN).
"""
import re
from copy import copy
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw" / "customers_kotor.csv"
OUT = ROOT / "data" / "clean" / "customers_bersih.xlsx"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def clean_phone(value) -> str | None:
    """Normalisasi nomor HP Indonesia ke format 08xx (digit saja)."""
    if pd.isna(value):
        return None
    digits = re.sub(r"\D", "", str(value))
    if not digits:
        return None
    if digits.startswith("62"):
        digits = "0" + digits[2:]
    elif digits.startswith("8"):
        digits = "0" + digits
    # '0...' dibiarkan; sisanya (terlalu pendek dsb.) dibiarkan apa adanya
    return digits or None


# Urutan penting: DD-MM-YYYY (Indonesia) SEBELUM MM-DD-YYYY (Amerika),
# agar tanggal ambigu (01-04-2024) ditafsir 1 April.
# Format Amerika tetap didukung: 03-15-2024 gagal di %d-%m-%Y lalu cocok di %m-%d-%Y.
_DATE_FORMATS = (
    "%Y-%m-%d",   # 2024-01-15
    "%d/%m/%Y",   # 15/02/2024
    "%d-%m-%Y",   # 12-05-2024 -> 12 Mei
    "%d.%m.%Y",   # 01.04.2024 -> 1 April
    "%Y/%m/%d",   # 2024/03/01
    "%Y.%m.%d",   # 2024.05.20
    "%m-%d-%Y",   # 03-15-2024 -> 15 Mar (fallback Amerika)
    "%d %b %Y",   # 15 Jan 2024
    "%b %d %Y",   # May 25 2024
)


def clean_date(value):
    """Parse bermacam format tanggal -> date; gagal -> NaT."""
    if pd.isna(value) or str(value).strip() in ("", "xxx"):
        return pd.NaT
    text = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return pd.to_datetime(text, format=fmt).date()
        except (ValueError, TypeError):
            continue
    return pd.NaT


def main() -> None:
    df = pd.read_csv(RAW, dtype=str, keep_default_na=True)
    # string hanya-spasi -> NA agar seragam (pandas 3: kolom string ber-dtype `str`)
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].str.strip().replace(r"^\s*$", pd.NA, regex=True)
    n_awal = len(df)

    # 1. Duplikat (abaikan kolom id karena id selalu unik)
    subset = [c for c in df.columns if c != "id"]
    duplikat = int(df.duplicated(subset=subset, keep="first").sum())
    df = df.drop_duplicates(subset=subset, keep="first").reset_index(drop=True)

    # 2. Nama & kota: title case
    for col in ("nama", "kota"):
        if col in df.columns:
            df[col] = df[col].str.title()

    # 3. Email: lowercase (validasi ringan, yang invalid dibiarkan + dilaporkan)
    if "email" in df.columns:
        df["email"] = df["email"].str.lower()
        invalid_email = df["email"].dropna()
        invalid_email = invalid_email[~invalid_email.str.match(EMAIL_RE)].tolist()
    else:
        invalid_email = []

    # 4. Tanggal -> YYYY-MM-DD
    tanggal_gagal = 0
    if "tanggal_daftar" in df.columns:
        sebelum = df["tanggal_daftar"].notna().sum()
        df["tanggal_daftar"] = df["tanggal_daftar"].apply(clean_date)
        tanggal_gagal = int(sebelum - df["tanggal_daftar"].notna().sum())
        df["tanggal_daftar"] = pd.to_datetime(df["tanggal_daftar"], errors="coerce")

    # 5. No HP -> 08xx
    if "no_hp" in df.columns:
        df["no_hp"] = df["no_hp"].apply(clean_phone)

    # 6. Umur -> numerik (yang bukan angka jadi kosong)
    if "umur" in df.columns:
        df["umur"] = pd.to_numeric(df["umur"], errors="coerce").astype("Int64")

    # 7. Simpan ke xlsx dengan header rapi
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Customers", index=False)
        ws = writer.sheets["Customers"]
        for cell in ws[1]:
            font = copy(cell.font)
            font.bold = True
            cell.font = font
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        widths = {"id": 6, "nama": 20, "email": 30, "no_hp": 16,
                  "tanggal_daftar": 15, "kota": 14, "umur": 8}
        for idx, col in enumerate(df.columns, start=1):
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = (
                widths.get(col, 15)
            )
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=list(df.columns).index("tanggal_daftar") + 1,
                                max_col=list(df.columns).index("tanggal_daftar") + 1):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = "YYYY-MM-DD"

    # Laporan ringkas
    print(f"Baris awal      : {n_awal}")
    print(f"Duplikat dibuang: {duplikat}")
    print(f"Baris akhir     : {len(df)}")
    print(f"Tanggal gagal parse (dikosongkan): {tanggal_gagal}")
    print(f"Email invalid   : {invalid_email if invalid_email else '-'}")
    print(f"Missing per kolom:\n{df.isna().sum()}")
    print(f"Tersimpan: {OUT}")


if __name__ == "__main__":
    main()
