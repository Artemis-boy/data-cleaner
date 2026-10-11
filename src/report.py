"""Generate laporan PDF dari hasil cleaning.

Membaca data/clean/customers_bersih.xlsx, menghitung ringkasan
(dengan menghitung ulang dari data raw agar angkanya konsisten
dengan clean.py), lalu menyimpan output/laporan_customers.pdf
berisi judul, tabel ringkasan, dan 10 baris contoh data.
"""
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from fpdf import FPDF

ROOT = Path(__file__).resolve().parent.parent
CLEAN_XLSX = ROOT / "data" / "clean" / "customers_bersih.xlsx"
RAW_CSV = ROOT / "data" / "raw" / "customers_kotor.csv"
OUT_PDF = ROOT / "output" / "laporan_customers.pdf"

sys.path.insert(0, str(Path(__file__).resolve().parent))
from clean import EMAIL_RE, clean_date  # noqa: E402


def compute_summary() -> dict:
    raw = pd.read_csv(RAW_CSV, dtype=str, keep_default_na=True)
    for col in raw.columns:
        if pd.api.types.is_string_dtype(raw[col]):
            raw[col] = raw[col].str.strip().replace(r"^\s*$", pd.NA, regex=True)

    total_raw = len(raw)
    subset = [c for c in raw.columns if c != "id"]
    duplikat = int(raw.duplicated(subset=subset, keep="first").sum())

    tanggal_gagal = 0
    if "tanggal_daftar" in raw.columns:
        parsed = raw["tanggal_daftar"].apply(clean_date)
        tanggal_gagal = int(raw["tanggal_daftar"].notna().sum() - parsed.notna().sum())

    email_invalid = []
    if "email" in raw.columns:
        emails = raw["email"].str.lower().dropna()
        email_invalid = emails[~emails.str.match(EMAIL_RE)].tolist()

    bersih = pd.read_excel(CLEAN_XLSX, sheet_name="Customers", dtype=str)
    return {
        "total_raw": total_raw,
        "duplikat": duplikat,
        "total_bersih": len(bersih),
        "tanggal_gagal": tanggal_gagal,
        "email_invalid": email_invalid,
        "df": bersih,
    }


def build_pdf(s: dict) -> None:
    pdf = FPDF(orientation="L", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 12, "Laporan Data Cleaning - Customers", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(
        0, 7,
        f"Tanggal laporan: {date.today().isoformat()}  |  Sumber: data/raw/customers_kotor.csv",
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.ln(4)

    # --- Tabel ringkasan ---
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 9, "Ringkasan", new_x="LMARGIN", new_y="NEXT")
    rows = [
        ("Total baris (raw)", str(s["total_raw"])),
        ("Duplikat dibuang", str(s["duplikat"])),
        ("Total baris (bersih)", str(s["total_bersih"])),
        ("Tanggal gagal parse (dikosongkan)", str(s["tanggal_gagal"])),
        ("Email invalid", ", ".join(s["email_invalid"]) if s["email_invalid"] else "-"),
    ]
    col_w = [90, 187]
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(col_w[0], 8, "Metrik", border=1, fill=True)
    pdf.cell(col_w[1], 8, "Nilai", border=1, fill=True, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    for metric, value in rows:
        h = 8
        # email invalid bisa panjang -> pakai multi_cell untuk kolom nilai
        if metric == "Email invalid" and len(value) > 100:
            y0 = pdf.get_y()
            pdf.multi_cell(col_w[0], h, metric, border=1)
            y1 = pdf.get_y()
            pdf.set_xy(pdf.l_margin + col_w[0], y0)
            pdf.multi_cell(col_w[1], h, value, border=1)
            pdf.set_y(max(y1, pdf.get_y()))
        else:
            pdf.cell(col_w[0], h, metric, border=1)
            pdf.cell(col_w[1], h, value, border=1, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)

    # --- 10 baris contoh ---
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 9, "Contoh Data Bersih (10 baris pertama)", new_x="LMARGIN", new_y="NEXT")
    df = s["df"].head(10)
    cols = list(df.columns)
    page_w = 297 - 20  # A4 landscape lebar usable (margin 10+10)
    widths = {"id": 12, "nama": 45, "email": 65, "no_hp": 32,
              "tanggal_daftar": 30, "kota": 35, "umur": 14}
    col_ws = [widths.get(c, 30) for c in cols]
    scale = page_w / sum(col_ws)
    col_ws = [w * scale for w in col_ws]

    def fmt(v):
        if pd.isna(v):
            return "-"
        if isinstance(v, (datetime, pd.Timestamp)):
            return v.strftime("%Y-%m-%d")
        s = str(v)
        m = re.match(r"^(\d{4}-\d{2}-\d{2})[ T]00:00:00$", s)
        return m.group(1) if m else s

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(230, 230, 230)
    for c, w in zip(cols, col_ws):
        pdf.cell(w, 7, c, border=1, fill=True)
    pdf.ln()
    pdf.set_font("Helvetica", "", 8)
    for _, row in df.iterrows():
        for c, w in zip(cols, col_ws):
            pdf.cell(w, 6, fmt(row[c]), border=1)
        pdf.ln()

    OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(OUT_PDF))
    print(f"Tersimpan: {OUT_PDF}")


def main() -> None:
    s = compute_summary()
    print(f"Total raw: {s['total_raw']}, duplikat: {s['duplikat']}, "
          f"bersih: {s['total_bersih']}, tgl gagal: {s['tanggal_gagal']}, "
          f"email invalid: {s['email_invalid'] or '-'}")
    build_pdf(s)


if __name__ == "__main__":
    main()
