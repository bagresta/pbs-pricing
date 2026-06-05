"""
Export a slim Parquet file for Streamlit Cloud
===============================================
Reads the full combined_df.csv, keeps only the columns the dashboard needs,
applies categorical dtypes, and saves as a compressed Parquet file.

The Parquet file is ~30-50 MB (vs 624 MB CSV) and loads much faster
with far less memory — solving Streamlit Cloud's 1 GB RAM limit.

Usage
-----
    python export_slim_parquet.py

Upload the output file to Google Drive and update CLOUD_URL in pbs_dashboard.py.
"""

from pathlib import Path
import pandas as pd

SOURCE_CSV  = Path(r"G:\My Drive\PBS Pricing\full data\combined_df.csv")
OUTPUT_FILE = Path(r"C:\Users\bagre\OneDrive\Documentos\psd data\Pricing database\combined_df_slim.csv.gz")

WANTED = ["price_date", "source", "item_code", "drug_name", "brand_name",
          "form", "aemp", "dpmq_dpma", "formulary", "atc"]

SOURCE_NORMALISE = {
    "non_efc1": "non_efc", "non_efc2": "non_efc",
    "nonefc1":  "non_efc", "nonefc2":  "non_efc",
    "nonefc":   "non_efc",
}

def detect_encoding(p: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(p, encoding=enc) as f:
                f.read(65536)
            return enc
        except (UnicodeDecodeError, LookupError):
            pass
    return "latin-1"

def main():
    print(f"Reading {SOURCE_CSV.name} ({SOURCE_CSV.stat().st_size / 1e6:.0f} MB) ...")
    enc = detect_encoding(SOURCE_CSV)

    header = pd.read_csv(SOURCE_CSV, nrows=0, encoding=enc, encoding_errors="replace")
    available = [c for c in WANTED if c in header.columns]

    df = pd.read_csv(SOURCE_CSV, usecols=available, low_memory=False,
                     encoding=enc, encoding_errors="replace")
    print(f"  Loaded {len(df):,} rows, {len(df.columns)} columns")

    # Normalise
    df["price_date"] = pd.to_datetime(df["price_date"], errors="coerce")
    df = df.dropna(subset=["price_date"])

    for col in ["aemp", "dpmq_dpma"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "source" in df.columns:
        df["source"] = df["source"].str.strip().str.lower().replace(SOURCE_NORMALISE)

    for col in ["drug_name", "brand_name"]:
        if col in df.columns:
            df[col] = df[col].str.strip().str.title()

    df = df.sort_values("price_date").reset_index(drop=True)

    # Save as gzip-compressed CSV — universally compatible, no version issues.
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False, compression="gzip")

    size_mb = OUTPUT_FILE.stat().st_size / 1e6
    print(f"\n✅ Saved: {OUTPUT_FILE}")
    print(f"   Size  : {size_mb:.1f} MB  (was {SOURCE_CSV.stat().st_size / 1e6:.0f} MB)")
    print(f"   Rows  : {len(df):,}")
    print(f"\nNext steps:")
    print(f"  1. Share this file on Google Drive and copy the link")
    print(f"  2. Update CLOUD_URL in pbs_dashboard.py with the new link")

if __name__ == "__main__":
    main()
