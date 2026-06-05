"""
PBS Ex-Manufacturer Price Updater
===================================
1. Scrapes the PBS ex-manufacturer price page to find all available XLSX files.
2. Reads the latest date already in combined_df.csv.
3. Downloads any EFC / non-EFC files newer than that date.
4. Normalises and appends them to the CSV.

Usage
-----
    python update_pbs_data.py                        # auto-detect latest date
    python update_pbs_data.py --from 2026-02         # force re-download from Feb 2026
    python update_pbs_data.py --dry-run              # check without saving
    python update_pbs_data.py --csv "C:/path/to/combined_df.csv"

Requirements
------------
    pip install pandas requests openpyxl beautifulsoup4
"""

import argparse
import io
import re
import sys
import time
from datetime import date
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ─── Configuration ─────────────────────────────────────────────────────────────
DEFAULT_CSV = Path(r"G:\My Drive\PBS Pricing\full data\combined_df.csv")

PBS_PAGE    = "https://www.pbs.gov.au/info/industry/pricing/ex-manufacturer-price"
BASE_DOMAIN = "https://www.pbs.gov.au"
HEADERS     = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}
REQUEST_DELAY = 0.3   # seconds between requests

# ─── Column mappings ──────────────────────────────────────────────────────────

EFC_RENAME = {
    "Item Code":                  "item_code",
    "Legal Instrument Drug":      "drug_name",
    "Legal Instrument Form":      "form",
    "Legal Instrument MoA":       "route",
    "Brand Name":                 "brand_name",
    "Formulary":                  "formulary",
    "Program":                    "program",
    "Manufacturer Code":          "manufacturer_code",
    "Responsible Person":         "responsible_person",
    "Pack Quantity":               "pack_quantity",
    "Pricing Quantity":            "pricing_quantity",
    "Vial Content":                "vial_content",
    "Maximum Amount":              "maximum_amount",
    "Number Repeats":              "maximum_repeats",
    "AEMP":                       "aemp",
    "PEMP":                       "pemp",
    "Ex-man Price per Vial":      "ex_man_price_per_vial",
    "DPMA":                       "dpmq_dpma",
    "Claimed Price for Pack":     "claimed_price_for_pack",
    "Claimed Price for vial":     "claimed_price_for_vial",
    "Claimed DPMA":               "claimed_dpmq_dpma",
    "Premium":                    "premium",
    "Commonwealth Pays Premium":  "cwlth_pays_premium",
    "Maximum Patient Charge":     "maximum_patient_charge",
    "ATC":                        "atc",
}

NON_EFC_RENAME = {
    "Item Code":                      "item_code",
    "Legal Instrument Drug":          "drug_name",
    "Legal Instrument Form":          "form",
    "Legal Instrument MoA":           "route",
    "Brand Name":                     "brand_name",
    "Formulary":                      "formulary",
    "Program":                        "program",
    "Manufacturer Code":              "manufacturer_code",
    "Responsible Person":             "responsible_person",
    "Pack Quantity":                   "pack_quantity",
    "Pricing Quantity":                "pricing_quantity",
    "Maximum Quantity":                "maximum_quantity",
    "Maximum Repeats":                 "maximum_repeats",
    "AEMP":                           "aemp",
    "PEMP":                           "pemp",
    "Price to Pharmacy":              "price_to_pharmacy",
    "DPMQ":                           "dpmq_dpma",
    "Claimed Price for Pack":         "claimed_price_for_pack",
    "Claimed Price to Pharmacist":    "claimed_price_to_pharmacist",
    "Claimed DPMQ":                   "claimed_dpmq_dpma",
    "Premium":                        "premium",
    "C'wlth Pays Premium":           "cwlth_pays_premium",
    "Maximum Patient Charge":         "maximum_patient_charge",
    "ATC":                            "atc",
}

AMT_VARIANTS = [
    "AMT Trade Product Pack",
    "AMT Trade Product pack",
    "AMT Trade product Pack",
    "ANT Trade Product Pack",
    "AMT Trade Product Pack Pack",
    "Amt Trade Product Pack",
    "TPP",
]

CANONICAL_COLS = [
    "price_date", "item_code", "drug_name", "brand_name", "form", "route",
    "formulary", "responsible_person", "pack_quantity", "pricing_quantity",
    "maximum_repeats", "maximum_amount", "aemp", "dpmq_dpma", "program",
    "manufacturer_code", "pemp", "claimed_price_for_pack", "premium",
    "maximum_patient_charge", "amt_trade_product_pack", "vial_content",
    "claimed_price_for_vial", "ex_man_price_per_vial", "maximum_quantity",
    "price_to_pharmacy", "claimed_price_to_pharmacist", "claimed_dpmq_dpma",
    "cwlth_pays_premium", "drug_type_code", "atc", "atc_type", "atc_print_option",
    "caution_flag", "note_flag", "markup_code", "df_type", "dd_fee_code",
    "bp", "tgp", "tg_ptp", "tg_dpmq", "m_ptp", "m_dpmq", "maximum_safety_net",
    "bioequivalence_indicator", "source", "price_date_raw",
]

# ─── PBS page scraper ─────────────────────────────────────────────────────────

def scrape_available_files() -> list[dict]:
    """
    Fetch the PBS ex-manufacturer price page and return every XLSX link found,
    each with its date and source type (efc / non_efc / pre_split).
    """
    print(f"  Scraping PBS page for available files ...")
    resp = requests.get(PBS_PAGE, headers=HEADERS, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    files = []

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()

        # Only XLSX links that match the ex-manufacturer price pattern
        if "ex-manufacturer-price" not in href.lower():
            continue
        if not href.lower().endswith(".xlsx"):
            continue

        # Make absolute URL
        if href.startswith("/"):
            href = BASE_DOMAIN + href
        elif not href.startswith("http"):
            continue

        # Classify: non-EFC must come before EFC (it contains "efc" too)
        name = href.lower()
        if "non-efc" in name:
            source = "non_efc"
        elif "efc" in name:
            source = "efc"
        else:
            source = "pre_split"

        # Extract the date from the filename (e.g. 2026-05-01)
        m = re.search(r"(\d{4})-(\d{2})-\d{2}", href)
        if not m:
            continue
        file_date = date(int(m.group(1)), int(m.group(2)), 1)

        files.append({"url": href, "source": source, "date": file_date})

    # Deduplicate (same URL may appear more than once on the page)
    seen = set()
    unique = []
    for f in files:
        key = f["url"]
        if key not in seen:
            seen.add(key)
            unique.append(f)

    unique.sort(key=lambda x: (x["date"], x["source"]))
    print(f"  Found {len(unique)} XLSX file(s) on PBS website.")
    return unique


# ─── Download helpers ─────────────────────────────────────────────────────────

def try_download(url: str) -> bytes | None:
    """Download a single URL, trying both .XLSX and .xlsx variants."""
    for candidate in [url, url.replace(".XLSX", ".xlsx"), url.replace(".xlsx", ".XLSX")]:
        try:
            resp = requests.get(candidate, headers=HEADERS, timeout=60)
            if resp.status_code == 200 and len(resp.content) > 1_000:
                return resp.content
        except requests.RequestException:
            pass
    return None


# ─── Parsing & normalisation ──────────────────────────────────────────────────

def read_xlsx(raw: bytes) -> pd.DataFrame | None:
    try:
        xl = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
        for sheet in xl.sheet_names:
            df = xl.parse(sheet, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            df.dropna(how="all", inplace=True)
            df = df.loc[:, ~df.columns.str.fullmatch(r"Unnamed.*")]
            if not df.empty:
                return df
    except Exception as e:
        print(f"    [WARN] Could not parse XLSX: {e}")
    return None


def coalesce_amt_column(df: pd.DataFrame) -> pd.DataFrame:
    primary = "AMT Trade Product Pack"
    if primary not in df.columns:
        df[primary] = pd.NA
    for variant in AMT_VARIANTS[1:]:
        if variant in df.columns:
            df[primary] = df[primary].fillna(df[variant])
            df.drop(columns=[variant], inplace=True, errors="ignore")
    return df


def normalise(df: pd.DataFrame, rename_map: dict, price_date: date, source: str) -> pd.DataFrame:
    df = coalesce_amt_column(df)
    df = df.rename(columns={**rename_map, "AMT Trade Product Pack": "amt_trade_product_pack"})
    date_str = price_date.isoformat()
    df.insert(0, "price_date", date_str)
    df["source"] = source
    df["price_date_raw"] = date_str
    return df


def align_to_schema(new_df: pd.DataFrame, existing_cols: list) -> pd.DataFrame:
    for col in existing_cols:
        if col not in new_df.columns:
            new_df[col] = pd.NA
    extras = [c for c in new_df.columns if c not in existing_cols]
    return new_df[existing_cols + extras]


# ─── Encoding detection ───────────────────────────────────────────────────────

def detect_encoding(csv_path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(csv_path, encoding=enc) as f:
                f.read(65536)
            return enc
        except (UnicodeDecodeError, LookupError):
            pass
    return "latin-1"


# ─── Main ─────────────────────────────────────────────────────────────────────

def run(csv_path: Path, from_date: date | None = None, dry_run: bool = False):
    if not csv_path.exists():
        print(f"ERROR: CSV not found at:\n  {csv_path}")
        print("Edit DEFAULT_CSV at the top of this script or use --csv")
        sys.exit(1)

    size_mb = csv_path.stat().st_size / 1_048_576
    enc = detect_encoding(csv_path)

    print(f"\n{'='*60}")
    print(f"  PBS Ex-Manufacturer Price Updater")
    print(f"{'='*60}")
    print(f"  File     : {csv_path.name}  ({size_mb:.0f} MB)")
    print(f"  Encoding : {enc}")

    # ── Find latest date already in the CSV ───────────────────────────────────
    print(f"  Reading latest date from dataset ...", end=" ", flush=True)
    dates_only = pd.read_csv(csv_path, usecols=["price_date"], low_memory=False,
                             encoding=enc, encoding_errors="replace")
    dates_only["price_date"] = pd.to_datetime(dates_only["price_date"], errors="coerce")
    latest_date = dates_only["price_date"].max().date()
    cutoff = from_date if from_date else latest_date
    print(f"{latest_date.strftime('%B %Y')}")
    if from_date:
        print(f"  Forcing re-download from : {from_date.strftime('%B %Y')}")

    # ── Scrape PBS page for available files ───────────────────────────────────
    print()
    available = scrape_available_files()

    # Filter to files newer than our cutoff date
    to_download = [f for f in available if f["date"] > cutoff]

    if not to_download:
        print(f"\n  ✅ Already up to date — latest data is {latest_date.strftime('%B %Y')}.")
        print(f"     (PBS page has {len(available)} files; none newer than {cutoff.strftime('%B %Y')})")
        return

    # Show what we're about to download
    months_found = sorted(set(f["date"].strftime("%B %Y") for f in to_download))
    print(f"\n  New data found: {', '.join(months_found)}")
    print(f"  Files to download: {len(to_download)}")
    print()

    if dry_run:
        print("  [DRY RUN] — files that would be downloaded:")
        for f in to_download:
            print(f"    {f['date'].strftime('%B %Y')}  {f['source']:8s}  {f['url'].split('/')[-1]}")
        print()
        return

    # ── Download and normalise ────────────────────────────────────────────────
    new_frames = []
    rename_map = {"efc": EFC_RENAME, "non_efc": NON_EFC_RENAME, "pre_split": {}}

    for f in to_download:
        label   = f["date"].strftime("%B %Y")
        source  = f["source"]
        url     = f["url"]
        fname   = url.split("/")[-1]

        time.sleep(REQUEST_DELAY)
        raw = try_download(url)
        if raw is None:
            print(f"  [{label}]  {source:8s}  ✗  (download failed)")
            continue

        df_raw = read_xlsx(raw)
        if df_raw is None:
            print(f"  [{label}]  {source:8s}  ✗  (parse error)")
            continue

        rmap = rename_map.get(source, {})
        df_norm = normalise(df_raw, rmap.copy(), f["date"], source)
        new_frames.append(df_norm)
        print(f"  [{label}]  {source:8s}  ✓  {len(df_norm):>6,} rows  ← {fname}")

    if not new_frames:
        print(f"\n  ✅ No new rows were successfully downloaded.")
        return

    total_new = sum(len(f) for f in new_frames)
    print(f"\n  {total_new:,} new rows ready to append.")

    # ── Load full CSV and append ──────────────────────────────────────────────
    print(f"  Loading full dataset ({size_mb:.0f} MB) ...")
    existing = pd.read_csv(csv_path, low_memory=False, dtype=str,
                           encoding=enc, encoding_errors="replace")
    existing_cols = list(existing.columns)
    print(f"  Existing rows : {len(existing):,}")

    new_df   = pd.concat(new_frames, ignore_index=True, sort=False)
    new_df   = align_to_schema(new_df, existing_cols)
    combined = pd.concat([existing, new_df], ignore_index=True, sort=False)

    combined["_sort_date"] = pd.to_datetime(combined["price_date"], errors="coerce")
    combined.sort_values(["drug_name", "item_code", "_sort_date"], inplace=True)
    combined.drop(columns=["_sort_date"], inplace=True)
    combined.reset_index(drop=True, inplace=True)

    print(f"  Saving {len(combined):,} rows → {csv_path.name} ...")
    combined.to_csv(csv_path, index=False, encoding="utf-8")

    print(f"\n{'='*60}")
    print(f"  ✅ Update complete!")
    print(f"     Rows added : {total_new:,}")
    print(f"     Total rows : {len(combined):,}")
    print(f"     Saved to   : {csv_path}")
    print(f"{'='*60}\n")


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Update PBS ex-manufacturer price CSV with new monthly data."
    )
    parser.add_argument(
        "--csv", default=str(DEFAULT_CSV), metavar="PATH",
        help=f"Path to combined_df.csv (default: {DEFAULT_CSV})"
    )
    parser.add_argument(
        "--from", dest="date_from", default=None, metavar="YYYY-MM",
        help="Force re-download from this month, e.g. --from 2026-05"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Show what would be downloaded without saving anything"
    )
    args = parser.parse_args()

    from_date = None
    if args.date_from:
        try:
            yr, mo = args.date_from.split("-")
            from_date = date(int(yr), int(mo), 1)
        except ValueError:
            print("ERROR: --from must be in YYYY-MM format, e.g. --from 2026-05")
            sys.exit(1)

    run(Path(args.csv), from_date, args.dry_run)
