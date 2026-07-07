"""
Export a slim Parquet file for Streamlit Cloud
===============================================
Reads the full combined_df.csv, keeps only the columns the dashboard needs,
applies categorical dtypes, and saves as a compressed Parquet file.

The Parquet file is ~30-50 MB (vs 624 MB CSV) and loads much faster
with far less memory — solving Streamlit Cloud's 1 GB RAM limit.

AEMP BACK-CALCULATION FOR THE APR 2007 - AUG 2013 GAP
-------------------------------------------------------
AEMP didn't exist as a legislated price point until 1 Oct 2012, so every row
dated before then has a blank "aemp" in the source data. This script now
automatically fills that gap on every run, using the item's Price to
Pharmacists (PTP) figure from the raw presplit-era file plus its PBS program
code, reversing whichever mark-up (if any) applied to that program - see
BACKCALC_NOTES below for the validated formula and its accuracy. This runs
every time (not a one-off), so it's always current with whatever's in
full data/presplit.csv - no separate R script or manual step needed.

TWO MONTHS PBS NEVER PUBLISHED AN UPDATE FOR
-----------------------------------------------
1 February 2012 and 1 November 2012 have no PTP at all in the raw source
(source tag "presplitmiss"), so the back-calculation above can't touch them.
This isn't a scrape failure - PBS's own ex-manufacturer-price page says, for
both months: "No updates this month - please refer to [prior month] prices
below." dpmq_dpma for these rows is a byte-for-byte carry-forward of the
prior month (verified), so the correct AEMP is likewise just the prior
month's AEMP, carried forward per item_code - see fill_no_update_months().

Usage
-----
    python export_slim_parquet.py

Upload the output file to Google Drive and update CLOUD_URL in pbs_dashboard.py.
"""

from pathlib import Path
import pandas as pd
import numpy as np

SOURCE_CSV   = Path(r"G:\My Drive\PBS Pricing\full data\combined_df.csv")
OUTPUT_FILE  = Path(r"C:\Users\bagre\OneDrive\Documentos\psd data\Pricing database\combined_df_slim.csv.gz")
PRESPLIT_CSV = OUTPUT_FILE.parent / "full data" / "presplit.csv"

WANTED = ["price_date", "source", "item_code", "drug_name", "brand_name",
          "form", "aemp", "dpmq_dpma", "formulary", "atc"]

SOURCE_NORMALISE = {
    "non_efc1": "non_efc", "non_efc2": "non_efc",
    "nonefc1":  "non_efc", "nonefc2":  "non_efc",
    "nonefc":   "non_efc",
}

# Source tag PBS itself used (via our scrape) for months where the ex-manufacturer
# price page said "No updates this month - please refer to [prior month] prices".
NO_UPDATE_SOURCE_TAG = "presplitmiss"

# ─── AEMP back-calculation parameters (validated against ~3,600 items whose
#     PTP->AEMP handover happened within 35 days of the Oct2012/Aug2013
#     changeover - see backcalculate_aemp_from_ptp.R for the full derivation
#     and accuracy breakdown by program code) ─────────────────────────────
BACKCALC_NOTES = """
  Wholesale mark-up formula (4CPA clause 14.2, confirmed effective 1 Jul 2006,
  unchanged through today): 7.52% of AEMP up to and including $930.06 AEMP
  (= $1000.00 PTP); flat $69.94 above that.
    PTP <= $1000.00  ->  AEMP = round(PTP / 1.0752, 2)
    PTP >  $1000.00  ->  AEMP = round(PTP - 69.94, 2)
  Applies to general-schedule/community program codes (GE, R1, PL, DB, PQ,
  DT, OT, SB). Validated exact-match rate: 83-90%.

  Section 100 program codes (HS, HB, GH, MD, CT, IF - Highly Specialised
  Drugs, Growth Hormone, IVF-related, most chemotherapy) carry NO wholesale
  mark-up at all: AEMP = PTP directly. Validated exact-match rate: 91-100%.

  Section 100 Efficient Funding of Chemotherapy codes (IP, IN) are NOT
  back-calculated - PTP:AEMP ratios for these vary per item in a way
  consistent with a pricing-quantity/vial-content unit mismatch, not a
  mark-up formula (0% match to either formula above).

  Feb 2012 and Nov 2012: PBS published no price update at all for these
  months (confirmed on pbs.gov.au's ex-manufacturer-price page). AEMP is
  carried forward from the prior month, per item_code, since dpmq_dpma for
  these rows is already an identical carry-forward of the prior month.
"""

WHOLESALE_RATE   = 0.0752
FLAT_MARKUP      = 69.94
PTP_THRESHOLD    = 930.06 * (1 + WHOLESALE_RATE)  # = $1000.00
MARKUP_CODES     = {"GE", "R1", "PL", "DB", "PQ", "DT", "OT", "SB"}
NO_MARKUP_CODES  = {"HS", "HB", "GH", "MD", "CT", "IF"}
UNRESOLVED_CODES = {"IP", "IN"}


def detect_encoding(p: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(p, encoding=enc) as f:
                f.read(65536)
            return enc
        except (UnicodeDecodeError, LookupError):
            pass
    return "latin-1"


def backcalculate_aemp(ptp: pd.Series, drug_type_code: pd.Series) -> pd.Series:
    """Program-aware AEMP back-calculation from PTP. Returns NaN for
    unresolved/unknown program codes."""
    is_no_markup = drug_type_code.isin(NO_MARKUP_CODES)
    is_markup    = drug_type_code.isin(MARKUP_CODES)

    above_threshold = ptp > PTP_THRESHOLD
    markup_result = np.where(
        above_threshold, (ptp - FLAT_MARKUP).round(2), (ptp / (1 + WHOLESALE_RATE)).round(2)
    )

    result = np.full(len(ptp), np.nan)
    result = np.where(is_no_markup, ptp.round(2), result)
    result = np.where(is_markup, markup_result, result)
    return pd.Series(result, index=ptp.index)


def fill_historical_aemp_gap(df: pd.DataFrame) -> pd.DataFrame:
    """Back-fill blank aemp for Apr 2007-Aug 2013 rows using PTP from the raw
    presplit file, program-code-aware. Adds an aemp_source column so
    back-calculated values stay distinguishable from genuinely published
    ones. No-op (with a warning) if presplit.csv isn't available locally."""
    df["aemp_source"] = np.where(df["aemp"].notna(), "published", pd.NA)

    if not PRESPLIT_CSV.exists():
        print(f"\n  [WARN] {PRESPLIT_CSV} not found - skipping AEMP back-calculation "
              f"for the Apr 2007-Aug 2013 gap. combined_df_slim.csv.gz will still have "
              f"blank aemp for that window this run.")
        df.loc[df["aemp_source"].isna(), "aemp_source"] = "still_missing"
        return df

    print(f"\n  Filling historical AEMP gap from {PRESPLIT_CSV.name} ...")
    presplit = pd.read_csv(
        PRESPLIT_CSV, dtype={"item_code": str}, low_memory=False,
        usecols=["price_date", "item_code", "brand_name", "price_to_pharmacy", "drug_type_code"],
    )
    presplit["price_date"] = pd.to_datetime(presplit["price_date"], errors="coerce")
    presplit["price_to_pharmacy"] = pd.to_numeric(presplit["price_to_pharmacy"], errors="coerce")
    presplit = presplit.dropna(subset=["price_date", "price_to_pharmacy"])
    presplit["brand_name"] = presplit["brand_name"].astype(str).str.strip().str.title()

    # IMPORTANT: item_code + price_date is NOT a unique key in this raw file - PBS
    # item codes are shared across multiple interchangeable brands (e.g. one item
    # code covers Zovirax/Acyclo-V/Acihexal/Lovir, each a different price). Merging
    # on item_code + price_date alone silently fans out ~47% of rows (one row per
    # brand all matching the same key), corrupting row counts. Joining on
    # item_code + price_date + brand_name instead brings duplicate keys down to
    # ~3.5% (residual multi-manufacturer/pack-size variants under the same brand);
    # those are deduplicated (keep first) so the join can never inflate row counts.
    presplit["aemp_backcalculated"] = backcalculate_aemp(
        presplit["price_to_pharmacy"], presplit["drug_type_code"]
    )
    presplit["backcalc_confidence"] = np.select(
        [presplit["drug_type_code"].isin(UNRESOLVED_CODES),
         presplit["drug_type_code"].isin(NO_MARKUP_CODES),
         presplit["drug_type_code"].isin(MARKUP_CODES)],
        ["unit_mismatch_unresolved", "high_no_markup", "high"],
        default="other_program_uncertain",
    )
    presplit = presplit.drop_duplicates(subset=["item_code", "price_date", "brand_name"], keep="first")

    df["item_code"] = df["item_code"].astype(str)
    df["brand_name"] = df["brand_name"].astype(str).str.strip().str.title()
    fill = presplit[["item_code", "price_date", "brand_name", "aemp_backcalculated", "backcalc_confidence"]]
    rows_before = len(df)
    df = df.merge(fill, on=["item_code", "price_date", "brand_name"], how="left")
    assert len(df) == rows_before, (
        f"Merge changed row count ({rows_before:,} -> {len(df):,}) - the join key is no "
        f"longer unique enough on the presplit side. Do not trust this output; investigate "
        f"before using it."
    )

    needs_fill = df["aemp"].isna() & df["aemp_backcalculated"].notna()
    n_filled = int(needs_fill.sum())
    df.loc[needs_fill, "aemp"] = df.loc[needs_fill, "aemp_backcalculated"]
    df.loc[needs_fill, "aemp_source"] = "backcalculated_" + df.loc[needs_fill, "backcalc_confidence"]
    df.loc[df["aemp_source"].isna(), "aemp_source"] = "still_missing"

    df = df.drop(columns=["aemp_backcalculated", "backcalc_confidence"])

    print(f"  Back-calculated AEMP for {n_filled:,} rows.")
    print("  aemp_source breakdown:")
    print(df["aemp_source"].value_counts().to_string())
    return df


def fill_no_update_months(df: pd.DataFrame) -> pd.DataFrame:
    """Fill AEMP for rows tagged as a PBS 'no updates this month' carry-forward
    (currently: Feb 2012, Nov 2012 - source == 'presplitmiss'). These months
    have no PTP at all, so fill_historical_aemp_gap() can't touch them - but
    PBS's own page says these months are identical to the prior month, and
    dpmq_dpma for these rows is already a verified byte-for-byte carry-forward
    of the prior month's dpmq_dpma. So the correct AEMP is likewise the prior
    month's AEMP, carried forward per item_code - not a fresh estimate, just
    reusing the number PBS itself says didn't change."""
    still_missing = df["aemp_source"] == "still_missing"
    no_update_rows = still_missing & (df["source"] == NO_UPDATE_SOURCE_TAG)
    if not no_update_rows.any():
        return df

    df = df.sort_values(["item_code", "price_date"]).reset_index(drop=True)
    carried = df.groupby("item_code")["aemp"].ffill()

    apply_mask = (df["aemp_source"] == "still_missing") & (df["source"] == NO_UPDATE_SOURCE_TAG)
    apply_mask = apply_mask & carried.notna()
    n_carried = int(apply_mask.sum())

    df.loc[apply_mask, "aemp"] = carried[apply_mask]
    df.loc[apply_mask, "aemp_source"] = "carried_forward_pbs_no_update"

    print(f"\n  Carried forward AEMP for {n_carried:,} rows PBS itself flagged as "
          f"'no updates this month' (source == '{NO_UPDATE_SOURCE_TAG}').")
    return df


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

    # Back-fill AEMP for the Apr 2007 - Aug 2013 gap, every run - permanent
    # and self-refreshing, no separate manual/R step needed.
    df = fill_historical_aemp_gap(df)

    # Then carry forward AEMP for the two "PBS published no update" months
    # that the step above can't reach (no PTP exists for them at all).
    df = fill_no_update_months(df)

    print("\n  Final aemp_source breakdown:")
    print(df["aemp_source"].value_counts().to_string())
    print(BACKCALC_NOTES)

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
