"""
Update the whole PBS pricing database, end to end
====================================================
Runs the full pipeline in order, stopping if any step fails so later steps
never run on stale/partial data:

  1. update_pbs_data.py       - scrape pbs.gov.au for any new monthly XLSX
                                 files and append them to the master file
                                 (combined_df.csv, on G:\\My Drive\\PBS Pricing\\...)
  2. export_slim_parquet.py   - rebuild the slim file the Streamlit dashboard
                                 reads (combined_df_slim.csv.gz, in this folder).
                                 This also automatically back-fills AEMP for the
                                 Apr 2007-Aug 2013 gap every run (folded directly
                                 into that script - no separate R step needed).

Usage
-----
    python update_all_data.py                  # run everything, normal monthly update
    python update_all_data.py --dry-run         # step 1 only checks for new files, changes nothing
    python update_all_data.py --skip-scrape     # skip step 1, just rebuild step 2 from the
                                                 # existing master CSV (e.g. after fixing something
                                                 # by hand, or if you already ran step 1 separately)
    python update_all_data.py --from 2026-05    # force step 1 to re-check/re-download from May 2026
    python update_all_data.py --master-csv "D:\\path\\combined_df.csv"
                                                 # override the master CSV location for step 1
                                                 # (note: export_slim_parquet.py has its OWN hardcoded
                                                 # SOURCE_CSV constant - edit that too if you move the
                                                 # master file permanently)

Requirements: pandas, requests, beautifulsoup4, openpyxl (no R needed)
"""

import argparse
import subprocess
import sys
import time
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent

UPDATE_SCRIPT      = REPO_DIR / "update_pbs_data.py"
SLIM_SCRIPT        = REPO_DIR / "export_slim_parquet.py"
DEFAULT_MASTER_CSV = r"G:\My Drive\PBS Pricing\full data\combined_df.csv"


def banner(title: str):
    print(f"\n{'=' * 70}\n  {title}\n{'=' * 70}")


def run_step(label: str, cmd: list) -> bool:
    """Run a subprocess step, streaming its output live. Returns True on success."""
    banner(label)
    print(f"  $ {' '.join(str(c) for c in cmd)}\n")
    t0 = time.time()
    result = subprocess.run(cmd, cwd=REPO_DIR)
    elapsed = time.time() - t0
    ok = result.returncode == 0
    status = "OK" if ok else f"FAILED (exit code {result.returncode})"
    print(f"\n  --> {label}: {status}  ({elapsed:.0f}s)")
    return ok


def main():
    parser = argparse.ArgumentParser(description="Run the full PBS pricing database update pipeline.")
    parser.add_argument("--dry-run", action="store_true",
                         help="Step 1 only checks for new PBS files, changes nothing. Step 2 is skipped "
                              "too in dry-run mode (there's nothing new for it to pick up).")
    parser.add_argument("--skip-scrape", action="store_true",
                         help="Skip step 1 (scrape) entirely - just rebuild the slim file from the "
                              "master CSV as it currently stands.")
    parser.add_argument("--from", dest="date_from", default=None, metavar="YYYY-MM",
                         help="Forwarded to update_pbs_data.py --from: force re-check/re-download from "
                              "this month onward.")
    parser.add_argument("--master-csv", default=DEFAULT_MASTER_CSV, metavar="PATH",
                         help=f"Master combined_df.csv location for step 1 (default: {DEFAULT_MASTER_CSV}). "
                              "NOTE: export_slim_parquet.py has its own hardcoded SOURCE_CSV constant - if "
                              "you override this, make sure that constant points to the same file.")
    args = parser.parse_args()

    banner("PBS PRICING DATABASE - FULL UPDATE")
    print(f"  Repo folder : {REPO_DIR}")
    print(f"  Master CSV  : {args.master_csv}")
    print(f"  Dry run     : {args.dry_run}")
    print(f"  Skip scrape : {args.skip_scrape}")

    # ---- Step 1: scrape + append new months to the master CSV -----------------
    if args.skip_scrape:
        print("\n[Step 1/2] Skipped (--skip-scrape).")
    else:
        if not UPDATE_SCRIPT.exists():
            print(f"\n[Step 1/2] ERROR: {UPDATE_SCRIPT.name} not found next to this script. Aborting.")
            sys.exit(1)
        cmd = [sys.executable, str(UPDATE_SCRIPT), "--csv", args.master_csv]
        if args.date_from:
            cmd += ["--from", args.date_from]
        if args.dry_run:
            cmd += ["--dry-run"]
        if not run_step("Step 1/2: Scrape pbs.gov.au for new monthly files", cmd):
            print("\nAborting pipeline - fix the error above before rebuilding the slim file "
                  "(step 2 would otherwise run on stale data without you noticing).")
            sys.exit(1)

    if args.dry_run:
        banner("DRY RUN COMPLETE")
        print("  Nothing was changed. Step 2 is skipped in dry-run mode.")
        return

    # ---- Step 2: rebuild the slim file the dashboard reads (incl. AEMP fill) --
    if not SLIM_SCRIPT.exists():
        print(f"\n[Step 2/2] ERROR: {SLIM_SCRIPT.name} not found next to this script. Aborting.")
        sys.exit(1)
    if not run_step("Step 2/2: Rebuild combined_df_slim.csv.gz (incl. AEMP back-fill)",
                     [sys.executable, str(SLIM_SCRIPT)]):
        sys.exit(1)

    banner("PIPELINE COMPLETE")
    print(f"  combined_df.csv         : {args.master_csv}")
    print(f"  combined_df_slim.csv.gz : {REPO_DIR / 'combined_df_slim.csv.gz'}")


if __name__ == "__main__":
    main()
