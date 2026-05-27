"""
PBS Ex-Manufacturer Price Dashboard
=====================================
- Checks current dataset for latest date
- Checks PBS website for newer monthly files
- Downloads and appends any new data
- Visualises AEMP and DPMQ/DPMA over time
- Annotates F1→F2 formulary transitions
- Annotates price decreases with % change
- Export charts as PNG

Run with:
    streamlit run pbs_dashboard.py
"""

import io
import time
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from plotly.subplots import make_subplots

# ─── Data source config ────────────────────────────────────────────────────────
#
# Option A — Local file (recommended).
#   Point this to your combined_df.csv.
#   If the file lives in a OneDrive / Google Drive synced folder, edits made by
#   update_pbs_data.py will automatically sync to the cloud.
#
DATA_FILE = Path(r"G:\My Drive\PBS Pricing\full data\combined_df.csv")
#
# Option B — Cloud download URL (OneDrive, Google Drive, Dropbox).
#   Leave as None to use DATA_FILE above.
#   Set to a direct-download URL and the app will download the file on first
#   load and cache it locally at CACHE_FILE.
#
#   OneDrive  : right-click file → Share → Copy link, then append ?download=1
#               e.g. "https://onedrive.live.com/download?cid=XXX&resid=YYY&authkey=ZZZ"
#   Google Drive: share file publicly, then use:
#               "https://drive.google.com/uc?export=download&id=FILE_ID"
#   Dropbox   : change ?dl=0 to ?dl=1 at the end of the share link
#
# When running LOCALLY: keep this as None — the dashboard reads directly from DATA_FILE above.
# When deployed to STREAMLIT CLOUD: paste your Google Drive share URL here.
CLOUD_URL: str | None = "https://drive.google.com/file/d/1UA7_JqQHi-f-lHF5I09Tr5gTo9-m8yH3/view?usp=sharing"
CACHE_FILE = Path(__file__).parent / "_combined_df_cache.parquet"

# ─── PBS config ────────────────────────────────────────────────────────────────
BASE_URL = "https://www.pbs.gov.au/industry/pricing/ex-manufacturer-price"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}
SPLIT_DATE = date(2013, 8, 1)

st.set_page_config(
    page_title="PBS Ex-Manufacturer Pricing",
    page_icon="💊",
    layout="wide",
)

# ─── PBS URL helpers ──────────────────────────────────────────────────────────

def urls_pre_split(yr: int, mo: int) -> list:
    fname = f"ex-manufacturer-prices-{yr}-{mo:02d}-01.XLSX"
    return [f"{BASE_URL}/{yr}/{fname}", f"{BASE_URL}/{yr}/{mo:02d}/{fname}"]

def urls_efc(yr: int, mo: int) -> list:
    fname = f"ex-manufacturer-prices-efc-{yr}-{mo:02d}-01.XLSX"
    return [f"{BASE_URL}/{yr}/{fname}", f"{BASE_URL}/{yr}/{mo:02d}/{fname}"]

def urls_non_efc(yr: int, mo: int) -> list:
    fname = f"ex-manufacturer-prices-non-efc-{yr}-{mo:02d}-01.XLSX"
    return [f"{BASE_URL}/{yr}/{fname}", f"{BASE_URL}/{yr}/{mo:02d}/{fname}"]

def try_download(urls: list):
    for url in urls:
        for candidate in [url, url.replace(".XLSX", ".xlsx"), url.replace(".xlsx", ".XLSX")]:
            try:
                resp = requests.get(candidate, headers=HEADERS, timeout=30)
                if resp.status_code == 200 and len(resp.content) > 1000:
                    return resp.content, candidate
            except Exception:
                pass
    return None, ""

def read_xlsx_bytes(data: bytes, price_date: date, source: str):
    try:
        xl = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
        for sheet in xl.sheet_names:
            df = xl.parse(sheet, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            df.dropna(how="all", inplace=True)
            df = df.loc[:, ~df.columns.str.fullmatch(r"Unnamed.*")]
            if df.empty:
                continue
            df.insert(0, "price_date", price_date.isoformat())
            df["source"] = source
            return df
    except Exception as e:
        st.warning(f"Could not parse XLSX: {e}")
    return None

def normalise_and_extract(df: pd.DataFrame, source: str) -> pd.DataFrame:
    if source == "efc":
        rename = {
            "Item Code": "item_code", "Legal Instrument Drug": "drug_name",
            "Legal Instrument Form": "form", "Legal Instrument MoA": "route",
            "Brand Name": "brand_name", "AEMP": "aemp", "DPMA": "dpmq_dpma",
            "Previous AEMP": "previous_aemp", "Price Change Event": "price_change_event",
            "Formulary": "formulary", "Program": "program", "ATC": "atc",
            "Pack Quantity": "pack_quantity", "Pricing Quantity": "pricing_quantity",
            "Premium": "premium",
        }
    elif source == "non_efc":
        rename = {
            "Item Code": "item_code", "Legal Instrument Drug": "drug_name",
            "Legal Instrument Form": "form", "Legal Instrument MoA": "route",
            "Brand Name": "brand_name", "AEMP": "aemp", "DPMQ": "dpmq_dpma",
            "Previous AEMP": "previous_aemp", "Price Change Event": "price_change_event",
            "Formulary": "formulary", "Program": "program", "ATC": "atc",
            "Pack Quantity": "pack_quantity", "Pricing Quantity": "pricing_quantity",
            "Premium": "premium",
        }
    else:
        rename = {
            "Item Code": "item_code", "Drug": "drug_name",
            "Form and Strength": "form", "Brand Name": "brand_name",
            "DPMQ": "dpmq_dpma", "Full ATC": "atc", "Pack Size": "pack_quantity",
        }
    return df.rename(columns=rename)

def month_range(start: date, end: date):
    yr, mo = start.year, start.month
    while date(yr, mo, 1) <= end:
        yield yr, mo
        mo += 1
        if mo > 12:
            mo, yr = 1, yr + 1

# ─── Data loading ─────────────────────────────────────────────────────────────

def _gdrive_file_id(url: str) -> str | None:
    """Extract a Google Drive file ID from a share URL, or return None."""
    import re
    m = re.search(r"/d/([A-Za-z0-9_-]+)", url)
    return m.group(1) if m else None


def _download_gdrive(file_id: str, dest: Path) -> None:
    """
    Download a (potentially large) Google Drive file, bypassing the virus-scan
    confirmation page that Google shows for files over ~100 MB.
    """
    session = requests.Session()

    # First request — may return a confirmation page for large files
    dl_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&authuser=0&confirm=t"
    resp = session.get(dl_url, stream=True, timeout=300, headers=HEADERS)
    resp.raise_for_status()

    # Stream to disk in chunks so we don't blow memory
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as fh:
        for chunk in resp.iter_content(chunk_size=1024 * 1024):   # 1 MB chunks
            if chunk:
                fh.write(chunk)


def resolve_data_file() -> Path:
    """
    Return the path to the CSV to use, downloading from CLOUD_URL if configured.
    Handles Google Drive large-file downloads automatically.
    Falls back to DATA_FILE if the cloud download fails.
    """
    if CLOUD_URL:
        if not CACHE_FILE.exists():
            with st.spinner("Downloading dataset from cloud — this may take a few minutes for a large file…"):
                try:
                    gd_id = _gdrive_file_id(CLOUD_URL)
                    if gd_id:
                        _download_gdrive(gd_id, CACHE_FILE)
                    else:
                        resp = requests.get(CLOUD_URL, stream=True, timeout=300, headers=HEADERS)
                        resp.raise_for_status()
                        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
                        with open(CACHE_FILE, "wb") as fh:
                            for chunk in resp.iter_content(chunk_size=1024 * 1024):
                                if chunk:
                                    fh.write(chunk)
                    st.success("Dataset downloaded and cached.")
                except Exception as e:
                    st.error(f"Cloud download failed: {e}")
                    # Clean up partial file so next load retries
                    if CACHE_FILE.exists():
                        CACHE_FILE.unlink()
                    if DATA_FILE.exists():
                        st.info("Falling back to local file.")
                        return DATA_FILE
                    return CACHE_FILE   # will fail gracefully in load_data
        return CACHE_FILE

    if not DATA_FILE.exists():
        fallback = Path(__file__).parent / "combined_df.csv"
        if fallback.exists():
            return fallback
        fallback2 = Path(__file__).parent / "pbs_ex_manufacturer_combined.csv"
        if fallback2.exists():
            return fallback2
    return DATA_FILE


@st.cache_data(show_spinner="Loading dataset… (this may take a minute for large files)")
def load_data() -> pd.DataFrame:
    path = resolve_data_file()
    if not path.exists():
        st.error(
            f"Data file not found: `{path}`\n\n"
            "Edit `DATA_FILE` at the top of `pbs_dashboard.py` to point to your CSV."
        )
        return pd.DataFrame()

    # Detect encoding — try UTF-8 first, fall back to Windows-1252 (common for
    # files saved by Excel or R on Windows), then Latin-1 as last resort.
    def detect_encoding(p):
        for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(p, encoding=enc) as f:
                    f.read(65536)   # read a chunk to confirm
                return enc
            except (UnicodeDecodeError, LookupError):
                pass
        return "latin-1"   # latin-1 never raises, accepts any byte

    # Parquet files are pre-filtered and much smaller — load directly
    if path.suffix.lower() == ".parquet":
        df = pd.read_parquet(path)
    else:
        enc = detect_encoding(path)
        WANTED = ["price_date", "source", "item_code", "drug_name", "brand_name",
                  "form", "aemp", "dpmq_dpma", "formulary", "atc"]
        header = pd.read_csv(path, nrows=0, encoding=enc, encoding_errors="replace")
        available = [c for c in WANTED if c in header.columns]
        df = pd.read_csv(path, usecols=available, low_memory=False,
                         encoding=enc, encoding_errors="replace")
    df["price_date"] = pd.to_datetime(df["price_date"], errors="coerce")
    for col in ["aemp", "dpmq_dpma"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["price_date"])

    # Normalise source labels — the original R dataset split the non-EFC file
    # into two parts (non_efc1 / non_efc2) because it was too large.  Collapse
    # these back to a single "non_efc" label so each PBS item only appears once.
    if "source" in df.columns:
        df["source"] = (
            df["source"]
            .str.strip()
            .str.lower()
            .replace({
                "non_efc1": "non_efc",
                "non_efc2": "non_efc",
                "nonefc1":  "non_efc",
                "nonefc2":  "non_efc",
                "nonefc":   "non_efc",
            })
        )

    # Normalise text case so the same drug doesn't appear as
    # USTEKINUMAB / Ustekinumab / ustekinumab etc.
    if "drug_name" in df.columns:
        df["drug_name"] = df["drug_name"].str.strip().str.title()
    if "brand_name" in df.columns:
        df["brand_name"] = df["brand_name"].str.strip().str.title()

    # Convert high-cardinality string columns to categoricals — reduces memory
    # usage by ~70% since pandas stores one copy of each unique string instead
    # of repeating it across millions of rows.
    CAT_COLS = ["source", "item_code", "drug_name", "brand_name",
                "form", "formulary", "atc"]
    for col in CAT_COLS:
        if col in df.columns:
            df[col] = df[col].astype("category")

    df = df.sort_values("price_date")
    return df

def get_latest_date(df: pd.DataFrame) -> date:
    return df["price_date"].max().date()

def next_month(d: date) -> date:
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)

def today_first() -> date:
    t = date.today()
    return date(t.year, t.month, 1)

# ─── Check for new PBS data ───────────────────────────────────────────────────

def check_for_new_data(latest_date: date) -> list:
    new_records = []
    check_start = next_month(latest_date)
    check_end = today_first()
    if check_start > check_end:
        return new_records
    months_to_check = list(month_range(check_start, check_end))
    progress_bar = st.progress(0, text="Checking PBS website for new data…")
    for i, (yr, mo) in enumerate(months_to_check):
        progress_bar.progress((i + 1) / len(months_to_check), text=f"Checking {yr}-{mo:02d}…")
        d = date(yr, mo, 1)
        time.sleep(0.2)
        if d < SPLIT_DATE:
            data, url = try_download(urls_pre_split(yr, mo))
            if data:
                new_records.append((yr, mo, "pre_split", data, url))
        else:
            data, url = try_download(urls_efc(yr, mo))
            if data:
                new_records.append((yr, mo, "efc", data, url))
            time.sleep(0.2)
            data, url = try_download(urls_non_efc(yr, mo))
            if data:
                new_records.append((yr, mo, "non_efc", data, url))
    progress_bar.empty()
    return new_records

def append_new_data(new_records: list) -> int:
    all_new_frames = []
    for yr, mo, source, data, url in new_records:
        d = date(yr, mo, 1)
        df = read_xlsx_bytes(data, d, source)
        if df is not None:
            df = normalise_and_extract(df, source)
            all_new_frames.append(df)
    if not all_new_frames:
        return 0
    new_df = pd.concat(all_new_frames, ignore_index=True, sort=False)
    existing = pd.read_csv(DATA_FILE, low_memory=False)
    combined = pd.concat([existing, new_df], ignore_index=True, sort=False)
    combined.sort_values(["drug_name", "item_code", "price_date"], inplace=True)
    combined.reset_index(drop=True, inplace=True)
    combined.to_csv(DATA_FILE, index=False, encoding="utf-8")
    return len(new_df)

# ─── Analysis helpers ─────────────────────────────────────────────────────────

def find_f1_to_f2_transitions(item_df: pd.DataFrame) -> list:
    """Return list of (item_code, source, date) where formulary changed F1→F2."""
    transitions = []
    for (code, src), grp in item_df.groupby(["item_code", "source"]):
        grp = grp.sort_values("price_date").dropna(subset=["formulary"])
        if grp.empty:
            continue
        prev_formulary = None
        for _, row in grp.iterrows():
            f = str(row["formulary"]).strip()
            if prev_formulary == "F1" and f == "F2":
                transitions.append({
                    "item_code": code,
                    "source": src,
                    "date": row["price_date"],
                    "brand": row.get("brand_name", ""),
                })
            if f in ("F1", "F2"):
                prev_formulary = f
    return transitions

def find_price_changes(item_df: pd.DataFrame, col: str) -> pd.DataFrame:
    """
    For each (item_code, source) group, compute month-on-month change in col.
    Returns rows where price changed by more than 0.01% (either direction).
    """
    rows = []
    for (code, src), grp in item_df.groupby(["item_code", "source"]):
        grp = grp.sort_values("price_date").dropna(subset=[col]).copy()
        if len(grp) < 2:
            continue
        grp["_prev"] = grp[col].shift(1)
        grp["_pct"] = (grp[col] - grp["_prev"]) / grp["_prev"] * 100
        changes = grp[grp["_pct"].abs() > 0.01].copy()
        changes["item_code"] = code
        changes["source"] = src
        rows.append(changes[["price_date", "item_code", "source", col, "_prev", "_pct"]])
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)

# ─── Chart builder ────────────────────────────────────────────────────────────

COLORS = [
    "#2196F3", "#F44336", "#4CAF50", "#FF9800", "#9C27B0",
    "#00BCD4", "#FF5722", "#795548", "#607D8B", "#E91E63",
    "#3F51B5", "#009688", "#CDDC39", "#FFC107", "#8BC34A",
]

def price_chart(drug_df: pd.DataFrame, drug_label: str,
                show_pct_decreases: bool = True,
                show_f1f2: bool = True) -> go.Figure:
    """Dual-panel AEMP / DPMQ chart with F1→F2 annotations and % decrease labels."""

    # ── Collapse duplicate sources for the same item ──────────────────────────
    # The same PBS item code can appear in both EFC and non-EFC files.
    # Prefer EFC rows; for each (item_code, price_date) keep only one row so the
    # chart draws a single line per item regardless of source file origin.
    SOURCE_PRIORITY = {"efc": 0, "non_efc": 1, "pre_split": 2}
    drug_df = drug_df.copy()
    drug_df["_src_rank"] = drug_df["source"].map(SOURCE_PRIORITY).fillna(9)
    drug_df = (
        drug_df.sort_values(["price_date", "_src_rank"])
        .drop_duplicates(subset=["item_code", "price_date"], keep="first")
        .drop(columns=["_src_rank"])
    )

    # One legend entry per item_code — use most recent brand_name
    items = (
        drug_df.sort_values("price_date")
        .groupby("item_code", as_index=False)["brand_name"]
        .last()
        .sort_values("item_code")
    )
    item_color = {
        r["item_code"]: COLORS[i % len(COLORS)]
        for i, (_, r) in enumerate(items.iterrows())
    }
    item_label = {
        r["item_code"]: f"{r['item_code']} – {r['brand_name']}"
        for _, r in items.iterrows()
    }

    has_aemp = drug_df["aemp"].notna().any()
    has_dpmq = drug_df["dpmq_dpma"].notna().any()

    if has_aemp and has_dpmq:
        fig = make_subplots(
            rows=2, cols=1, shared_xaxes=True,
            subplot_titles=("AEMP (Ex-Manufacturer Price)", "DPMQ / DPMA"),
            vertical_spacing=0.10,
        )
    elif has_aemp:
        fig = make_subplots(rows=1, cols=1, subplot_titles=("AEMP",))
    else:
        fig = make_subplots(rows=1, cols=1, subplot_titles=("DPMQ / DPMA",))

    dpmq_row = 2 if (has_aemp and has_dpmq) else 1

    # ── Price change data (increases & decreases) ─────────────────────────────
    aemp_drops = find_price_changes(drug_df, "aemp") if (show_pct_decreases and has_aemp) else pd.DataFrame()
    dpmq_drops = find_price_changes(drug_df, "dpmq_dpma") if (show_pct_decreases and has_dpmq) else pd.DataFrame()

    # ── F1→F2 transitions ─────────────────────────────────────────────────────
    f1f2 = find_f1_to_f2_transitions(drug_df) if show_f1f2 else []

    # ── Traces ────────────────────────────────────────────────────────────────
    for _, row in items.iterrows():
        code = row["item_code"]
        color = item_color[code]
        label = item_label[code]
        sub = drug_df[drug_df["item_code"] == code].copy()
        sub = sub.sort_values("price_date")

        if has_aemp and sub["aemp"].notna().any():
            fig.add_trace(
                go.Scatter(
                    x=sub["price_date"], y=sub["aemp"],
                    name=label, line=dict(color=color, width=2),
                    mode="lines+markers", marker=dict(size=4),
                    hovertemplate=(
                        "<b>%{fullData.name}</b><br>"
                        "Date: %{x|%b %Y}<br>AEMP: $%{y:,.2f}<extra></extra>"
                    ),
                    legendgroup=label,
                ),
                row=1, col=1,
            )

        if has_dpmq and sub["dpmq_dpma"].notna().any():
            fig.add_trace(
                go.Scatter(
                    x=sub["price_date"], y=sub["dpmq_dpma"],
                    name=label, line=dict(color=color, width=2),
                    mode="lines+markers", marker=dict(size=4),
                    hovertemplate=(
                        "<b>%{fullData.name}</b><br>"
                        "Date: %{x|%b %Y}<br>DPMQ/DPMA: $%{y:,.2f}<extra></extra>"
                    ),
                    legendgroup=label,
                    showlegend=not has_aemp,
                ),
                row=dpmq_row, col=1,
            )

    # ── % price change annotations (green = increase, red = decrease) ─────────
    def add_change_annotations(changes_df: pd.DataFrame, val_col: str, panel_row: int):
        if changes_df.empty:
            return
        for _, dr in changes_df.iterrows():
            code, src = dr["item_code"], dr["source"]
            pct = dr["_pct"]
            val = dr[val_col]
            is_decrease = pct < 0
            ann_color = "#D32F2F" if is_decrease else "#2E7D32"   # red / green
            arrow_dir = -36 if is_decrease else 36                 # below / above
            fig.add_annotation(
                x=dr["price_date"],
                y=val,
                text=f"<b>{pct:+.1f}%</b>",
                showarrow=True,
                arrowhead=2,
                arrowsize=1,
                arrowwidth=1.5,
                arrowcolor=ann_color,
                ax=0,
                ay=arrow_dir,
                font=dict(size=10, color=ann_color),
                bgcolor="rgba(255,255,255,0.88)",
                bordercolor=ann_color,
                borderwidth=1,
                borderpad=3,
                row=panel_row, col=1,
            )

    if not aemp_drops.empty:
        add_change_annotations(aemp_drops, "aemp", 1)
    if not dpmq_drops.empty:
        add_change_annotations(dpmq_drops, "dpmq_dpma", dpmq_row)

    # ── F1→F2 vertical lines ──────────────────────────────────────────────────
    for i, t in enumerate(f1f2):
        xval = t["date"]
        brand = t.get("brand", "")
        label_text = f"F1→F2<br>{brand}" if brand else "F1→F2"
        color = item_color.get(t["item_code"], "#E91E63")

        # Add vline on both panels
        for r in ([1, dpmq_row] if has_aemp and has_dpmq else [1]):
            fig.add_vline(
                x=xval,
                line=dict(color=color, width=1.5, dash="dot"),
                row=r, col=1,
            )
        # Annotation on top panel only
        fig.add_annotation(
            x=xval,
            yref="paper",
            y=1.0,
            text=label_text,
            showarrow=False,
            font=dict(size=9, color=color),
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor=color,
            borderwidth=1,
            borderpad=3,
            xanchor="left",
            yanchor="top",
        )

    # ── Layout ────────────────────────────────────────────────────────────────
    chart_h = 650 if (has_aemp and has_dpmq) else 400
    fig.update_layout(
        title=dict(text=f"Price History: {drug_label}", font=dict(size=18)),
        height=chart_h,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="left", x=0),
        template="plotly_white",
        margin=dict(l=70, r=40, t=70, b=140),
    )
    fig.update_yaxes(tickprefix="$", tickformat=",.2f")
    return fig

# ─── Export helpers ───────────────────────────────────────────────────────────

def fig_to_png(fig: go.Figure) -> bytes:
    """Render figure to PNG bytes using kaleido."""
    return fig.to_image(format="png", width=1400, height=800, scale=2)

def fig_to_svg(fig: go.Figure) -> bytes:
    return fig.to_image(format="svg", width=1400, height=800)

# ─── Misc ─────────────────────────────────────────────────────────────────────

def drug_search(df: pd.DataFrame, query: str) -> pd.DataFrame:
    q = query.strip().upper()
    mask = (
        df["drug_name"].str.upper().str.contains(q, na=False) |
        df["item_code"].str.upper().str.contains(q, na=False) |
        df["brand_name"].str.upper().str.contains(q, na=False)
    )
    return df[mask]

def summary_stats(drug_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (code, src), grp in drug_df.groupby(["item_code", "source"]):
        grp_s = grp.sort_values("price_date")
        brand = grp_s["brand_name"].dropna().iloc[-1] if grp_s["brand_name"].notna().any() else "—"
        latest_row = grp_s.iloc[-1]
        earliest_row = grp_s.iloc[0]
        current_formulary = latest_row.get("formulary", "—")
        row = {
            "Item Code": code, "Brand": brand, "Source": src.upper(),
            "First Date": earliest_row["price_date"].strftime("%b %Y"),
            "Latest Date": latest_row["price_date"].strftime("%b %Y"),
            "Formulary": current_formulary if pd.notna(current_formulary) else "—",
        }
        if grp_s["aemp"].notna().any():
            first_v = grp_s[grp_s["aemp"].notna()].iloc[0]["aemp"]
            last_v = latest_row["aemp"]
            row["AEMP (first)"] = f"${first_v:,.2f}" if pd.notna(first_v) else "—"
            row["AEMP (latest)"] = f"${last_v:,.2f}" if pd.notna(last_v) else "—"
            if pd.notna(first_v) and pd.notna(last_v) and first_v > 0:
                pct = (last_v - first_v) / first_v * 100
                row["AEMP Δ%"] = f"{pct:+.1f}%"
            else:
                row["AEMP Δ%"] = "—"
        if grp_s["dpmq_dpma"].notna().any():
            first_v = grp_s[grp_s["dpmq_dpma"].notna()].iloc[0]["dpmq_dpma"]
            last_v = latest_row["dpmq_dpma"]
            row["DPMQ/DPMA (first)"] = f"${first_v:,.2f}" if pd.notna(first_v) else "—"
            row["DPMQ/DPMA (latest)"] = f"${last_v:,.2f}" if pd.notna(last_v) else "—"
            if pd.notna(first_v) and pd.notna(last_v) and first_v > 0:
                pct = (last_v - first_v) / first_v * 100
                row["DPMQ/DPMA Δ%"] = f"{pct:+.1f}%"
            else:
                row["DPMQ/DPMA Δ%"] = "—"
        rows.append(row)
    return pd.DataFrame(rows)

# ─── Main app ─────────────────────────────────────────────────────────────────

LOGO_PATH = Path(r"C:\Users\bagre\OneDrive\Documentos\psd data\psd_search_app_r\www\IQVIA-Logo.png")


def check_password() -> bool:
    """Show a login form; return True only when credentials are correct."""
    if st.session_state.get("authenticated"):
        return True

    # Centre the login card
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=180)
        st.title("PBS Ex-Manufacturer Price Dashboard")
        st.caption("Please log in to continue.")
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in", use_container_width=True)
        if submitted:
            # Credentials come from .streamlit/secrets.toml locally,
            # or the Streamlit Cloud secrets manager when deployed.
            valid_user = st.secrets.get("auth", {}).get("username", "iqvia")
            valid_pass = st.secrets.get("auth", {}).get("password", "iqvia")
            if username == valid_user and password == valid_pass:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect username or password — please try again.")
    return False


def main():
    if not check_password():
        st.stop()

    # ── Header row: title left, IQVIA logo right ──────────────────────────────
    title_col, logo_col = st.columns([6, 1])
    with title_col:
        st.title("💊 PBS Ex-Manufacturer Price Dashboard")
        st.caption("Pharmaceutical Benefits Scheme — Ex-Manufacturer Pricing Monitor")
    with logo_col:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=130)
    # ─────────────────────────────────────────────────────────────────────────

    df = load_data()
    if df.empty:
        st.stop()

    latest_date = get_latest_date(df)

    # Build sorted autocomplete list once: "DRUG NAME (Brand1; Brand2)" entries + item codes
    @st.cache_data
    def build_search_options(_df):
        # Collect all unique brand names per drug and join with "; "
        # e.g. "Ustekinumab  (Stelara; Steqeyma)"
        def join_brands(s):
            brands = sorted({str(b).strip() for b in s if pd.notna(b) and str(b).strip()})
            return "; ".join(brands) if brands else ""

        drug_brands = (
            _df.dropna(subset=["drug_name"])
            .groupby("drug_name", as_index=False)["brand_name"]
            .agg(join_brands)
        )
        drug_opts = []
        for _, r in drug_brands.iterrows():
            if r["brand_name"]:
                drug_opts.append(f"{r['drug_name']}  ({r['brand_name']})")
            else:
                drug_opts.append(r["drug_name"])
        # Also add unique item codes
        item_opts = sorted(_df["item_code"].dropna().unique().tolist())
        return sorted(drug_opts), item_opts

    drug_opts, item_opts = build_search_options(df)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Latest Data", latest_date.strftime("%b %Y"))
    col2.metric("Total Records", f"{len(df):,}")
    col3.metric("Unique Drugs", f"{df['drug_name'].nunique():,}")
    col4.metric("Months Covered", f"{df['price_date'].nunique()}")

    st.divider()

    tab_dash, tab_update = st.tabs(["📈 Price Dashboard", "🔄 Check for New Data"])

    # ════════════════════════════════════════════════════════════════════════
    with tab_dash:
        st.subheader("Search for a Drug")
        # Selectbox with built-in search — type to filter the dropdown
        all_opts = [""] + drug_opts + item_opts
        selected_opt = st.selectbox(
            "Drug name, brand name, or item code",
            options=all_opts,
            index=0,
            placeholder="Start typing to search…",
            format_func=lambda x: "Start typing to search…" if x == "" else x,
        )
        # Extract the plain drug name / item code from the selection
        if selected_opt and "  (" in selected_opt:
            query = selected_opt.split("  (")[0].strip()
        else:
            query = selected_opt.strip()

        if not query:
            st.info("👆 Enter a drug name, brand name, or item code to get started.")
            st.subheader("Top 10 Highest AEMP (latest month)")
            latest_month = df["price_date"].max()
            latest_slice = df[(df["price_date"] == latest_month) & df["aemp"].notna()]

            # Deduplicate: group by drug + brand + aemp value so the same drug
            # with multiple item codes at the same price only appears once.
            # If a drug has item codes at DIFFERENT prices, each price is kept.
            top10 = (
                latest_slice
                .groupby(["drug_name", "brand_name", "aemp"], as_index=False)
                .agg(
                    item_codes=("item_code", lambda x: ", ".join(sorted(x.unique()))),
                    dpmq_dpma=("dpmq_dpma", "first"),
                )
                .nlargest(10, "aemp")
                .reset_index(drop=True)
            )
            top10.index += 1
            top10 = top10.rename(columns={
                "drug_name": "Drug", "brand_name": "Brand",
                "item_codes": "Item Code(s)", "aemp": "AEMP ($)", "dpmq_dpma": "DPMQ/DPMA ($)",
            })
            top10["AEMP ($)"] = top10["AEMP ($)"].map("${:,.2f}".format)
            top10["DPMQ/DPMA ($)"] = top10["DPMQ/DPMA ($)"].map(
                lambda v: f"${v:,.2f}" if pd.notna(v) else "—"
            )
            st.dataframe(top10, use_container_width=True, hide_index=False)

            st.divider()

            # ── Helper: deduplicate by drug+brand+price, return top/bottom N ──
            def make_price_table(slice_df, n=10, largest=True):
                grouped = (
                    slice_df
                    .groupby(["drug_name", "brand_name", "aemp"], as_index=False)
                    .agg(
                        item_codes=("item_code", lambda x: ", ".join(sorted(x.unique()))),
                        dpmq_dpma=("dpmq_dpma", "first"),
                    )
                )
                ranked = grouped.nlargest(n, "aemp") if largest else grouped.nsmallest(n, "aemp")
                ranked = ranked.reset_index(drop=True)
                ranked.index += 1
                ranked = ranked.rename(columns={
                    "drug_name": "Drug", "brand_name": "Brand",
                    "item_codes": "Item Code(s)", "aemp": "AEMP ($)", "dpmq_dpma": "DPMQ/DPMA ($)",
                })
                ranked["AEMP ($)"] = ranked["AEMP ($)"].map("${:,.2f}".format)
                ranked["DPMQ/DPMA ($)"] = ranked["DPMQ/DPMA ($)"].map(
                    lambda v: f"${v:,.2f}" if pd.notna(v) else "—"
                )
                return ranked

            # ── Top 10 Lowest AEMP ────────────────────────────────────────────
            st.subheader("Top 10 Lowest AEMP (latest month)")
            bottom10 = make_price_table(latest_slice, n=10, largest=False)
            st.dataframe(bottom10, use_container_width=True, hide_index=False)

            st.divider()

            # ── Newest Drugs — first appearance IS the latest month ────────────
            # Only show drugs whose sole month of data is the dataset's most recent month.
            # (Excludes drugs that appeared once years ago and dropped off the PBS.)
            newest_label = latest_month.strftime("%B %Y")
            st.subheader(f"Newest Drugs on PBS ({newest_label} additions)")

            # Drugs that appear in the latest month
            in_latest = set(
                df[df["price_date"] == latest_month][["drug_name", "brand_name"]]
                .drop_duplicates()
                .apply(tuple, axis=1)
            )
            # Drugs whose FIRST ever appearance is the latest month
            first_seen = (
                df.groupby(["drug_name", "brand_name"], as_index=False)["price_date"].min()
                .rename(columns={"price_date": "first_date"})
            )
            truly_new = first_seen[first_seen["first_date"] == latest_month].copy()

            # Pull AEMP / DPMQ / item code from the latest month slice
            latest_detail = (
                latest_slice
                .groupby(["drug_name", "brand_name"], as_index=False)
                .agg(
                    item_codes=("item_code", lambda x: ", ".join(sorted(x.unique()))),
                    aemp=("aemp", "first"),
                    dpmq_dpma=("dpmq_dpma", "first"),
                )
            )
            newest = truly_new.merge(latest_detail, on=["drug_name", "brand_name"], how="left")
            newest = newest.sort_values("drug_name").reset_index(drop=True)
            newest.index += 1
            n_newest = len(newest)
            newest["aemp"] = newest["aemp"].map(lambda v: f"${v:,.2f}" if pd.notna(v) else "—")
            newest["dpmq_dpma"] = newest["dpmq_dpma"].map(lambda v: f"${v:,.2f}" if pd.notna(v) else "—")
            newest = newest.rename(columns={
                "drug_name": "Drug", "brand_name": "Brand", "item_codes": "Item Code(s)",
                "aemp": "AEMP ($)", "dpmq_dpma": "DPMQ/DPMA ($)",
            })[["Drug", "Brand", "Item Code(s)", "AEMP ($)", "DPMQ/DPMA ($)"]]
            st.caption(f"{n_newest} drug{'s' if n_newest != 1 else ''} appearing for the first time in {newest_label}")
            st.dataframe(newest, use_container_width=True, hide_index=False)

        else:
            results = drug_search(df, query)

            if results.empty:
                st.warning(f"No results found for **{query}**.")
            else:
                unique_drugs = results[["drug_name"]].drop_duplicates().sort_values("drug_name")

                if len(unique_drugs) > 1:
                    drug_choice = st.selectbox(
                        f"Found {len(unique_drugs)} drugs — select one:",
                        options=unique_drugs["drug_name"].tolist(),
                    )
                    drug_df = results[results["drug_name"] == drug_choice].copy()
                else:
                    drug_choice = unique_drugs.iloc[0]["drug_name"]
                    drug_df = results.copy()

                # ── Date range filter ─────────────────────────────────────
                min_d = drug_df["price_date"].min().date()
                max_d = drug_df["price_date"].max().date()
                date_range = st.slider(
                    "Date range", min_value=min_d, max_value=max_d,
                    value=(min_d, max_d), format="MMM YYYY",
                )
                drug_df = drug_df[
                    (drug_df["price_date"].dt.date >= date_range[0]) &
                    (drug_df["price_date"].dt.date <= date_range[1])
                ]

                # ── Item filter ───────────────────────────────────────────
                # Deduplicate by item_code only for filter display
                all_items = (
                    drug_df.sort_values("price_date")
                    .groupby("item_code", as_index=False)["brand_name"].last()
                )
                if len(all_items) > 6:
                    item_opts_filt = all_items.apply(
                        lambda r: f"{r['item_code']} – {r['brand_name']}", axis=1
                    ).tolist()
                    selected_items = st.multiselect(
                        "Filter by item/brand (leave empty for all)", options=item_opts_filt, default=[],
                    )
                    if selected_items:
                        selected_codes = [s.split(" – ")[0] for s in selected_items]
                        drug_df = drug_df[drug_df["item_code"].isin(selected_codes)]

                # ── Chart options ─────────────────────────────────────────
                opt_col1, opt_col2 = st.columns(2)
                show_pct = opt_col1.checkbox("Show % price changes on chart", value=True)
                show_f1f2 = opt_col2.checkbox("Show F1→F2 transitions", value=True)

                # ── F1→F2 info banner ─────────────────────────────────────
                if show_f1f2:
                    transitions = find_f1_to_f2_transitions(drug_df)
                    if transitions:
                        # Group by transition date — one clean banner per date
                        by_date = {}
                        for t in transitions:
                            key = t["date"].strftime("%B %Y")
                            by_date.setdefault(key, t)   # keep first (any brand will do)
                        for date_label, t in by_date.items():
                            brand = t.get("brand", "")
                            st.info(
                                f"🔄 **{drug_choice}**"
                                + (f" ({brand})" if brand else "")
                                + f" moved from **F1 → F2** in **{date_label}**"
                            )

                # ── Build & display chart ─────────────────────────────────
                fig = price_chart(drug_df, drug_choice, show_pct_decreases=show_pct, show_f1f2=show_f1f2)
                st.plotly_chart(fig, use_container_width=True)

                # ── Export buttons ────────────────────────────────────────
                exp_col1, exp_col2, exp_col3 = st.columns([1, 1, 4])
                safe_name = drug_choice.replace(" ", "_").lower()

                with exp_col1:
                    try:
                        png_bytes = fig_to_png(fig)
                        st.download_button(
                            label="📷 Export PNG",
                            data=png_bytes,
                            file_name=f"pbs_{safe_name}.png",
                            mime="image/png",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.caption(f"PNG export unavailable: {e}")

                with exp_col2:
                    # Interactive HTML export (always works, no kaleido needed)
                    html_str = fig.to_html(include_plotlyjs="cdn", full_html=True)
                    st.download_button(
                        label="🌐 Export HTML",
                        data=html_str.encode(),
                        file_name=f"pbs_{safe_name}.html",
                        mime="text/html",
                        use_container_width=True,
                    )

                # ── Summary table ─────────────────────────────────────────
                with st.expander("📊 Summary Statistics", expanded=True):
                    stats = summary_stats(drug_df)
                    st.dataframe(stats, use_container_width=True, hide_index=True)

                # ── Price change events table ─────────────────────────────
                aemp_changes = find_price_changes(drug_df, "aemp")
                dpmq_changes = find_price_changes(drug_df, "dpmq_dpma")
                if not aemp_changes.empty or not dpmq_changes.empty:
                    with st.expander("📊 Price Change Events"):
                        def fmt_change_table(changes_df, price_col, label):
                            if changes_df.empty:
                                return
                            st.markdown(f"**{label}**")
                            disp = changes_df.copy()
                            disp["Date"] = disp["price_date"].dt.strftime("%b %Y")
                            disp["Before"] = disp["_prev"].map("${:,.2f}".format)
                            disp["After"]  = disp[price_col].map("${:,.2f}".format)
                            disp["Change"] = disp["_pct"].map("{:+.2f}%".format)
                            disp["Direction"] = disp["_pct"].map(
                                lambda p: "▼ Decrease" if p < 0 else "▲ Increase"
                            )
                            st.dataframe(
                                disp[["Date", "item_code", "Before", "After", "Change", "Direction"]]
                                .rename(columns={"item_code": "Item"}),
                                use_container_width=True, hide_index=True,
                            )
                        fmt_change_table(aemp_changes, "aemp", "AEMP Changes")
                        fmt_change_table(dpmq_changes, "dpmq_dpma", "DPMQ/DPMA Changes")

                # ── Raw data ──────────────────────────────────────────────
                with st.expander("🗃 Raw Data"):
                    show_df = drug_df.drop(columns=["atc", "previous_aemp", "price_change_event"], errors="ignore").copy()
                    show_df["price_date"] = show_df["price_date"].dt.strftime("%Y-%m-%d")
                    st.dataframe(show_df, use_container_width=True, hide_index=True)
                    csv_bytes = show_df.to_csv(index=False).encode()
                    st.download_button(
                        "⬇ Download as CSV", data=csv_bytes,
                        file_name=f"pbs_{safe_name}.csv", mime="text/csv",
                    )

    # ════════════════════════════════════════════════════════════════════════
    with tab_update:
        st.subheader("Update Dataset")
        next_expected = next_month(latest_date)
        today_month = today_first()
        active_file = resolve_data_file()

        col_a, col_b = st.columns(2)
        col_a.metric("Latest month in dataset", latest_date.strftime("%B %Y"))
        col_b.metric("Next expected month", next_expected.strftime("%B %Y"))

        if next_expected > today_month:
            st.success("✅ Your dataset is already up to date.")
        else:
            months_behind = sum(1 for _ in month_range(next_expected, today_month))
            st.warning(
                f"⚠️ Dataset may be up to **{months_behind} month(s)** behind. "
                "Run the update script below to fetch the latest data."
            )

        st.divider()
        st.subheader("How to update")

        st.markdown(
            "Run `update_pbs_data.py` from the command line. "
            "It will automatically download any missing months from the PBS website "
            "and append them to your CSV.\n\n"
            "**In PowerShell / Command Prompt:**"
        )
        csv_display = str(DATA_FILE).replace("\\", "/")
        st.code(f'python update_pbs_data.py', language="bash")
        st.markdown("Or, to re-download from a specific month:")
        st.code(f'python update_pbs_data.py --from 2026-03', language="bash")
        st.markdown(
            "After the script finishes, **reload this page** (press F5) to see the updated data."
        )

        st.divider()
        st.subheader("Cloud storage — how to set up")
        st.markdown("""
**Your file is already in OneDrive** — it syncs automatically whenever `update_pbs_data.py` saves changes.

To load from a direct cloud URL instead of a local path (useful if running this dashboard on another machine):

1. **OneDrive** — right-click the file → *Share* → *Copy link*, then add `?download=1` to the end.
   Example: `https://onedrive.live.com/download?cid=XXX&resid=YYY&authkey=ZZZ`

2. **Google Drive** — share the file publicly, then use:
   `https://drive.google.com/uc?export=download&id=FILE_ID`
   *(For files >100 MB, Google adds a virus-scan warning — use a direct share link instead.)*

3. **Dropbox** — change `?dl=0` → `?dl=1` at the end of the share link.

Once you have the URL, open `pbs_dashboard.py` and set:
```python
CLOUD_URL = "https://your-direct-download-url-here"
```
The app will download and cache the file on first load, then reload from cache on subsequent runs.
        """)

        st.divider()
        st.subheader("Dataset File Info")
        if active_file.exists():
            size_mb = active_file.stat().st_size / 1_048_576
            st.write(f"📁 **File:** `{active_file}`")
            st.write(f"📦 **Size:** {size_mb:.1f} MB")
            st.write(f"📅 **Last modified:** {datetime.fromtimestamp(active_file.stat().st_mtime).strftime('%d %B %Y %H:%M')}")
        else:
            st.error(f"File not found: `{active_file}`")


if __name__ == "__main__":
    main()
