#!/usr/bin/env python3
"""
HFBC_Praise_Worship – KPI ETL
Reads raw spreadsheets/CSVs, validates, applies business rules,
and writes curated KPI CSVs to /data for the dashboard.

Requires: pandas, openpyxl
"""

from __future__ import annotations
import sys
from pathlib import Path
from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo
import pandas as pd

# -------- CONFIG --------
REPO_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR   = REPO_ROOT / "data"
SETLIST_XLSX = REPO_ROOT / "setlist" / "setlist.xlsx"        # Required
SONGS_CATALOG_CSV = REPO_ROOT / "setlist" / "songs_catalog.csv"  # Required

# Business rules
TZ = ZoneInfo("America/Chicago")
ROLL_HOUR = 12; ROLL_MIN = 30  # Sunday 12:30 pm rollover
WEEK_WINDOW = 52               # rolling weeks
EXCLUDE_TITLES = {"NA", "N/A", "Church Close", "Church Close - Flood"}  # extend as needed

# Column expectations (soft mapping → more robust to naming drifts)
SETLIST_COLMAP = {
    "date": {"date", "service_date", "Date", "DATE"},
    "title": {"title", "song", "Song", "Title"},
    "source": {"source", "category", "Category", "Source"},  # optional
}
SONGSCAT_COLMAP = {
    "song_number": {"number", "no", "song_no", "Song #", "Song_Number"},
    "title": {"title", "song", "Song", "Title"},
    "in_hymnal": {"in_hymnal", "hymnal", "In_Hymnal"},
}

# -------- UTIL --------
def pick_col(df: pd.DataFrame, candidates: set[str], required: bool, label: str) -> str:
    for c in df.columns:
        if c in candidates:
            return c
    # case-insensitive match
    lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    if required:
        raise ValueError(f"Required column '{label}' not found. Have: {list(df.columns)}")
    return None

def sunday_rollover_cutoff(now_ct: datetime) -> datetime:
    # Find the most recent Sunday 12:30 pm CT
    # If now is before this week’s Sunday 12:30, use last week’s.
    weekday = now_ct.weekday()  # Monday=0 ... Sunday=6
    # days since Sunday
    days_since_sun = (weekday + 1) % 7
    this_sunday = (now_ct - timedelta(days=days_since_sun)).replace(
        hour=ROLL_HOUR, minute=ROLL_MIN, second=0, microsecond=0
    )
    if now_ct >= this_sunday:
        return this_sunday
    else:
        return this_sunday - timedelta(days=7)

def start_of_window(cutoff: datetime) -> datetime:
    return cutoff - timedelta(weeks=WEEK_WINDOW)

def normalize_titles(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()

def load_setlist() -> pd.DataFrame:
    if not SETLIST_XLSX.exists():
        raise FileNotFoundError(f"Missing file: {SETLIST_XLSX}")
    df = pd.read_excel(SETLIST_XLSX)
    c_date = pick_col(df, SETLIST_COLMAP["date"], True, "date")
    c_title = pick_col(df, SETLIST_COLMAP["title"], True, "title")
    c_source = pick_col(df, SETLIST_COLMAP["source"], False, "source")

    out = pd.DataFrame({
        "Date": pd.to_datetime(df[c_date], errors="coerce"),
        "Title": normalize_titles(df[c_title]),
    })
    if c_source:
        out["Source"] = normalize_titles(df[c_source])
    else:
        out["Source"] = "Unknown"

    # Validation
    if out["Date"].isna().any():
        bad = out[out["Date"].isna()]
        raise ValueError(f"Unparseable dates in setlist.xlsx rows:\n{bad.to_string(index=False)}")
    if out["Title"].eq("").any():
        raise ValueError("Empty song titles found in setlist.xlsx")

    # Deduplicate per (Date, Title) to prevent double counting
    out = out.drop_duplicates(subset=["Date", "Title"]).reset_index(drop=True)
    return out

def load_songs_catalog() -> pd.DataFrame:
    if not SONGS_CATALOG_CSV.exists():
        raise FileNotFoundError(f"Missing file: {SONGS_CATALOG_CSV}")
    df = pd.read_csv(SONGS_CATALOG_CSV)
    c_num = pick_col(df, SONGSCAT_COLMAP["song_number"], False, "song_number")
    c_title = pick_col(df, SONGSCAT_COLMAP["title"], True, "title")
    c_inh = pick_col(df, SONGSCAT_COLMAP["in_hymnal"], False, "in_hymnal")

    out = pd.DataFrame({"Title": normalize_titles(df[c_title])})
    if c_num: out["Song_Number"] = df[c_num]
    if c_inh: out["In_Hymnal"] = df[c_inh].astype(str).str.lower().isin({"1","true","yes","y"})
    else:     out["In_Hymnal"] = True  # assume from hymnal unless specified
    out = out.drop_duplicates(subset=["Title"]).reset_index(drop=True)
    return out

def filter_business_rules(df: pd.DataFrame, now_ct: datetime) -> pd.DataFrame:
    # Apply Sunday 12:30 pm CT window
    cutoff = sunday_rollover_cutoff(now_ct)
    start = start_of_window(cutoff)
    # Window filter
    dfw = df[(df["Date"] >= start) & (df["Date"] < cutoff)].copy()

    # Exclusions by Title (case-insensitive)
    excl_norm = {t.strip().lower() for t in EXCLUDE_TITLES}
    dfw = dfw[~dfw["Title"].str.lower().isin(excl_norm)].copy()
    return dfw

# -------- KPI COMPUTATIONS --------
def kpi_top10(dfw: pd.DataFrame) -> pd.DataFrame:
    counts = dfw.groupby("Title").size().reset_index(name="Plays").sort_values("Plays", ascending=False)
    return counts.head(10)

def kpi_by_source(dfw: pd.DataFrame) -> pd.DataFrame:
    counts = dfw.groupby(["Source", "Title"]).size().reset_index(name="Plays")
    # Aggregate per Source for a compact KPI
    kpi = counts.groupby("Source")["Plays"].sum().reset_index().sort_values("Plays", ascending=False)
    return kpi

def kpi_hymnal_coverage(dfw: pd.DataFrame, catalog: pd.DataFrame) -> pd.DataFrame:
    used_titles = dfw["Title"].drop_duplicates()
    cat = catalog[["Title", "In_Hymnal"]].copy()
    cat["Used"] = cat["Title"].isin(used_titles)
    # Coverage among hymnal songs
    denom = (cat["In_Hymnal"] == True).sum()
    num = ((cat["In_Hymnal"] == True) & (cat["Used"] == True)).sum()
    coverage = 0.0 if denom == 0 else round(100 * num / denom, 2)
    return pd.DataFrame([{"Hymnal_Songs": denom, "Used_Hymnal_Songs": num, "Coverage_%": coverage}])

def kpi_hymnal_unused(dfw: pd.DataFrame, catalog: pd.DataFrame) -> pd.DataFrame:
    used = set(dfw["Title"].unique())
    hymnal = catalog[catalog["In_Hymnal"] == True].copy()
    hymnal["Used"] = hymnal["Title"].isin(used)
    unused = hymnal[hymnal["Used"] == False][["Song_Number", "Title"]].sort_values("Song_Number", na_index=False, na_position="last")
    return unused.reset_index(drop=True)

# -------- MAIN --------
def main():
    now_ct = datetime.now(TZ)
    print(f"[INFO] Running ETL at {now_ct.isoformat()}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    setlist = load_setlist()
    catalog = load_songs_catalog()

    # Window + rules
    dfw = filter_business_rules(setlist, now_ct)
    if dfw.empty:
        print("[WARN] No rows within the 52-week window after exclusions.")

    # KPIs
    top10 = kpi_top10(dfw)
    bysrc = kpi_by_source(dfw)
    cov   = kpi_hymnal_coverage(dfw, catalog)
    unused = kpi_hymnal_unused(dfw, catalog)

    # Outputs
    out_top10 = DATA_DIR / "kpi_top10.csv"
    out_bysrc = DATA_DIR / "kpi_by_source.csv"
    out_cov   = DATA_DIR / "kpi_hymnal_coverage.csv"
    out_unused= DATA_DIR / "hymnal_unused.csv"

    top10.to_csv(out_top10, index=False)
    bysrc.to_csv(out_bysrc, index=False)
    cov.to_csv(out_cov, index=False)
    unused.to_csv(out_unused, index=False)

    print(f"[OK] Wrote: {out_top10.relative_to(REPO_ROOT)}")
    print(f"[OK] Wrote: {out_bysrc.relative_to(REPO_ROOT)}")
    print(f"[OK] Wrote: {out_cov.relative_to(REPO_ROOT)}")
    print(f"[OK] Wrote: {out_unused.relative_to(REPO_ROOT)}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[FAIL] {e}", file=sys.stderr)
        sys.exit(1)

