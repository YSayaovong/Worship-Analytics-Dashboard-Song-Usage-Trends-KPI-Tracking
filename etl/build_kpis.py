#!/usr/bin/env python3
"""
HFBC_Praise_Worship – KPI ETL
Reproducible KPIs with Sunday 12:30 PM CT window and exclusions.
Requires: pandas, openpyxl
"""

from __future__ import annotations
import sys
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd

# ---------- CONFIG ----------
REPO_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = REPO_ROOT / "data"
SETLIST_XLSX = REPO_ROOT / "setlist" / "setlist.xlsx"
SONGS_CATALOG_CSV = REPO_ROOT / "setlist" / "songs_catalog.csv"

TZ = ZoneInfo("America/Chicago")
ROLL_HOUR, ROLL_MIN = 12, 30        # Sunday 12:30 PM CT
WEEK_WINDOW = 52
EXCLUDE_TITLES = {"NA", "N/A", "Church Close", "Church Close - Flood"}

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

# ---------- UTIL ----------
def pick_col(df: pd.DataFrame, candidates: set[str], required: bool, label: str) -> str:
    for c in df.columns:
        if c in candidates:
            return c
    lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    if required:
        raise ValueError(f"Required column '{label}' not found. Have: {list(df.columns)}")
    return None

def sunday_rollover_cutoff(now_ct: datetime) -> datetime:
    """Return tz-aware CT datetime for most recent Sunday 12:30 PM."""
    wd = now_ct.weekday()  # Mon=0 ... Sun=6
    days_since_sun = (wd + 1) % 7
    this_sun = (now_ct - timedelta(days=days_since_sun)).replace(
        hour=ROLL_HOUR, minute=ROLL_MIN, second=0, microsecond=0
    )
    return this_sun if now_ct >= this_sun else this_sun - timedelta(days=7)

def get_naive_window(now_ct: datetime) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Compute [start, cutoff) as tz-naive pandas Timestamps (to match Excel)."""
    cutoff_ct = sunday_rollover_cutoff(now_ct)         # tz-aware
    cutoff = pd.Timestamp(cutoff_ct).tz_localize(None) # strip tz -> naive
    start = cutoff - pd.Timedelta(weeks=WEEK_WINDOW)
    return start, cutoff

def normalize_titles(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()

# ---------- LOAD ----------
def load_setlist() -> pd.DataFrame:
    if not SETLIST_XLSX.exists():
        raise FileNotFoundError(f"Missing file: {SETLIST_XLSX}")
    df = pd.read_excel(SETLIST_XLSX)

    c_date = pick_col(df, SETLIST_COLMAP["date"], True, "date")
    c_title = pick_col(df, SETLIST_COLMAP["title"], True, "title")
    c_source = pick_col(df, SETLIST_COLMAP["source"], False, "source")

    out = pd.DataFrame({
        "Date": pd.to_datetime(df[c_date], errors="coerce"),  # -> datetime64[ns] (naive)
        "Title": normalize_titles(df[c_title]),
    })
    out["Source"] = normalize_titles(df[c_source]) if c_source else "Unknown"

    # drop bad rows; keep dataset naive (no tz)
    out = out.dropna(subset=["Date"])
    out = out[out["Title"].astype(str).str.len() > 0]
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
    out["Song_Number"] = df[c_num] if c_num else pd.NA
    out["In_Hymnal"] = df[c_inh].astype(str).str.lower().isin({"1", "true", "yes", "y"}) if c_inh else True
    return out.drop_duplicates(subset=["Title"]).reset_index(drop=True)

# ---------- RULES & KPIs ----------
def apply_window_and_rules(df: pd.DataFrame, now_ct: datetime) -> pd.DataFrame:
    start, cutoff = get_naive_window(now_ct)  # both tz-naive
    dfw = df[(df["Date"] >= start) & (df["Date"] < cutoff)].copy()
    excl = {t.strip().lower() for t in EXCLUDE_TITLES}
    dfw = dfw[~dfw["Title"].str.lower().isin(excl)]
    return dfw

def kpi_top10(dfw: pd.DataFrame) -> pd.DataFrame:
    if dfw.empty:
        return pd.DataFrame(columns=["Title", "Plays"])
    return (dfw.groupby("Title").size()
            .reset_index(name="Plays")
            .sort_values("Plays", ascending=False)
            .head(10))

def kpi_by_source(dfw: pd.DataFrame) -> pd.DataFrame:
    if dfw.empty:
        return pd.DataFrame(columns=["Source", "Plays"])
    return (dfw.groupby("Source").size()
            .reset_index(name="Plays")
            .sort_values("Plays", ascending=False))

def kpi_hymnal_coverage(dfw: pd.DataFrame, catalog: pd.DataFrame) -> pd.DataFrame:
    if catalog.empty:
        return pd.DataFrame([{"Hymnal_Songs": 0, "Used_Hymnal_Songs": 0, "Coverage_%": 0.0}])
    used = set(dfw["Title"].unique())
    cat = catalog[["Title", "In_Hymnal"]].copy()
    cat["Used"] = cat["Title"].isin(used)
    denom = int((cat["In_Hymnal"] == True).sum())
    num = int(((cat["In_Hymnal"] == True) & (cat["Used"] == True)).sum())
    cov = 0.0 if denom == 0 else round(100 * num / denom, 2)
    return pd.DataFrame([{"Hymnal_Songs": denom, "Used_Hymnal_Songs": num, "Coverage_%": cov}])

def kpi_hymnal_unused(dfw: pd.DataFrame, catalog: pd.DataFrame) -> pd.DataFrame:
    if catalog.empty:
        return pd.DataFrame(columns=["Song_Number", "Title"])
    used = set(dfw["Title"].unique())
    hymnal = catalog[catalog["In_Hymnal"] == True].copy()
    hymnal["Used"] = hymnal["Title"].isin(used)
    unused = hymnal[hymnal["Used"] == False][["Song_Number", "Title"]].sort_values(
        by="Song_Number", na_position="last"
    )
    return unused.reset_index(drop=True)

# ---------- MAIN ----------
def main():
    now_ct = datetime.now(TZ)
    print(f"[INFO] Running ETL at {now_ct.isoformat()}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    setlist = load_setlist()
    catalog = load_songs_catalog()

    dfw = apply_window_and_rules(setlist, now_ct)
    if dfw.empty:
        print("[WARN] No rows in last 52 weeks after exclusions.")

    top10 = kpi_top10(dfw)
    bysrc = kpi_by_source(dfw)
    cov   = kpi_hymnal_coverage(dfw, catalog)
    unused= kpi_hymnal_unused(dfw, catalog)

    (DATA_DIR / "kpi_top10.csv").write_text(top10.to_csv(index=False))
    (DATA_DIR / "kpi_by_source.csv").write_text(bysrc.to_csv(index=False))
    (DATA_DIR / "kpi_hymnal_coverage.csv").write_text(cov.to_csv(index=False))
    (DATA_DIR / "hymnal_unused.csv").write_text(unused.to_csv(index=False))

    print("[OK] data/*.csv refreshed")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[FAIL] {e}", file=sys.stderr)
        sys.exit(1)
