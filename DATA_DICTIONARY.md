# Data Dictionary – HFBC Praise & Worship

## Business Rules
- **Window:** Rolling 52 weeks ending at the most recent **Sunday 12:30 PM America/Chicago**.
- **Play definition:** A song **Title** appearing on a setlist date (deduped by `(Date, Title)`).
- **Exclusions:** Titles in `{"NA", "N/A", "Church Close", "Church Close - Flood"}` (case-insensitive).
- **Source:** Optional category/collection the song belongs to (e.g., Hymnal / Special / Other).

---

## Raw Sources

### `setlist/setlist.xlsx`
| Column (expected) | Type      | Required | Notes                                   |
|---|---|---|---|
| Date              | date      | Yes      | Service date. Parsed to UTC-agnostic date; windowed in CT. |
| Title             | string    | Yes      | Song title; trimmed, case-insensitive compare. |
| Source            | string    | No       | Category/source; default `"Unknown"` if missing. |

### `setlist/songs_catalog.csv`
| Column (expected) | Type   | Required | Notes                                          |
|---|---|---|---|
| Song_Number       | int    | No       | Hymnal index when available.                   |
| Title             | string | Yes      | Unique canonical title.                        |
| In_Hymnal         | bool   | No       | If omitted, assumed `True`.                    |

> Column names are matched case-insensitively; the ETL accepts common variants (e.g., `service_date`, `song`, `Song #`).

---

## Curated Outputs (`/data`)

### `kpi_top10.csv`
Top 10 most played songs in the 52-week window.
| Column | Type   | Notes            |
|---|---|---|
| Title  | string | Song title       |
| Plays  | int    | Count in window  |

### `kpi_by_source.csv`
Total plays by source/category in the 52-week window.
| Column | Type   | Notes                  |
|---|---|---|
| Source | string | Category or “Unknown”  |
| Plays  | int    | Total plays            |

### `kpi_hymnal_coverage.csv`
Coverage of hymnal titles used at least once in the window.
| Column            | Type | Notes                           |
|---|---|---|
| Hymnal_Songs      | int  | Count of `In_Hymnal=True`       |
| Used_Hymnal_Songs | int  | Subset used at least once       |
| Coverage_%        | float| `Used/Hymnal * 100` (2 decimals)|

### `hymnal_unused.csv`
Hymnal songs not used in the window.
| Column      | Type   | Notes                      |
|---|---|---|
| Song_Number | int    | May be null if unknown     |
| Title       | string | Canonical title            |

