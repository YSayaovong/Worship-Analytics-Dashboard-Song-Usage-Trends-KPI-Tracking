# HFBC — Praise & Worship Dashboard

This project is a live, interactive dashboard designed for HFBC’s Worship & Praise ministry. It automatically reads schedules, announcements, setlists, analytics, and Bible study logs from spreadsheets and CSV files stored in this repository, and displays them in a clear, modern web interface.

## What It Does

- **Weekly Practices**: Automatically calculates and shows the next Thursday and Sunday worship practice sessions.
- **Additional Practices**: Displays extra scheduled practices (excluding Thursday and Sunday) directly from the practice log.
- **Announcements**: Shows announcements in both English and Hmong, with the newest displayed first and older than 31 days hidden automatically.
- **Worship Team Members**: Lists all team members and their roles.
- **Setlist**: Splits songs into upcoming week and last week, showing date, song title, and topic while preventing duplicate titles per service date.
- **Song Analytics**:
  - Top 10 most-played songs overall, displayed as a 3D pie chart.
  - A breakdown of the top 10 songs played in the current year, with counts of how many times each was sung.
- **Bible Study Log**: Tracks Bible study sessions including date, topic/passage, leader, and notes, displayed with most recent first.

## Key Points

- Updates automatically whenever the linked Excel or CSV files are updated in this repository.
- Handles data cleaning by ignoring invalid values such as “NA” or “church close.”
- Built entirely client-side with JavaScript, using SheetJS for spreadsheets, PapaParse for CSVs, and Google Charts for visualizations.
- Provides both English and Hmong content where applicable.

---

This dashboard is built to keep worship planning, music tracking, and study records in one central place, making it easy for the team to stay informed and focused. 
