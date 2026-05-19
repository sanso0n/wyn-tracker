# WYN AD Tracker · Portable folder

Self-contained website that shows the current state of the [WYN] World You Need alliance in the Last Z: Survival Alliance Duel.

## How to open the website

**Double-click `index.html`.** Any modern browser works (Chrome, Safari, Firefox, Edge).

The page is fully self-contained — all data is embedded in the HTML. **No internet connection required.**

Two views are linked from each other:

- `index.html` — public-facing version (cover, podium, day-by-day, Hall of Fame, history, join CTA).
- `dashboard.html` — detailed dashboard (full ranking, time zones with world map, coffers, history).

## Files in this folder

| File | What it is |
|---|---|
| `index.html` | Main page (the "show") |
| `dashboard.html` | Detailed dashboard |
| `current-week-extended.json` | Embedded data: 100 warriors of the current week |
| `history.json` | Embedded data: Diamond April + Gold May history |
| `WYN Logo.jpeg` | Alliance logo (shown in the cover) |
| `Mapa Mundi WYN.jpeg` | World map (shown in the Time zones tab of the dashboard) |
| `template.html` | HTML template used by the refresh script |
| `refresh_from_excel.py` | Optional · regenerates the data from Excel files |
| `README.md` | This file |

## Updating the data (optional)

If you want to refresh the website with new weekly data, you need:

1. Python 3.x installed.
2. The `openpyxl` package: `pip3 install --user openpyxl`
3. The Excel files placed in a parent folder, with this structure:

```
SomeFolder/
  ├─ Clasificacion_WYN.xlsx                      ← master file (current cycle)
  ├─ Gold Group Mayo/
  │   └─ 01-WYN vs OM3N/
  │       └─ Datos_parciales_WYN_vs_OM3N.xlsx    ← weekly daily data
  ├─ Diamond Group Abril/
  │   └─ Clasificacion_WYN Abril.xlsx            ← closed cycle (historical)
  └─ WYN-Tracker-Portable/                       ← this folder
      └─ refresh_from_excel.py
```

Then from a Terminal in this folder:

```bash
python3 refresh_from_excel.py            # full refresh
python3 refresh_from_excel.py --preview  # only print a summary, no files changed
```

The script:
- Reads the Excel files.
- Re-builds `current-week-extended.json` and `history.json`.
- Re-embeds the data into `index.html` and `dashboard.html`.

If the Excel files are not present in the expected place, the script will print a warning and stop without modifying anything.

## How it was designed

- **Source of truth: the Excel files.** The website only displays what's in them.
- **Apocalypse Time = UTC.** The Alliance Duel week closes at `00:00 Apocalypse`. Players are grouped into three operational time bands: AMER (UTC −3/−10), EMEA (UTC −1/+4) and APAC (UTC +5/+12). A 4th group, UNK, holds players whose nationality is unknown.
- **Language: English** as the default. The structure is ready for a future ES/EN toggle.

## Notes

- The page works fully offline. The only external dependency is Google Fonts (Oswald + Inter), which the browser will try to download once and then cache. If you have no internet, the page still works — it just falls back to system fonts.
- All player data is embedded in the HTML at the time of refresh. You do **not** need a web server to view the page.

---

*Generated: May 2026 · Last Z Survival · Server 221*
