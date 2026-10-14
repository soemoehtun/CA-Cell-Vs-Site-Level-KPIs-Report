# Cell Report & CA Data Aggregator (Excel VBA)

A VBA macro module that automates the import, aggregation, formatting, and export of cellular network KPI data. It combines raw cell‑level data with CA (Cell Access) query results and a Tracker reference sheet to produce **cell‑level detail KPIs** and a **site‑level summary**.

---

## 📌 Overview

| Step | What it does |
|------|--------------|
| **Import** | Prompts for two files: one with `ca` in the filename, one with `cell` in the filename. Extracts specific columns into staging sheets. |
| **Aggregate (Cell Level)** | Groups raw data by **Cell Name**, calculates oldest→latest date range, converts dates to **ISO week numbers**, sums traffic KPIs, averages others, and pulls CA Category & Availability Status via lookups. |
| **Highlight** | Conditional formatting: zeros in KPI columns are highlighted **red & bold** unless excluded by 2CC/3CC rules. |
| **Summarize (Site Level)** | Rolls up cell data to **Site ID**, tracks band status (L1800/L2100/L900), CA Solution, and lists KPI issues per band & site. |
| **Format & Export** | Applies consistent styling, borders, header colors, number/general formats, auto‑fit, then exports the Detail + Summary sheets to a dated workbook. |

---

## 🛠 Prerequisites

- **Excel** 2016+ (Windows/macOS). 64‑bit recommended for large datasets.
- **Macro security**: Enable macros for the workbook.
- **Required sheets in the host workbook**:
  - `Control`
  - `Tracker` *(used for CA Category lookup: key in col A, value in col D)*
- **Source file requirements**:
  - `RAW DATA` must contain columns whose normalized headers include `cellname` and `date`.
  - Headers are normalized by removing spaces/tabs and converting to lower case before matching.
- **⚠️ Warning**: `ImprotReportData` deletes all sheets except `Control` and `Tracker` before creating new ones. Back up any other sheets first.

---

## 🚀 How to Use

1. Open the workbook that contains this VBA module.
2. Ensure `Control` and `Tracker` sheets exist and are named exactly as above.
3. Run the main macro (via `Alt+F8`, button, or shortcut):
   - `ImprotReportData` – full import + aggregation + formatting + export.
   - `AggregateByCell` – only rebuilds the cell‑level detail from already‑imported raw data.
4. In the file picker, select:
   - The **CA data file** (filename must contain `ca`)
   - The **Cell Report file** (filename must contain `cell`)
5. The macro will:
   - Create `RAW DATA` and `Cell Report` staging sheets.
   - Build `Cell Level Detail KPI`.
   - Build `Site-Summary`.
   - Apply formatting & zero‑KPI highlighting.
   - Export `Cell Level Detail KPI` + `Site-Summary` to a new workbook named:  
     `Export CA Cell Vs Site Level KPIs Result - dd-mm-yyyy.xlsx`

---

## 📊 Sheet Structure

| Sheet | Purpose |
|-------|---------|
| `RAW DATA` | Staged columns imported from the CA file (configurable column list). |
| `Cell Report` | Staged columns imported from the cell report file. |
| `Cell Level Detail KPI` | One row per **Cell Name**. Contains Site, Band, Availability Status, CA Category, KPI values (summed/averaged), KPI Period (week range + actual dates), and CA KPI Status (zero‑KPI remarks). |
| `Site-Summary` | One row per **Site ID**. Shows CA Solution, band achievement status, overall remark, and KPI issue lists per band. |

---

## 🔑 Key Logic & Notes

### Date → Week Number
- For each cell, the code scans all RAW DATA dates, keeps the **oldest** and **latest**.
- Week numbers are calculated using **ISO week** (`vbMonday`, `vbFirstFourDays`).
- KPI Period text format:
  - Same year: `Week 8 ~ Week 12 (17-Feb-2025 to 21-Mar-2025)`
  - Cross year: `Week 52 2024 ~ Week 2 2025 (23-Dec-2024 to 10-Jan-2025)`

### KPI Aggregation
- Traffic columns (`totalcatraffic(gb)`, `pcelltraffic(mbit)`, `scelltraffic(mbit)`) are **summed**.
- All other numeric KPI columns are **averaged**.
- Non‑numeric or empty values are skipped.

### Zero‑KPI Highlighting
- Zeros are highlighted **red & bold** unless they match ignore rules:
  - 2CC cells ignore 3CC `avguser` / `throughput` zeros.
  - 3CC cells ignore 2CC `avguser` / `throughput` zeros.
  - “Achieve” status is never flagged.
- Highlighting runs after formatting and can be triggered separately via `HighlightZeroKPIs` or `HighlightZeros`.

### Lookups
- **Report lookup**: Cell Name → Availability Status from `Cell Report` sheet (col 1 → col 2).
- **Tracker lookup**: Site ID → CA Category from `Tracker` sheet (col A → col D).

---

## ⚙️ Customization

| What to change | Where |
|----------------|-------|
| Imported columns | `CopyColumnsFromFile` calls in `ImprotReportData` (arrays `Array(1, 4, 8, …)` and `Array(9, 18)`). |
| Sheet names | Constants at the top: `RAW_SHEET`, `REPORT_SHEET`, `DETAIL_SHEET`, `SUMMARY_SHEET`. |
| Zero‑KPI ignore rules | `ShouldIgnoreZeroKPI` function. |
| Week/period formatting | `FormatKpiPeriod` and `WeekNumber` / `IsoWeekYear` helpers. |
| Site summary columns / band logic | `BuildSiteSummary` and `SS2_BandIndex` / `SS2_StatusHasIssue`. |
| Styling | `FormatAllSheets` (fonts, colors, borders, number formats). |

---

## 🐛 Troubleshooting

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| `Select a CA data file and a Cell Report file…` | Selected files don’t contain `ca` and `cell` in filenames (case‑insensitive). | Rename files or select different ones. |
| `Could not find 'Tracker' sheet…` | Missing or misnamed sheet. | Add/rename a sheet to `Tracker` (CA Category in col D). |
| `No data rows in RAW DATA` | Source file has only headers or is empty. | Ensure at least one data row below the header. |
| `Too many KPI columns` | `METADATA_COLS + outCol` exceeds `KPI_PERIOD_COL`. | Reduce imported columns or adjust constants. |
| Week numbers look off for edge dates | ISO week rules (first 4‑day week, Monday start). | Expected behavior; adjust `WeekNumber` if your org uses a different week rule. |
| Export fails / path error | Workbook not saved yet (`ThisWorkbook.Path` is empty). | Save the host workbook first, then run export. |

---

## 📄 License & Author

- **Author**: Upstage AI (Solar Pro4) / adapted for internal use.
- **License**: Replace with your organization’s terms if distributing internally or externally.
- **Version**: 1.0

---

## 🔗 Quick Reference (Macro List)

| Macro | Purpose |
|-------|---------|
| `ImprotReportData` | Full pipeline: import → aggregate → summarize → format → export. |
| `AggregateByCell` | Rebuild cell‑level detail from existing raw data. |
| `CreateSiteSummary` | Rebuild site summary from existing detail sheet. |
| `HighlightZeroKPIs` / `HighlightZeros` | Re‑apply zero‑KPI red/bold formatting. |
| `ExportCaCellVsSiteKpis` | Export Detail + Summary to a dated workbook. |
| `FormatAllSheets` | Apply workbook‑wide styling (called automatically in the main flow). |

---

*Place this `README.md` in the same folder as the Excel workbook or in your version‑control repo for quick onboarding and maintenance.*
