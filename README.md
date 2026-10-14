# CA Traffic Degrade Analysis Tool

A VBA-based Excel automation tool for processing Carrier Aggregation (CA) traffic data, cell reports, and generating KPI summaries at both cell and site levels.

---

## 📋 Overview

This tool automates the process of:
- Importing CA query data and Cell Report data
- Aggregating KPIs at the cell level
- Generating site-level summaries
- Highlighting zero KPI values for quick issue identification
- Exporting results to a separate workbook

---

## ✨ Features

- **Multi-file Import**: Select both CA and Cell Report files in one go
- **Cell-Level Aggregation**: Groups RAW DATA by cell name with min/max date tracking
- **KPI Period Tracking**: Shows week numbers from oldest to latest RAW DATA date per cell
- **Site Summary**: Aggregates cell data by site and band (L1800, L2100, L900)
- **Zero KPI Highlighting**: Automatically highlights zero values in red (with CA category logic)
- **Formatted Output**: Professional formatting with headers, borders, and auto-fit columns
- **Export Function**: Saves results to a dated Excel file

---

## 📁 Required Workbook Structure

Your master workbook **must** contain these sheets (case-insensitive):

| Sheet Name   | Purpose                          |
|--------------|----------------------------------|
| `Control`    | Control parameters (required)    |
| `Tracker`    | Site-to-CA Category mapping      |

The tool will create these sheets automatically:

| Sheet Name                | Purpose                              |
|---------------------------|--------------------------------------|
| `RAW DATA`                | Imported CA query data               |
| `Cell Report`             | Imported Cell Report data            |
| `Cell Level Detail KPI`   | Aggregated cell-level KPIs           |
| `Site-Summary`            | Site-level summary by band           |

---

## 📥 Input File Requirements

### File Naming Convention

| File Type      | Filename Must Contain | Example                          |
|----------------|-----------------------|----------------------------------|
| CA Data        | `ca`                  | `CA_Query_2025.xlsx`             |
| Cell Report    | `cell`                | `Cell_Report_Feb2025.xlsx`       |

### CA Data File Columns (Expected)

| Column | Header              |
|--------|---------------------|
| 1      | Cell Name           |
| 4      | Site ID             |
| 8-16   | KPI Metrics         |
| 17     | Date                |
| 20     | Additional KPI      |

### Cell Report File Columns (Expected)

| Column | Header              |
|--------|---------------------|
| 9      | Cell Name           |
| 18     | Availability Status |

### Tracker Sheet Columns (Required)

| Column | Header        |
|--------|---------------|
| A      | Site ID       |
| D      | CA Category   |

---

## 🚀 Usage Instructions

### Step 1: Prepare Master Workbook

1. Save this workbook with the VBA code
2. Ensure `Control` and `Tracker` sheets exist
3. Populate `Tracker` with Site ID → CA Category mappings

### Step 2: Run the Import Process

1. Open the workbook
2. Press `ALT + F11` to open VBA Editor
3. Press `F5` or run `ImprotReportData` subroutine
4. **OR** assign a button to `ImprotReportData` for easier access

### Step 3: Select Files

- A file dialog will appear
- Select **both** CA data file and Cell Report file (multi-select enabled)
- Click **Open**

### Step 4: Review Output

| Sheet                   | Description                                      |
|-------------------------|--------------------------------------------------|
| `Cell Level Detail KPI` | Per-cell KPIs with week range from RAW DATA      |
| `Site-Summary`          | Site-level summary by band with issue tracking   |

### Step 5: Export Results

- Results are automatically exported to:  
  `Export CA Cell Vs Site Level KPIs Result - DD-MM-YYYY.xlsx`
- Saved in the same folder as the master workbook

---

## 📊 Output Sheet Details

### Cell Level Detail KPI

| Column | Header              | Description                              |
|--------|---------------------|------------------------------------------|
| A      | Site                | Site ID (extracted from Cell Name)       |
| B      | Band                | L900 / L1800 / L2100 / Indoor            |
| C      | Availability Status | From Cell Report file                    |
| D      | CA Category         | From Tracker sheet lookup                |
| E-Q    | KPI Metrics         | Aggregated values (Sum or Average)       |
| Q      | KPI Period          | **Week X ~ Week Y (dd-mmm-yyyy to dd-mmm-yyyy)** |
| R      | CA KPI Status       | List of KPIs with zero values            |

**KPI Period Example:**
- Same year: `Week 8 ~ Week 12 (17-Feb-2025 to 21-Mar-2025)`
- Across years: `Week 52 2024 ~ Week 2 2025 (23-Dec-2024 to 10-Jan-2025)`

### Site-Summary

| Column | Header            | Description                        |
|--------|-------------------|------------------------------------|
| A      | Site ID           | Site identifier                    |
| B      | CA Solution       | CA Category from Tracker           |
| C-E    | L1800/L2100/L900  | Achieve / Not Achieve per band     |
| F      | Remark            | Overall site status                |
| G      | Issue             | All KPI issues for the site        |
| H-J    | Band KPI Issues   | Issues broken down by band         |

---

## 🔧 Available Macros

| Macro Name              | Description                                  |
|-------------------------|----------------------------------------------|
| `ImprotReportData`      | Main process: import, aggregate, export      |
| `AggregateByCell`       | Rebuild cell-level detail only               |
| `CreateSiteSummary`     | Rebuild site summary only                    |
| `HighlightZeros`        | Re-apply zero KPI highlighting               |
| `FormatAllSheets`       | Re-apply formatting to all sheets            |
| `ExportCaCellVsSiteKpis`| Export results to new workbook               |
| `DeleteAllSheet`        | Clear all generated sheets (keep Control/Tracker) |

---

## ⚙️ KPI Period Logic (Week Number Tracking)

For each cell, the tool scans all RAW DATA rows and tracks:

1. **Oldest Date**: Minimum date found for that cell
2. **Latest Date**: Maximum date found for that cell
3. **Week Numbers**: ISO week numbers (Monday start, 4-day rule)
4. **ISO Week Year**: Handles year boundaries correctly

**Formula:**
```
KPI Period = Week [Oldest] ~ Week [Latest] (dd-mmm-yyyy to dd-mmm-yyyy)
```

**Functions Used:**
- `WeekNumber(d)` - ISO week number
- `IsoWeekYear(d)` - ISO week year (Thursday rule)
- `FormatKpiPeriod(oldDate, latestDate)` - Formatted output string

---

## 🎨 Formatting Rules

| Element           | Style                                      |
|-------------------|--------------------------------------------|
| Headers           | Bold, White text, Dark Blue background     |
| Data Cells        | Calibri 11pt, Centered, Borders            |
| Zero KPI (F:P)    | **Red Bold** (unless CA category ignores)  |
| Non-Zero KPI      | Black, Normal                              |
| Column Width      | AutoFit                                    |
| Gridlines         | Hidden                                     |

---

## ⚠️ Zero KPI Highlighting Logic

Zero values in columns F-P are highlighted **RED BOLD** unless:

| CA Category | Ignored KPI Pattern                        |
|-------------|--------------------------------------------|
| 2CC         | 3CC AvgUser, 3CC Throughput                |
| 3CC         | 2CC AvgUser, 2CC Throughput                |

This prevents false positives when a KPI is not applicable to the CA configuration.

---

## 🐛 Troubleshooting

| Issue                              | Solution                                      |
|------------------------------------|-----------------------------------------------|
| "Select a CA data file and Cell Report file" | Ensure filenames contain `ca` and `cell`     |
| "Workbook must contain Control and Tracker"  | Create these sheets before running            |
| "No data rows in RAW DATA"         | Check source file has data from row 2 onwards |
| "Cell Name not found"              | Verify column headers in source files         |
| "Date not found"                   | Ensure Date column exists in CA file          |
| Export fails                       | Save master workbook before exporting         |
| Week numbers incorrect             | Dates must be valid Excel date format         |

---

## 📝 Notes

- **ISO Week Standard**: Uses Monday as first day, first 4-day rule for Week 1
- **Date Parsing**: Handles string, numeric, and date variants
- **Case-Insensitive**: Sheet names and lookups are case-insensitive
- **Error Handling**: Most errors show descriptive messages
- **Screen Updating**: Disabled during processing for performance

---

## 📄 Version History

| Version | Date       | Changes                              |
|---------|------------|--------------------------------------|
| 1.0     | 2025-01-01 | Initial release                      |
| 1.1     | 2025-03-01 | Added KPI Period with week numbers   |
| 1.2     | 2025-03-01 | ISO week year support for cross-year |

---

## 📞 Support

For issues or enhancements:
1. Check error messages in message boxes
2. Verify file naming conventions
3. Ensure Tracker sheet has correct Site-to-CA mappings
4. Review RAW DATA column headers match expected format

---

## ⚖️ License

Internal use only. Not for distribution outside organization.
