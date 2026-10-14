# CA Traffic Degrade Analysis Tool

A VBA-based Excel automation tool for processing Carrier Aggregation (CA) traffic data, aggregating cell-level KPIs, generating site-level summaries, and exporting formatted reports.

## Overview

The tool automates the following tasks:

- Imports CA query data and Cell Report data.
- Aggregates KPI metrics by cell, including the oldest and latest dates.
- Tracks the ISO week range covered by each cell's data.
- Creates site-level summaries by CA category and frequency band.
- Highlights zero KPI values to help identify potential issues.
- Exports the generated reports to a separate Excel workbook.

## Features

- **Multi-file import:** Select the CA data file and Cell Report file in one file-selection process.
- **Cell-level aggregation:** Groups `RAW DATA` records by cell name.
- **KPI period tracking:** Displays the ISO week range from the oldest to the latest date for each cell.
- **Site-level summary:** Summarizes results by site and band (`L1800`, `L2100`, and `L900`).
- **Zero KPI highlighting:** Highlights zero values in red and bold, with exceptions for KPIs that do not apply to a CA category.
- **Formatted output:** Applies header styling, borders, centered values, AutoFit column widths, and hidden gridlines.
- **Automatic export:** Saves the generated reports to a dated Excel workbook.

## Requirements

- Microsoft Excel for Windows with macros enabled.
- A master workbook saved as an Excel macro-enabled workbook (`.xlsm`).
- The required `Control` and `Tracker` worksheets.
- CA data and Cell Report source files in a supported Excel format.

## Required Workbook Structure

The master workbook must contain these worksheets. Sheet names are matched without regard to case.

| Worksheet | Purpose |
|---|---|
| `Control` | Required control worksheet. |
| `Tracker` | Maps Site ID to CA Category. |

The tool creates or populates these generated worksheets:

| Worksheet | Purpose |
|---|---|
| `RAW DATA` | Imported CA query data. |
| `Cell Report` | Imported Cell Report data. |
| `Cell Level Detail KPI` | Aggregated cell-level KPIs and KPI period. |
| `Site-Summary` | Site-level summary by frequency band. |

> Keep the `Control` and `Tracker` worksheets in the master workbook. The `DeleteAllSheet` macro is intended to clear generated worksheets while preserving these required sheets.

## Input File Requirements

### File naming convention

The import process identifies files by keywords in their names.

| File type | Filename must contain | Example |
|---|---|---|
| CA data | `ca` | `CA_Query_2025.xlsx` |
| Cell Report | `cell` | `Cell_Report_Feb2025.xlsx` |

Use filenames that clearly identify each file type.

### CA data file

The source data is expected to contain the following fields in these columns:

| Column | Expected content |
|---|---|
| A (1) | Cell Name |
| D (4) | Site ID |
| H:P (8–16) | KPI metrics |
| Q (17) | Date |
| T (20) | Additional KPI |

### Cell Report file

| Column | Expected content |
|---|---|
| I (9) | Cell Name |
| R (18) | Availability Status |

### Tracker worksheet

| Column | Required content |
|---|---|
| A | Site ID |
| D | CA Category |

Make sure the Site ID values in `Tracker` match the identifiers used by the source data.

## How to Use

### 1. Prepare the master workbook

1. Open the Excel workbook containing the VBA code.
2. Save it as an Excel Macro-Enabled Workbook (`.xlsm`).
3. Confirm that the `Control` and `Tracker` worksheets exist.
4. Populate the `Tracker` worksheet with the Site ID-to-CA Category mappings.

### 2. Run the import process

1. Enable macros if Excel prompts you.
2. Press `Alt + F11` to open the VBA Editor.
3. Locate the `ImprotReportData` procedure.
4. Place the cursor inside the procedure and press `F5` to run it.

Alternatively, assign the `ImprotReportData` macro to a worksheet button for easier access.

### 3. Select the source files

1. When the file dialog opens, select both the CA data file and the Cell Report file.
2. Confirm that the filenames contain the expected `ca` and `cell` keywords.
3. Click **Open** to start processing.

### 4. Review the output

Review the generated worksheets:

- **`Cell Level Detail KPI`** — cell-level KPI results, availability status, CA category, KPI period, and zero-KPI issue details.
- **`Site-Summary`** — site-level results by band, overall status, and issue details.

### 5. Export the report

The tool exports the results to a dated Excel workbook in the same folder as the saved master workbook.

Expected filename pattern:

```text
Export CA Cell Vs Site Level KPIs Result - DD-MM-YYYY.xlsx
```

Save the master workbook before exporting so Excel has a destination folder.

## Output Sheet Details

### Cell Level Detail KPI

| Column(s) | Field | Description |
|---|---|---|
| A | Site | Site identifier extracted from the Cell Name. |
| B | Band | Frequency band, such as `L900`, `L1800`, `L2100`, or `Indoor`. |
| C | Availability Status | Availability information from the Cell Report file. |
| D | CA Category | CA category looked up from the `Tracker` worksheet. |
| E onward | KPI metrics | Aggregated KPI values. The aggregation method depends on the KPI. |
| KPI Period field | KPI Period | Week range and oldest/latest dates covered by the cell's raw records. |
| Final status field | CA KPI Status | List of KPI fields with zero values, subject to CA category exceptions. |

**Note:** Confirm the exact KPI column positions in the current VBA code if you change the source data layout. The KPI Period and CA KPI Status descriptions above refer to their logical fields; their physical column positions depend on the final output layout.

Example KPI period strings:

- Same year: `Week 8 ~ Week 12 (17-Feb-2025 to 21-Mar-2025)`
- Across years: `Week 52 2024 ~ Week 2 2025 (23-Dec-2024 to 10-Jan-2025)`

### Site-Summary

| Column(s) | Field | Description |
|---|---|---|
| A | Site ID | Site identifier. |
| B | CA Solution | CA Category from the `Tracker` worksheet. |
| C–E | Band status | Achievement status for `L1800`, `L2100`, and `L900`. |
| F | Remark | Overall site status. |
| G | Issue | Combined KPI issues for the site. |
| H–J | Band KPI Issues | KPI issues separated by band. |

## Available Macros

| Macro | Description |
|---|---|
| `ImprotReportData` | Main process: imports source files, aggregates KPIs, formats output, and exports results. |
| `AggregateByCell` | Rebuilds the cell-level detail worksheet. |
| `CreateSiteSummary` | Rebuilds the site-level summary. |
| `HighlightZeros` | Reapplies zero-KPI highlighting. |
| `FormatAllSheets` | Reapplies formatting to the worksheets handled by the macro. |
| `ExportCaCellVsSiteKpis` | Exports the CA cell- and site-level KPI results to a new workbook. |
| `DeleteAllSheet` | Clears generated worksheets while keeping `Control` and `Tracker`. |

> The macro name `ImprotReportData` is intentionally shown as provided. If you rename it in the VBA editor, update any assigned buttons or documentation references as well.

## KPI Period Logic

For each cell, the tool scans the matching records in `RAW DATA` and determines:

1. **Oldest date:** The minimum date recorded for that cell.
2. **Latest date:** The maximum date recorded for that cell.
3. **ISO week numbers:** Week numbers calculated using the ISO standard.
4. **ISO week-year:** Correctly handles dates near the end or beginning of a calendar year.

The resulting KPI Period is formatted as:

```text
Week [Oldest] ~ Week [Latest] (dd-mmm-yyyy to dd-mmm-yyyy)
```

The VBA helper functions are:

- `WeekNumber(d)` — returns the ISO week number.
- `IsoWeekYear(d)` — returns the ISO week-year.
- `FormatKpiPeriod(oldDate, latestDate)` — creates the formatted KPI period text.

The ISO week standard uses Monday as the first day of the week and the rule that Week 1 is the week containing the first Thursday of the year.

## Formatting Rules

| Element | Format |
|---|---|
| Headers | Bold white text on a dark blue background. |
| Data cells | Calibri 11 pt, centered, with borders. |
| Zero KPI values | Red and bold in the applicable KPI columns, except for ignored KPI patterns. |
| Non-zero KPI values | Black, normal text. |
| Column widths | AutoFit. |
| Gridlines | Hidden. |

## Zero KPI Highlighting Logic

Zero values in the applicable KPI range (described as columns F–P in the original configuration) are highlighted **red and bold**, unless the KPI is not applicable to the cell's CA category.

| CA Category | KPI patterns ignored |
|---|---|
| `2CC` | `3CC AvgUser`, `3CC Throughput` |
| `3CC` | `2CC AvgUser`, `2CC Throughput` |

This exception logic helps avoid false-positive issue flags for KPIs that are not relevant to the configured CA solution.

## Troubleshooting

| Problem | Suggested solution |
|---|---|
| “Select a CA data file and Cell Report file” | Select both source files and ensure their filenames contain `ca` and `cell`, respectively. |
| “Workbook must contain Control and Tracker” | Add the required worksheets to the master workbook. |
| “No data rows in RAW DATA” | Confirm the CA source file contains data rows starting from row 2. |
| “Cell Name not found” | Verify that the Cell Name field is in the expected source column. |
| “Date not found” | Confirm that the CA data file has valid dates in column Q. |
| Export fails | Save the master workbook before running the export macro, and check folder permissions. |
| Week numbers are incorrect | Ensure the source dates are valid Excel dates or are formatted in a date format the VBA code can parse. |
| CA category is missing | Check that the site's ID matches a value in column A of `Tracker` and that the category is in column D. |

## License

**Internal use only.** Do not distribute outside the organization without authorization.
