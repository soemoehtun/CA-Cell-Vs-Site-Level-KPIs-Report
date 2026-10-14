# CA Traffic Degrade Analysis Tool

An Excel VBA automation tool that processes CA Query Data and Cell Report data, generates cell-level and site-level KPI reports, and highlights zero-KPI values for issue investigation.

## 1. Input Files

- **CA Query Data** — Contains CA traffic and KPI data.
- **Cell Report Data** — Contains cell-level information required for KPI processing.

## 2. Required Workbook Structure

The master workbook must contain the following worksheets. Worksheet names are matched without regard to case.

| Worksheet | Purpose |
|---|---|
| `Control` | Required control worksheet. |
| `Tracker` | Maps Site ID to CA Category. |

The tool creates or populates the following worksheets during processing:

| Worksheet | Purpose |
|---|---|
| `RAW DATA` | Imported CA query data. |
| `Cell Report` | Imported Cell Report data. |
| `Cell Level Detail KPI` | Aggregated cell-level KPIs and KPI period. |
| `Site-Summary` | Site-level summary by frequency band. |

**Important:** Keep the `Control` and `Tracker` worksheets in the master workbook. The `DeleteAllSheet` macro clears generated worksheets while preserving these two required worksheets.

## 3. Input File Selection

1. Run the processing macro to open the file selection dialog.
2. Select both the CA data file and the Cell Report file.
3. Ensure the filenames contain the expected `ca` and `cell` keywords.
4. Click **Open** to start processing.

## 4. Processing Logic

The tool performs the following operations:

1. **Import source data** — Load the selected CA Query Data and Cell Report data into `RAW DATA` and `Cell Report`.
2. **Process cell-level KPIs** — Aggregate the required KPI data and determine the applicable KPI period.
3. **Match cell information** — Use the required cell information and the `Tracker` worksheet to associate Site IDs with their CA Categories.
4. **Generate cell-level results** — Populate `Cell Level Detail KPI` with the processed KPI values, availability status, CA category, KPI period, and zero-KPI issue details.
5. **Generate site-level summaries** — Aggregate the processed cell-level results by site and frequency band in `Site-Summary`.
6. **Identify zero-KPI values** — Highlight applicable zero KPI values to help identify potential issues while excluding KPIs that do not apply to the cell's CA Category.

## 5. Zero KPI Highlighting Logic

Zero values in the applicable KPI range, configured as columns F–P in the original setup, are highlighted in **red and bold** unless the KPI is not applicable to the cell's CA Category.

| CA Category | KPI Patterns Ignored |
|---|---|
| `2CC` | `3CC AvgUser`, `3CC Throughput` |
| `3CC` | `2CC AvgUser`, `2CC Throughput` |

This exception logic prevents false-positive issue flags for KPIs that are not relevant to the configured CA solution.

## 6. Review the Output

Review the following generated worksheets after processing:

- **`Cell Level Detail KPI`** — Contains cell-level KPI results, availability status, CA category, KPI period, and zero-KPI issue details.
- **`Site-Summary`** — Contains site-level results by frequency band, overall status, and issue details.

## 7. Export the Report

The tool exports the processed results to a dated Excel workbook in the same folder as the saved master workbook.

**Expected filename pattern:**

```text
Export CA Cell Vs Site Level KPIs Result - DD-MM-YYYY.xlsx
```

Save the master workbook before exporting to ensure Excel has a valid destination folder.

## 8. Output Summary

The tool generates the following results:

- Imported CA Query Data and Cell Report data.
- Cell-level KPI details and KPI period.
- Site-level KPI summaries by frequency band.
- CA Category and availability status.
- Zero-KPI highlighting with CA Category-specific exceptions.
- A dated Excel report containing the processed results.

