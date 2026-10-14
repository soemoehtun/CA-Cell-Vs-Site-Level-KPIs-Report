# CA Cell vs Site KPI Report

Excel VBA tool that imports 7-day CA traffic data, builds cell-level
KPIs and a site summary, then exports both sheets next to the workbook.

## Requirements

-   Microsoft Excel with macros enabled
-   Workbook must already contain these sheets:
    -   `Control`
    -   `Tracker`
-   `Tracker` column A = Site ID, column D = CA Category (`2CC` / `3CC`)
-   Save the workbook at least once (`ThisWorkbook.Path` is used for
    export)

## Input files

Run `DegradeTrafficProcess` and select **two** Excel files:

  File              Filename must contain   Used for
  ----------------- ----------------------- ---------------------
  CA traffic data   `ca`                    RAW DATA KPIs
  Cell report       `cell`                  Availability Status

If either file is missing, the process stops.

## Process overview

``` text
Select CA + Cell Report files
        ↓
Delete all sheets except Control and Tracker
        ↓
Import RAW DATA and Cell Report
        ↓
Build Cell Level Detail KPI 9 sites
        ↓
Build Site-Summary
        ↓
Format sheets
        ↓
Export xlsx to ThisWorkbook.Path
```

## How to run

1.  Open the macro workbook.
2.  Enable macros.
3.  Run `DegradeTrafficProcess` (or `ImportDataforTrafficDegrade`).
4.  Select the CA file and the Cell Report file.
5.  Wait for the completion message.

Optional macros:

  Macro                      Purpose
  -------------------------- -------------------------------------------
  `DegradeTrafficProcess`    Full import → detail → summary → format
  `AggregateByCell`          Rebuild cell detail only
  `CreateSiteSummary`        Rebuild site summary from existing detail
  `FormatAllSheets`          Re-apply formatting
  `ExportCaCellVsSiteKpis`   Export the two result sheets

## Sheets

  Sheet                             Role
  --------------------------------- -------------------------------------------------------
  `Control`                         Kept. Not deleted.
  `Tracker`                         Kept. CA Category lookup (`VLOOKUP` Site → column D).
  `RAW DATA`                        Imported CA columns.
  `Cell Report`                     Imported cell name + availability.
  `Cell Level Detail KPI 9 sites`   One row per cell.
  `Site-Summary`                    One row per site.

## Cell Level Detail KPI 9 sites

Built from `RAW DATA` + `Cell Report` + `Tracker`.

  Column     Content
  ---------- ----------------------------------------
  A          Site (from cell name, chars 3--10)
  B          Band
  C          Availability Status (from Cell Report)
  D          CA Category (from Tracker)
  E onward   KPI values
  Q          KPI Period (`Week n ~ Week n`)
  R          CA KPI Status

Band from cell-name prefix:

  Prefix   Band
  -------- --------
  `LE`     L1800
  `LT`     L2100
  `LN`     L900
  Other    Indoor

KPI aggregation:

-   Most KPIs = **average**
-   These are **sum**:
    -   `Total CA Traffic(GB)`
    -   `PCell Traffic(Mbit)`
    -   `SCell Traffic(Mbit)`

### CA KPI Status rules

A KPI is an issue only if its value is **0** and it is **not** an
expected 2CC/3CC zero.

  CA Category   KPI = 0                             Result
  ------------- ----------------------------------- ---------------------------------
  `2CC`         `CA 3CC Avg User`                   Ignored
  `2CC`         `CA 3CC DL User Throughput(Mbps)`   Ignored
  `3CC`         `CA 2CC Avg User`                   Ignored
  `3CC`         `CA 2CC DL User Throughput(Mbps)`   Ignored
  Any           Any other KPI = 0                   KPI name listed in status
  Any           No real issues                      **Blank** (no `Achieve` remark)

Examples:

-   2CC, only 3CC Avg User / Throughput are 0 → `CA KPI Status` = blank
-   3CC, only 2CC Avg User / Throughput are 0 → `CA KPI Status` = blank
-   2CC, Total CA Traffic also 0 → `Total CA Traffic(GB)`

### Formatting on this sheet

-   Header row: bold, themed fill
-   Font: Calibri 11, center, thin borders
-   Gridlines off
-   Columns **F to P**:
    -   Number format = **General** (shows `0`, not `0.00`)
    -   Value = 0 → **red bold**
    -   Value ≠ 0 → normal black

## Site-Summary

Rolled up from `Cell Level Detail KPI 9 sites`.

  Column   Content
  -------- -----------------------------------
  A        Site ID
  B        CA Solution (CA Category)
  C        L1800 (`Achieve` / `Not Achieve`)
  D        L2100
  E        L900
  F        Remark (site overall)
  G        Issue (all KPI names)
  H        L1800 KPI Issues
  I        L2100 KPI Issues
  J        L900 KPI Issues

Rules:

-   Blank `CA KPI Status` = Achieve for that band
-   Any remaining KPI name in status = Not Achieve for that band
-   Remark = `Not Achieve` if any present band failed, else `Achieve`
-   Indoor band is not included in L1800 / L2100 / L900

## Export

Macro: `ExportCaCellVsSiteKpis`

Saved to the same folder as this workbook:

``` text
Export CA Cell Vs Site Level KPIs - Export Date(DD-MM-YYYY).xlsx
```

Example:

``` text
D:\CA Report\Export CA Cell Vs Site Level KPIs - Export Date(25-03-2026).xlsx
```

Contents:

1.  `Cell Level Detail KPI 9 sites`
2.  `Site-Summary`

Formatting is kept, including red-bold zeros. An existing file with the
same name is overwritten.

If the workbook has never been saved, export stops and asks you to save
first.

To export automatically after processing, call `ExportCaCellVsSiteKpis`
after `FormatAllSheets`.

## Setup

1.  Put the process code in one standard module.
2.  Put the export code in a second standard module.
3.  Keep `Control` and `Tracker` in the workbook.
4.  Fill `Tracker` with Site ID (A) and CA Category (D).
5.  Save the workbook.
6.  Run `DegradeTrafficProcess`.

## Notes

-   Source files are opened read-only when not already open, then closed
    if this tool opened them.
-   Only sheets named `Control` and `Tracker` survive a full run; all
    other sheets are deleted first.
-   Expected 2CC/3CC zeros are still shown as `0` in F:P (red bold).
    They are only excluded from `CA KPI Status`.
