# Sales Data Transformation with Excel Power Query: Transaction Log to Monthly Sales-Rep Summary

## Brief One Line Summary

Reshaped 4,930 transaction-level sales records into a 72-row by 12-month summary of sales per sales representative using Excel Power Query (month extraction, pivot, and aggregation).

---

## Overview

This project demonstrates how Power Query can turn a long, transaction-level sales log into a compact report that is ready for analysis. The source data records one row per sale, with a date, a sales representative, and a sales amount. The goal was to produce one row per representative with the total sales for each month of the year, so that individual and monthly performance can be compared at a glance.

The workbook contains three sheets:

| Sheet | Purpose |
|---|---|
| `Extra Practice Sheet #3` | The original raw data (titled "XYZ Corp. Sales Data F.Y. 2013"), 4,930 transactions, with a small `Final Outcome` block that shows the target layout |
| `Sheet1` | The data table used as the Power Query source (the same 4,930 transactions, with the title row removed) |
| `Table1` | The final output of the query: 72 sales representatives by 12 months |

---

## Problem Statement

Transaction-level data is the right format for recording sales, but it is difficult to use for reporting. In the raw sheet, a single sales representative can appear dozens or hundreds of times across the year, and there is no way to see monthly performance without aggregating the data first.

The specific problems to solve were:

1. **Wrong shape for analysis.** The data was in long format (one row per transaction). Management reporting needs wide format: one row per sales representative and one column per month.
2. **No month field.** Only a full date was available, so sales could not be grouped by month until the month was extracted from it.
3. **Multiple transactions per representative per month.** Sales for the same representative and month were spread across many rows and had to be summed. The same representative can even have several transactions on the same day (1,363 repeated date and representative combinations).
4. **Missing combinations.** Not every representative made a sale in every month. Pivoting these gaps produces empty (null) cells, which must be converted to zero so that the table is numerically complete and safe to use in calculations.
5. **Data types.** The date, representative name, and sales amount needed explicit and correct data types (date, text, whole number) before they could be transformed.
6. **Title row above the data.** On the raw sheet, a report title sat above the column headers, so the data did not begin in the first row.

---

## Dataset

**Source file:** `Data_cleaning_practice_3.xlsx`
**Source sheet:** `Extra Practice Sheet #3` (copied to `Sheet1` for the query)

### Raw data structure

| Column | Description | Type |
|---|---|---|
| `Date` | Date of the sale | Date |
| `SalesRep` | Name of the sales representative | Text |
| `Total Sales` | Sales amount for the transaction | Whole number |

Examples of raw rows:

| Date | SalesRep | Total Sales |
|---|---|---|
| 2013-01-01 | Isabel Cross | 1,980 |
| 2013-01-02 | Cecilia Manning | 876 |
| 2013-01-02 | Isabel Cross | 426 |
| 2013-01-02 | Shari Silva | 520 |

### Dataset profile

| Measure | Value |
|---|---|
| Total transactions | 4,930 |
| Date range | 1 January 2013 to 31 December 2013 |
| Distinct dates with sales | 362 |
| Sales representatives | 72 |
| Missing values | 0 |
| Total sales across the year | 5,030,335 |
| Smallest / largest single sale | 100 / 12,175 |
| Average / median sale | about 1,020 / 624 |

> The source does not state a currency, so amounts are reported as plain numbers.

---

## Tools and Technologies

- **Microsoft Excel** (Microsoft 365 / Excel 2016 or later)
- **Power Query** (Get & Transform Data) for all transformations
- **Power Query M language** (steps generated through the editor interface)

---

## Methods

The entire transformation is a single Power Query. Every step is recorded in the Applied Steps pane, so the process is transparent and can be refreshed whenever new transactions are added.

| Step | Power Query action | Purpose |
|---|---|---|
| 1 | Load the source table | Bring the 4,930 transactions into Power Query |
| 2 | **Change Type**: `Date` to Date, `SalesRep` to Text, `Total Sales` to Whole Number | Ensure every column has the correct data type before transforming |
| 3 | **Insert Month Name** (Date > Month > Name of Month) | Create a `Month Name` column such as "January" from each date |
| 4 | **Remove Columns** (`Date`) | The full date is no longer needed once the month is extracted |
| 5 | **Reorder Columns**: `Month Name`, `SalesRep`, `Total Sales` | Arrange the columns for the pivot |
| 6 | **Pivot Column** on `Month Name`, values from `Total Sales`, aggregated with **Sum** | Turn the month names into 12 columns and total each representative's sales per month |
| 7 | **Replace Values** (null with 0) across all 12 month columns | Fill the months in which a representative made no sales |

**Result:** 72 rows (one per representative) and 13 columns (`SalesRep` plus January to December), loaded to the `Table1` sheet.

**Note on column order:** Power Query creates the pivoted month columns in the order in which each month first appears in the data. Because the source is sorted by date, the columns come out in calendar order. If the source were unsorted, the columns would need to be reordered after the pivot.

### Accuracy check

The output reconciles exactly with the source:

- The sum of all values in the 12 month columns of `Table1` equals the sum of `Total Sales` in the source: **5,030,335**.
- All 72 representatives appear exactly once in the output.

You can repeat this check in Excel by comparing `SUM` of the month columns in `Table1` with `SUM` of `Total Sales` in `Sheet1`.

---

## Key Insights

**Monthly sales performance**

| Month | Transactions | Total Sales | Average per Transaction |
|---|---|---|---|
| January | 149 | 166,293 | 1,116 |
| February | 127 | 131,601 | 1,036 |
| March | 149 | 137,902 | 926 |
| April | 144 | 153,421 | 1,065 |
| May | 174 | 159,982 | 919 |
| June | 119 | 152,474 | 1,281 |
| July | 117 | 130,540 | 1,116 |
| August | 123 | 113,203 | 920 |
| September | 117 | 131,157 | 1,121 |
| October | 129 | 155,222 | 1,203 |
| November | 1,781 | 1,766,051 | 992 |
| December | 1,801 | 1,832,489 | 1,017 |

**What the summary shows**

- **November and December dominate the year.** Together they account for about 71.5% of annual sales (3,598,540 of 5,030,335). The cause is volume rather than larger sales: each of these months has roughly 1,800 transactions, compared with 117 to 174 in every other month, while the average sale stays at about 1,000 throughout. This pattern is worth confirming with the data owner (for example, a seasonal peak or a change in how sales were recorded) before drawing business conclusions.
- **August is the weakest month** (113,203) and **December the strongest** (1,832,489).
- **Top five representatives by annual sales:** Shari Silva (209,677), Drew Rogers (189,847), Robert Munoz (182,478), Max Rodgers (157,667), and Rachel Gomez (153,323).
- **Lowest annual total:** Jaime Pearson (25,074).
- **Participation changes through the year.** Between 50 and 63 representatives recorded sales in each month from January to October, while all 72 recorded sales in both November and December.
- **Zero cells.** 162 of the 864 representative-month cells (about 19%) are zeros, created by replacing nulls in step 7. These represent months in which that representative had no recorded sales.

**Data quality observations and next steps**

These are points identified on review that I would address in a further iteration:

- **Double spaces in names.** Every representative name contains two spaces between the first and last name (for example `Elvira  Bowman`), in both the source and the output. Adding a Trim and a space-collapsing step (or a Replace Values step from two spaces to one) would standardise the names and avoid problems in lookups and joins.
- **No totals.** The output has no annual total per representative and no grand total row. Adding an `Annual Total` column would make ranking representatives straightforward.
- **One exact duplicate row.** One transaction (Mable Lindsey, 28 December 2013, 483) appears twice. It may be a genuine repeat sale, but it should be checked against the original records.
- **Layout reference only.** The `Final Outcome` block on the raw sheet is a small layout example. The values in `Table1` are calculated from the full data and reconcile to the source total.

---

## Dashboard/Model/Output

### Before and after

**Before** (transaction log, `Sheet1`): one row per sale. These are all of Isabel Cross's January transactions.

| Date | SalesRep | Total Sales |
|---|---|---|
| 2013-01-01 | Isabel Cross | 1,980 |
| 2013-01-02 | Isabel Cross | 426 |
| 2013-01-08 | Isabel Cross | 870 |
| 2013-01-27 | Isabel Cross | 6,960 |

These four sales add up to 10,236, which is her January value in the output below.

**After** (`Table1`): one row per representative, with one column per month.

| SalesRep | January | February | March | April | May | June | July | August | September | October | November | December |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Isabel Cross | 10,236 | 3,404 | 4,973 | 683 | 1,868 | 9,421 | 10,702 | 1,022 | 1,232 | 5,628 | 39,795 | 57,111 |
| Elvira Bowman | 11,588 | 2,508 | 2,852 | 677 | 1,099 | 0 | 1,152 | 376 | 376 | 0 | 23,535 | 15,165 |
| Robert Munoz | 11,458 | 15,311 | 2,658 | 6,643 | 2,354 | 3,579 | 9,266 | 2,692 | 11,146 | 739 | 71,664 | 44,968 |
| Lee Mack | 11,249 | 0 | 1,744 | 840 | 0 | 0 | 7,692 | 0 | 0 | 0 | 14,589 | 6,322 |
| Mable Lindsey | 11,212 | 1,078 | 3,300 | 5,388 | 764 | 9,514 | 810 | 620 | 216 | 1,941 | 48,903 | 24,761 |
| Rachel Gomez | 10,849 | 5,414 | 5,838 | 3,631 | 3,849 | 2,497 | 2,387 | 8,923 | 644 | 4,432 | 54,408 | 50,451 |

The summary table can be used directly for a PivotTable, a heat map or conditional formatting, or loaded into Power BI or Tableau to build a sales performance dashboard.

---

## Results & Conclusion

- **Outcome:** 4,930 transaction records were summarised into a 72 by 12 table of monthly sales per representative, and the output reconciles exactly to the source total of 5,030,335.
- **Approach:** the work was completed entirely with Power Query's built-in transformations (change type, insert month name, remove and reorder columns, pivot with sum aggregation, and replace values), with no formulas or manual editing. The query can be refreshed whenever new transactions are added.
- **Main finding:** November and December together generate about 71.5% of annual sales, driven by transaction volume rather than sale size.
- **Skills demonstrated:** reshaping data from long to wide format, date-part extraction, aggregation, handling null values created by a pivot, and validating a result against its source.
- **Next steps:** standardise the double-spaced names, add annual totals, review the duplicate row, and connect the summary to a Power BI or Tableau dashboard.

---

## Connect with Me

**SAMANPREET KAUR**
BA Student | Aspiring Data Analyst, Business Analyst, and BI Developer

- LinkedIn: [https://www.linkedin.com/in/kaursamanpreet/](https://www.linkedin.com/in/kaursamanpreet/)
- Email: [kaursamanpreet970@gmail.com](mailto:kaursamanpreet970@gmail.com)
- GitHub: [https://github.com/kaursamanpreet970-pixel](https://github.com/kaursamanpreet970-pixel)
