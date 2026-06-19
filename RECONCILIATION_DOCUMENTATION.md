# Payment Gateway API Reconciliation Documentation

## 1. Purpose

This project is a browser-based reconciliation app for matching payment gateway aggregator data with invoice data received from API exports.

The app lets the accounts team:

1. Upload aggregator Excel files, such as HDFC and BillDesk.
2. Upload API invoice Excel files, such as CME, CMS, LMA, and YLC.
3. Review or override category mapping.
4. Run reconciliation.
5. Review matched and exception records.
6. Export a final Tally-ready reconciled workbook.

The first exported sheet is `Invoice`, matching the final Tally import format shown in `Sample.xlsx`.

## 2. Project Structure

| File | Purpose |
| --- | --- |
| `index.html` | Vite application entry page. |
| `src/main.jsx` | Main React app, Excel parsing, reconciliation algorithm, and export logic. |
| `src/styles.css` | App layout and visual styling. |
| `package.json` | Dependencies and run/build scripts. |
| `Sample.xlsx` | Reference final Tally-ready `Invoice` sheet format. |
| `outputs/` | Generated reports and QA screenshots. |

## 3. How To Run

Install dependencies:

```bash
npm install
```

Start the app:

```bash
npm run dev -- --host 127.0.0.1
```

Build production files:

```bash
npm run build
```

The app is a client-side browser app. Excel files are parsed locally in the browser using the `xlsx` package. No backend is required.

## 4. Input Files

### 4.1 Aggregator Sheets

Aggregator files are payment gateway reports. Current supported formats are:

| Aggregator | Detection Logic | Expected Key Columns |
| --- | --- | --- |
| HDFC | File name contains `hdfc`, or sheet rows contain `CCAvenue Ref#`. | `CCAvenue Ref#`, `Order No`, `Order Amount`, `Gross Amount`, `Order Datetime`, `Order Status`, `Order Stlmt Date` |
| BillDesk | File name contains `billdesk`, sheet name contains `AIMAOTHERS`, or rows contain `PGI Ref. No.`. | `Dept`, `Ref. 1`, `PGI Ref. No.`, `Date of Txn`, `|`, `Charges`, `GST`, `Net Amount`, `Settlement Date`, `STATUS` |

The BillDesk gross amount is read from the `|` column. Charges, GST, and net amount are retained for review but are not used as the primary match amount.

### 4.2 API Invoice Sheets

API files are invoice exports by department/category.

Supported categories:

| Category | Example Files |
| --- | --- |
| CMS | `CMS INVOICE.xlsx` |
| CME | `CME INVOICE.xlsx` |
| LMA | `LMA INVOICE.xlsx` |
| YLC | `YLC INVOICE.xlsx` |
| Others | Any unclassified invoice file |

Expected API columns include:

| Column | Purpose |
| --- | --- |
| `Reference_No` | Primary invoice/payment reference used for matching. |
| `Transaction_No` | Alternate transaction reference used for matching. |
| `Total Amt` | Gross invoice amount used for amount comparison. |
| `Transaction_Date` | API invoice transaction date. |
| `Name of the Candidate` | Customer/candidate name for review. |
| `Invoice No` | Invoice number, when available. |
| `Cost Centre` or `Project Code` | Tally/project classification. |

The original API row headers and values are preserved so that the final `Invoice` export can reproduce the Tally-ready source row format.

## 5. Header Detection

Different Excel sheets may start their actual header row on different rows. For example, BillDesk may have report title rows before the table starts.

The app detects the header row by scanning the first 15 rows of each sheet.

Header scoring uses:

1. Number of non-empty cells in the row.
2. Extra weight for business keywords:
   - `ref`
   - `amount`
   - `date`
   - `order`
   - `invoice`
   - `category`
   - `dept`
   - `total`

The row with the highest score is treated as the header row.

## 6. Category Detection

API category is inferred from file name and sample row text.

Detection order:

1. Text contains `CMS` or `MAT` → `CMS`
2. Text contains `CME` → `CME`
3. Text contains `LMA` → `LMA`
4. Text contains `YLC` → `YLC`
5. Otherwise → `Others`

Users can override the category with the dropdown beside each API file.

The `Default Rules` button resets API file categories to automatic detection.

## 7. Aggregator Parsing Rules

## 7.0 Column-To-Column Matching Summary

The reconciliation checks reference columns first, then checks amount.

Important: the same payment gateway reference can appear in different API columns depending on the department/category. Because of this, the app does not use one fixed column for every source. It uses the mappings below.

### HDFC MAT / CMS Matching

For HDFC `MAT` data, the aggregator rows are treated as `CMS`.

| Check Order | HDFC Aggregator Column | API CMS Column | Meaning |
| --- | --- | --- | --- |
| 1 | `CCAvenue Ref#` | `Reference_No` | Primary CMS reference match. |
| 2 | `Order No` | `Transaction_No` | Alternate CMS transaction match. |

After either reference check matches, amount is compared:

```text
HDFC Order Amount = API Total Amt
```

### HDFC Others / CME Matching

For HDFC sheets other than `MAT`, the aggregator rows are treated as `CME`.

| Check Order | HDFC Aggregator Column | API CME Column | Meaning |
| --- | --- | --- | --- |
| 1 | `Order No` | `Reference_No` | Primary CME reference match. |
| 2 | `CCAvenue Ref#` | `Transaction_No` | Alternate CME transaction match. |

This is intentionally different from CMS. In the sample data, CME stores the HDFC `Order No` in API `Reference_No`, and stores the HDFC `CCAvenue Ref#` in API `Transaction_No`.

After either reference check matches, amount is compared:

```text
HDFC Order Amount = API Total Amt
```

### BillDesk / API Matching

For BillDesk, the same logic is used for CME, CMS, LMA, and YLC, based on the BillDesk `Dept` category.

| Check Order | BillDesk Aggregator Column | API Column | Meaning |
| --- | --- | --- | --- |
| 1 | `Ref. 1` | `Reference_No` | Primary BillDesk reference match. |
| 2 | `PGI Ref. No.` | `Transaction_No` | Payment gateway transaction match. |
| 3 | `PGI Ref. No.` | `Reference_No` | Fallback cross-reference check. |
| 4 | `Ref. 1` | `Transaction_No` | Fallback cross-reference check. |

`PGI Ref. No.` is the BillDesk payment gateway reference/transaction number. In the API invoice sheets, that same value commonly appears in `Transaction_No`.

After any reference check matches, amount is compared:

```text
BillDesk gross amount column "|" = API Total Amt
```

### Final Status Decision

For every aggregator row:

1. Try the configured reference checks.
2. If no API row is found, mark `Aggregator only`.
3. If an API row is found, compare amount.
4. If amount difference is within `0.01`, mark `Matched`.
5. If amount difference is greater than `0.01`, mark `Amount mismatch`.
6. After all aggregator rows are processed, any unused API row becomes `API only`.

### 7.1 HDFC MAT

HDFC sheets whose sheet name contains `MAT` are treated as CMS aggregator rows.

Parsed fields:

| Internal Field | Source Column |
| --- | --- |
| Category | `CMS` |
| Primary key | `CCAvenue Ref#` |
| Secondary key | `Order No` |
| Date | `Order Datetime` |
| Gross amount | `Order Amount`, fallback `Gross Amount` |
| Settlement date | `Order Stlmt Date` |
| Gateway status | `Order Status` |

Match attempts:

1. `CCAvenue Ref#` = API `Reference_No`
2. `Order No` = API `Transaction_No`

This means two columns are checked for HDFC MAT/CMS:

```text
HDFC CCAvenue Ref# -> API Reference_No
HDFC Order No      -> API Transaction_No
```

### 7.2 HDFC Others

All other HDFC sheets are currently treated as CME aggregator rows.

Parsed fields:

| Internal Field | Source Column |
| --- | --- |
| Category | `CME` |
| Primary key | `Order No` |
| Secondary key | `CCAvenue Ref#` |
| Date | `Order Datetime` |
| Gross amount | `Order Amount`, fallback `Gross Amount` |
| Settlement date | `Order Stlmt Date` |
| Gateway status | `Order Status` |

Match attempts:

1. `Order No` = API `Reference_No`
2. `CCAvenue Ref#` = API `Transaction_No`

This means two columns are checked for HDFC Others/CME:

```text
HDFC Order No      -> API Reference_No
HDFC CCAvenue Ref# -> API Transaction_No
```

### 7.3 BillDesk

BillDesk rows use the `Dept` column as the category when available. If `Dept` is blank, the app falls back to category inference from file/row content.

Parsed fields:

| Internal Field | Source Column |
| --- | --- |
| Category | `Dept` |
| Primary key | `Ref. 1` |
| Secondary key | `PGI Ref. No.` |
| Date | `Date of Txn` |
| Gross amount | `|`, fallback `Amount`, fallback `Gross Amount` |
| Charges | `Charges` |
| GST | `GST` |
| Net amount | `Net Amount` |
| Settlement date | `Settlement Date` |
| Gateway status | `STATUS` |

Match attempts:

1. `Ref. 1` = API `Reference_No`
2. `PGI Ref. No.` = API `Transaction_No`
3. `PGI Ref. No.` = API `Reference_No`
4. `Ref. 1` = API `Transaction_No`

This means two primary columns are checked for BillDesk:

```text
BillDesk Ref. 1       -> API Reference_No
BillDesk PGI Ref. No. -> API Transaction_No
```

The app also performs two fallback cross-checks:

```text
BillDesk PGI Ref. No. -> API Reference_No
BillDesk Ref. 1       -> API Transaction_No
```

## 8. Reconciliation Algorithm

### 8.1 Normalization

Before matching:

1. Empty, null, or undefined values become blank strings.
2. Dates are converted to `yyyy-mm-dd hh:mm:ss` where possible.
3. Numeric references ending in `.0` are converted to plain references.
4. Numeric values are parsed after removing commas.

This prevents common Excel formatting issues, especially references being read as numeric values.

### 8.2 API Lookup Table

The app builds a lookup map from API rows.

Lookup keys are:

```text
category::referenceNo::<API Reference_No>
category::transactionNo::<API Transaction_No>
```

Example:

```text
CMS::referenceNo::114404880034
CME::transactionNo::114408331372
```

Each lookup key stores all matching API row indexes. This supports duplicate keys while still allowing each API row to be used only once.

### 8.3 Matching Pass

The app loops through each aggregator row.

For each aggregator row:

1. Read its category.
2. Run the configured match attempts for that aggregator type.
3. Find unused API rows with the same category and matching reference.
4. If multiple API candidates are found, choose the candidate with the smallest amount difference.
5. Mark the selected API row as used.

Amount difference:

```text
Aggregator gross amount - API Total Amt
```

Status rules:

| Condition | Status |
| --- | --- |
| Reference matched and amount difference is within `0.01` | `Matched` |
| Reference matched but amount difference is greater than `0.01` | `Amount mismatch` |
| Aggregator row has no matching API row | `Aggregator only` |
| API row was never used by any aggregator row | `API only` |

### 8.4 Pseudo-Code

```text
read aggregator workbooks
read API workbooks

for each API row:
  add API row to lookup by category + Reference_No
  add API row to lookup by category + Transaction_No

for each aggregator row:
  candidates = []

  for each configured match attempt:
    lookup category + attempted field + attempted value
    add unused API rows to candidates

  if candidates is empty:
    mark row as Aggregator only
    continue

  sort candidates by smallest absolute amount difference
  select first candidate
  mark selected API row as used

  difference = aggregator gross amount - API Total Amt

  if abs(difference) <= 0.01:
    mark as Matched
  else:
    mark as Amount mismatch

for each unused API row:
  mark row as API only

build summary
build Tally-ready Invoice sheet from clean Matched API rows
export workbook
```

## 9. Summary Metrics

The app builds category-level summary rows for:

| Metric | Meaning |
| --- | --- |
| Aggregator Records | Count of aggregator rows for the category. |
| API Records | Count of API rows for the category. |
| Matched Records | Count of reference-matched rows, including amount mismatches. |
| Amount Mismatches | Count of matched references where amounts differ. |
| Aggregator Only | Count of aggregator rows not found in API data. |
| API Only | Count of API rows not found in aggregator data. |
| Aggregator Gross Total | Sum of aggregator gross amount for the category. |
| API Total | Sum of API `Total Amt` for the category. |
| Matched Agg Gross | Sum of matched aggregator gross amounts. |
| Matched API Total | Sum of matched API amounts. |
| Unmatched Agg Gross | Sum of aggregator-only amounts. |
| Unmatched API Total | Sum of API-only amounts. |

## 10. Result Tabs In The App

| Tab | Purpose |
| --- | --- |
| `All` | Combined preview of all result types. |
| `Matched` | Clean matched records only. |
| `Amount mismatch` | Same reference matched, amount differs. |
| `Aggregator only` | Payment exists in aggregator, missing from API data. |
| `API only` | Invoice/API record exists, missing from aggregator. |

The table preview shows the first 150 rows for performance. Export includes all rows.

## 11. Exported Workbook

Clicking `Export to Excel` creates:

```text
reconciled_tally_report_YYYY-MM-DD.xlsx
```

### 11.1 Exported Sheets

| Sheet | Purpose |
| --- | --- |
| `Invoice` | Final Tally-ready invoice sheet from clean matched records. |
| `Summary` | Category-level reconciliation summary. |
| `Matching Rules` | Files and matching logic used for the run. |
| `Matched` | Matched record audit. |
| `Amount Mismatch` | Amount mismatch exceptions. |
| `Aggregator Only` | Aggregator-only exceptions. |
| `API Only` | API-only exceptions. |
| `All Results` | Combined audit data. |

### 11.2 Final `Invoice` Sheet

The `Invoice` sheet is designed to match the format in `Sample.xlsx`.

Rules:

1. Only clean `Matched` records are included.
2. `Amount mismatch`, `Aggregator only`, and `API only` rows are excluded from the final Tally import sheet.
3. Original API headers and values are preserved.
4. Rows are grouped by API header structure.
5. A blank row is inserted between groups.
6. The header row is repeated for each group, matching the sample format.
7. Group order is:
   - `CMS`
   - `CME`
   - `LMA`
   - `YLC`
   - `Others`

This is why the final export can contain different repeated headers for CMS, CME, and membership categories.

## 12. Button Behavior

| Button / Control | Behavior |
| --- | --- |
| Aggregator upload | Reads selected aggregator Excel files and lists their sheets. |
| API upload | Reads selected API Excel files and auto-assigns category. |
| API category dropdown | Manually overrides API file category. |
| `Default Rules` | Resets API categories using automatic file/category detection. |
| `Run Reconciliation` | Runs parsing, matching, exception generation, and summary creation. |
| Result tabs | Switch preview table between result sets. |
| `Export to Excel` | Downloads final Tally-ready reconciliation workbook. |
| `Reset` | Clears files, category overrides, results, and errors. |
| Remove file `x` | Removes that file and clears stale results. |

## 13. Known Assumptions

1. HDFC `MAT` sheet is treated as CMS.
2. Other HDFC sheets are treated as CME.
3. BillDesk category comes from `Dept`.
4. BillDesk gross payment amount is the `|` column.
5. API `Total Amt` is the amount used for matching.
6. One API row can be matched only once.
7. If duplicate API candidates exist, the one with the closest amount is selected.
8. Final Tally import should include only clean `Matched` rows.

## 14. QA Performed

The app was tested with the sample workbooks in this workspace:

| Test | Result |
| --- | --- |
| Upload HDFC and BillDesk aggregator workbooks | Passed |
| Upload CME, CMS, LMA, and YLC API workbooks | Passed |
| Category dropdown changes | Passed |
| `Default Rules` category reset | Passed |
| `Run Reconciliation` | Passed |
| Result tabs | Passed |
| `Export to Excel` | Passed |
| `Reset` | Passed |
| API wording scan | Passed |
| Final export contains `Invoice` sheet first | Passed |
| Final export follows `Sample.xlsx` repeated-header structure | Passed |

## 15. Future Enhancements

Useful future improvements:

1. Add configurable matching rules from the UI.
2. Add per-category tolerance settings.
3. Add direct validation warnings for missing required columns.
4. Add saved reconciliation profiles.
5. Add support for more aggregator formats.
6. Add a downloadable rejected-records-only workbook.
7. Add row-level comments explaining why a row is unmatched.
8. Add password/login if the app is deployed for multiple users.
