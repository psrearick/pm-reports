# PM Reports – User Operations Guide

This document explains every feature that is available in the PM Reports workbook from the perspective of someone operating the Google Sheet. It focuses on what you can do in the interface, what the menu commands accomplish, and how the supporting sheets and Drive folders work together.

## 1. Platform Overview

The solution is centred on a Google Sheet that stores raw transactions, staging data for edits, configuration tables, and reporting templates. Custom menu actions (the **PM Reports** menu) orchestrate imports, cleans, reporting, and exports. Google Drive is used for versioned report storage and PDF exports. The workflow is:

1. Maintain the master **Transactions** table (directly or via **Import Credits** + **Clean Transactions Data**).
2. Use **Entry Controls** and **Entry & Edit** to review or adjust a date- and property-filtered slice of transactions.
3. Persist changes back to the master list with **Save Staging Data**.
4. Generate a reporting spreadsheet for a closing period, then export PDFs if required.
5. Review the automatically maintained log sheets for traceability.

## 2. Key Sheets and Their Roles

| Sheet | Purpose |
| --- | --- |
| **Transactions** | Master ledger. One row per transaction with IDs, amounts, flags, and timestamps. All reports pull from here. |
| **Entry & Edit** | Staging area populated by **Load Staging Data**. Users make edits/additions here before saving. Includes a `Delete Permanently` column for hard deletes. |
| **Entry Controls** | Control panel (cells B2–B7) that drives staging filters, report generation, report labels, and admin fee overrides. |
| **Properties** | Configuration for each property: MAF %, Markup %, Airbnb participation, admin fee settings, and keyword aliases to recognise property names during imports/cleans. |
| **Configuration** | Key/value settings that tell the automations where to find credit source files, target Drive folders, and column headers inside import files. |
| **ReportBodyTemplate** / **ReportTotalsTemplate** / **ReportAirbnbTemplate** | Hidden templates consumed by the templating engine when building the report spreadsheet. |
| **Import Log** | Audit trail for every credit file import (file ID, timestamp, rows imported, notes). Prevents duplicate processing. |
| **Report Log** | History of every generated report (label, period, spreadsheet ID/URL, version, admin-fee decisions). Used when exporting PDFs. |
| *(Optional)* **Reference** | You can paste the reference page content (see “Reference Page Content” document) here so end users have quick instructions inside the workbook. |

## 3. PM Reports Menu Actions

All automation is exposed under **PM Reports** in the sheet menu. Each command validates prerequisites, shows progress via toast notifications, and reports completion or errors.

### 3.1 Import Credits

- **Purpose:** Pull rent/fee data from configured Drive files into the master Transactions sheet.
- **Configuration required:** `Credits Document`, `Credits Folder ID`, header mappings (`Credits Date`, `Credits Amount`, `Credits Property`, `Credits Unit`, `Credits Category`, `Credits Subcategory`) and optional status/method/payer headers in **Configuration**.
- **Inputs:** Google Sheets or Excel files. Excel files are temporarily converted. Duplicate prevention uses the combination of File ID + Last Modified timestamp recorded in **Import Log**.
- **Process:**
  1. Locate files by direct ID and within the configured folder. Only process Google Sheets or Excel.
  2. For each file, locate the header row based on the configured header names.
  3. Skip rows that are blank, totals, lack a date, or have zero amount.
  4. Skip categories labelled “Property General Expense” or “Owner Distribution” and classify the remaining rows into Credits, Fees, or Security Deposits based on the category text (deposit → security deposit, tenant charges/fees/late fee → fees, otherwise credits).
  5. Resolve the property name using the **Properties** table keywords. Rows that cannot be matched are skipped.
  6. Create new transaction rows (with UUID IDs) and append them to **Transactions**.
  7. Optionally copy the parsed data into a new audit sheet if `Add Credits Sheet` is set to TRUE.
  8. Log the outcome in **Import Log**.
- **Results:** Newly appended transactions appear in **Transactions** (no staging step). Review with **Clean Transactions Data** if the import source uses inconsistent formatting.

### 3.2 Load Staging Data

- **Purpose:** Populate **Entry & Edit** with a filtered subset of master transactions for review or editing.
- **Controls:** Uses **Entry Controls** cells:
  - `B2` **Start Date** and `B3` **End Date** define the inclusive date range (leave blank for open-ended).
  - `B4` **Property** limits to one property; blank loads all.
  - `B6` **Show Deleted** toggles whether deleted rows appear.
- **Behaviour:**
  - Fetches transactions that match the filters and sorts them by date ascending.
  - Populates **Entry & Edit** with the standard headers plus `Delete Permanently`. Existing content below row 1 is cleared.
  - Boolean fields (`Deleted`, `Markup Included`) are presented as TRUE/FALSE checkboxes.

### 3.3 Save Staging Data

- **Purpose:** Push edits from **Entry & Edit** back into **Transactions**.
- **How it interprets rows:**
  - **New rows:** Rows without a Transaction ID are treated as new. Blank rows or rows flagged `Delete Permanently` are ignored. New rows receive a UUID, optional deleted timestamp, and computed markup if `Markup Included` is TRUE.
  - **Updates:** Rows whose IDs exist in master are updated. The deleted timestamp is preserved or cleared based on the `Deleted` flag. Markup revenue is recomputed.
  - **Delete Permanently:** When TRUE for an existing row, the corresponding master record is removed entirely.
  - **Implicit deletions:** Master records within the filtered set that are not present in staging (and not permanently deleted) are automatically toggled to Deleted=TRUE with a timestamp. This allows you to mark entries as deleted simply by removing them from staging.
- **After-save:** Master rows are updated in place, new rows appended, and deletions applied. Use **Load Staging Data** again to confirm results.

### 3.4 Clean Transactions Data

- **Purpose:** Normalise and repair the master **Transactions** sheet, especially after bulk copy/paste or external imports.
- **Actions performed:**
  - Generates missing or duplicate Transaction IDs.
  - Trims text fields and canonicalises property names using the **Properties** keywords.
  - Normalises unit, explanation, and internal notes text.
  - Standardises currency columns (Credits, Fees, Debits, Security Deposits, Markup Revenue) to plain numbers with two-decimal rounding and removes stray symbols or invalid values.
  - Converts boolean columns (`Markup Included`, `Deleted`) to TRUE/FALSE, and aligns Deleted Timestamps (adds current timestamp when missing, clears invalid values).
  - Validates dates and clears those that cannot be parsed.
- **Logging:** Any corrections or issues are written to the Apps Script execution log (view under *Extensions → Apps Script → Executions*). The toast message summarises how many rows were updated and how many issues were logged.
- **Recommendation:** Run this after manual data entry or imports to maintain clean downstream reporting.

### 3.5 Generate Report

- **Purpose:** Create a versioned reporting spreadsheet that includes a summary sheet, optional Airbnb sheet, and one tab per property.
- **Inputs:** The same **Entry Controls** values used for staging (Start Date, End Date, Property filter) plus:
  - `B5` **Report Label** – used in report naming and displayed in templates.
  - `B7` **Admin Fee Override** – leave blank to respect each property’s `Admin Fee Enabled` flag; set TRUE to force the admin fee on for all properties in scope; set FALSE to suppress it.
- **Processing steps:**
  1. Collect not-deleted transactions within the date range (and property filter if provided).
  2. Group by property (alphabetically sorted) and compute totals:
     - Credits, Fees, Debits, Security Deposits, Markup Revenue.
     - Unit count (unique, non-empty, excluding `CAM`).
     - MAF = `(Credits × property MAF rate) + (Unit count × 5)` plus optional Admin Fee.
     - Airbnb totals/fees when the property is flagged `Has Airbnb` and matching transactions include “airbnb” in Internal Notes.
     - New Lease/Renewal fees (based on exact Debit/Credit Explanation match).
     - Combined Credits, Total Debits, Due to Owners, Total to PM.
  3. Create a versioned spreadsheet in the configured **Reports** Drive folder (`Report Label`, or `Report Label_2`, etc.).
  4. Render property tabs using `ReportBodyTemplate`, the summary tab using `ReportTotalsTemplate`, and (when applicable) an Airbnb tab using `ReportAirbnbTemplate`.
  5. Delete the default `Sheet1`, then reorder sheets so **Summary** is first, **Airbnb** second (only if it was created), followed by property tabs in alphabetical order. Summary is left active.
  6. Write an entry to **Report Log** capturing the period, spreadsheet ID/URL, version number, properties included, and admin fee decisions.
- **Outputs:** One property tab per property, plus the summary (with per-property rows and aggregate placeholders) and optional Airbnb tab.

### 3.6 Export Report

- **Purpose:** Produce per-sheet PDFs for a previously generated report.
- **Inputs:** When invoked, you are prompted for a report label. Leave blank to export the most recent entry in **Report Log**, or enter a specific label to locate that report.
- **Process:**
  - Reads **Report Log** to find the spreadsheet ID.
  - Opens the spreadsheet, creates a versioned subfolder inside the configured **Exports** folder, and iterates through visible sheets except the hidden templates.
  - Each sheet is exported to PDF (landscape Letter, fit to width). A short pause between exports and a retry mechanism handle Google’s 429 rate limits automatically.
  - Files are named after the sheet (`<Sheet Name>.pdf`).
- **Outputs:** A timestamped subfolder containing PDFs for Summary, Airbnb (if present), and each property tab.

## 4. Supporting Configuration and Data Tables

### 4.1 Entry Controls Sheet (B2–B7)

| Cell | Name | Usage |
| --- | --- | --- |
| B2 | Start Date | Inclusive start for staging and reporting. Leave blank for “no lower bound”. |
| B3 | End Date | Inclusive end for staging and reporting. Leave blank for “no upper bound”. |
| B4 | Property | Optional single-property filter for staging and reporting. Must match the canonical property name. |
| B5 | Report Label | Text displayed on generated reports and used for naming/versioning. |
| B6 | Show Deleted | TRUE to include deleted rows in staging load. |
| B7 | Admin Fee Override | Blank = respect property default. TRUE = force admin fee for all properties. FALSE = suppress admin fee for all properties in the run. |

### 4.2 Configuration Sheet Keys

The Configuration sheet has two columns (`Setting`, `Value`). Required keys:

| Setting | Description |
| --- | --- |
| Credits Document | Optional direct file ID for a specific credits spreadsheet. |
| Credits Folder ID | Folder ID containing credit files to import. |
| Output Folder ID | Drive folder that will hold the Reports and Exports subfolders. |
| Reports Folder Name | Subfolder (created if absent) inside Output Folder for generated report spreadsheets. |
| Exports Folder Name | Subfolder for PDF exports. |
| Credits Date / Amount / Property / Unit / Category / Subcategory | Column headers in credit source files. |
| Credits Status / Credits Method / Credits Payer | Optional headers for additional notes. |
| Add Credits Sheet | TRUE to save an audit copy of imported credit rows in the workbook. |

You can cache-bust values by editing them directly; the script reads fresh values each time.

### 4.3 Properties Sheet Columns

| Column | Purpose |
| --- | --- |
| Property | Canonical property name used everywhere else. |
| MAF | Numeric percentage (e.g., `0.10` for 10%) used on property credits. |
| Markup | Percentage applied to `Debits` when `Markup Included` is TRUE in staging. |
| Airbnb | Percentage used to calculate the Airbnb collection fee. |
| Has Airbnb | TRUE if Airbnb totals should be tracked for this property. |
| Admin Fee | Flat admin fee amount added to MAF when admin fee is applied. |
| Admin Fee Enabled | Default boolean indicating whether the admin fee should apply when no override is supplied. |
| Key | Comma-separated list of keywords to match property names that appear differently in imports. |

Maintain this sheet so property mapping, MAF, and markup calculations stay accurate.

## 5. Drive Outputs and Versioning

- **Reports Folder:** Every report run creates a new spreadsheet in the configured reports subfolder. Names increment (`Label`, `Label_2`, …) to avoid overwriting prior runs.
- **Exports Folder:** Each export run creates a new subfolder (`Report Name`, `Report Name_2`, …) containing the PDFs.
- **Temporary Conversions:** Excel credit files are converted to Google Sheets under the hood; the temporary copies are trashed after importing.

## 6. Logs and Traceability

- **Import Log:** Provides proof of which files have been pulled, when, and how many rows were ingested. The system also uses it to skip already-processed files.
- **Report Log:** Lists every report spreadsheet with its version, dates, Drive link, and admin fee choices. Use this when exporting or auditing historical runs.
- **Execution Logs:** The Apps Script execution log captures warnings (missing template data, clean-up issues, etc.). Check this when a menu action reports logged issues.

## 7. Suggested Operational Workflow

1. **Initial setup:** Review Properties, Configuration, and ensure Output folder IDs are correct. Run **Clean Transactions Data** once to catch legacy issues.
2. **Regular cycle:**
   - Import new credits (if applicable).
   - Load staging data for the current pay period; review/edit in **Entry & Edit**.
   - Save staging data to commit changes and let the script handle IDs, markups, and deletes.
   - Update Entry Controls with the reporting period and label.
   - Generate the report and review the resulting spreadsheet tabs.
   - Export PDFs when ready to send to stakeholders.
3. **Maintenance:** Monitor Import/Report logs, refresh property keywords as naming conventions change, and keep configuration IDs up to date if Drive folders move.

By following this guide, operators can confidently manage data intake, reconciliation, reporting, and distribution without interacting with the underlying Apps Script code.
