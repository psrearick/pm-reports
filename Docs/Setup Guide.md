# Google Drive Setup Guide

This guide walks through creating the entire property management reporting system from an empty Google Drive folder. Follow each step in order; the Apps Script project assumes the sheet names, headers, and configuration values exactly match what is described here.

## 1. Create the Drive Structure

1. In Google Drive, create a new top-level folder (e.g., Reports). This is the `Output Folder ID` that the configuration will reference.
2. Create two optional subfolders (you can name them differently if desired):

    - Reports – the script stores generated Google Sheets reports here.
    - Exports – the script stores PDF exports here.

    If you skip either subfolder, leave the related configuration value blank later.

3. Record the folder IDs:
    - Right-click each folder > Copy link. The alphanumeric portion between `/folders/` and any query string is the Drive ID.

## 2. Create the Master Spreadsheet

1. In the top-level folder, create a new Google Sheets file named Transactions Master (you can rename later; only sheet/tab names must match the guide).
2. Open the spreadsheet and rename the default `Sheet1` tab to Transactions (Master).

## 3. Build the Core Sheets

### 3.1 Transactions (Master)

1. In `Transactions (Master)`, enter the following headers in row 1, starting at cell A1:
    - A1 `Transaction ID`
    - B1 `Date`
    - C1 `Property`
    - D1 `Unit`
    - E1 `Credits`
    - F1 `Fees`
    - G1 `Debits`
    - H1 `Security Deposits`
    - I1 `Debit/Credit Explanation`
    - J1 `Markup Included`
    - K1 `Markup Revenue`
    - L1 `Internal Notes`
    - M1 `Deleted`
    - N1 `Deleted Timestamp`
2. Format column `J` as checkboxes (Insert > Checkbox).
3. Format column `M` as checkboxes.
4. Leave the rest of the sheet empty for now; the script will populate rows.

### 3.2 Entry & Edit (Staging)

1. Add a new sheet named Entry & Edit (Staging).
2. Copy the headers A1:N1 from `Transactions (Master)` and paste into row 1 (cells A1–N1).
3. In cell O1, type `Delete Permanently`. The script fills this column with checkboxes when data is loaded.
4. Leave the remainder of the sheet empty; the script will only clear and write within columns A–O, leaving anything to the right (notes, helper formulas, etc.) untouched.

### 3.3 Entry Controls

1. Add a sheet named Entry Controls.
2. Enter the following labels in column A, leaving column B blank for user input:
    - A2 `Start Date`
    - A3 `End Date`
    - A4 `Property`
    - A5 `Report Label`
    - A6 `Show Deleted`
3. Format cells B6 and B7 as checkboxes.
4. (Optional) Apply a dropdown on B4 that references `Properties!A2:A` for faster property selection.
5. The script reads control values from column B. Leave the Property cell blank to load/edit/export all properties; enter a property name to work with that property only.

### 3.4 Properties

1. Add a sheet named Properties.
2. Enter the headers in row 1:
    - A1 `Property`
    - B1 `MAF`
    - C1 `Markup`
    - D1 `Airbnb`
    - E1 `Has Airbnb`
    - F1 `Admin Fee`
    - G1 `Admin Fee Enabled`
    - H1 `Key`
3. Populate rows with property data. Example values based on `Data/Transactions Properties.csv`:

---

| Property                | MAF  | Markup | Airbnb | Has Airbnb | Admin Fee | Admin Fee Enabled | Key           |
| ----------------------- | ---- | ------ | ------ | ---------- | --------- | ----------------- | ------------- |
| 1505-1515 Franklin Park | 0.06 | 0.1    | 0      | FALSE      | 0         | FALSE             | franklin,park |
| 189 W Patterson Ave     | 0.07 | 0.1    | 0      | FALSE      | 0         | FALSE             | patterson     |
| 196 Miller Ave          | 0.06 | 0.1    | 0      | FALSE      | 0         | FALSE             | miller        |
| 22 Wilson Ave           | 0.08 | 0.1    | 0      | FALSE      | 0         | FALSE             | wilson        |
| 2536 Adams Ave          | 0    | 0.1    | 0      | FALSE      | 15        | TRUE              | adams         |
| 705 Ann                 | 0.02 | 0.1    | 0      | FALSE      | 0         | FALSE             | ann           |
| Ohio & Bryden           | 0.06 | 0.1    | 0      | FALSE      | 0         | FALSE             | ohio,bryden   |
| Park and State          | 0.08 | 0      | 0.04   | TRUE       | 0         | FALSE             | park,state    |
| Schiller Terrace        | 0.07 | 0.1    | 0      | FALSE      | 0         | FALSE             | schiller      |
| 1476 S High St          | 0.08 | 0.1    | 0      | FALSE      | 0         | FALSE             | high          |

---

4. Format `MAF`, `Markup`, `Airbnb` as decimals; `Has Airbnb` and `Admin Fee Enabled` as checkboxes; `Admin Fee` as currency (optional).

### 3.5 Configuration

1. Add a sheet named Configuration.
2. Enter headers: `A1` `Setting`, `B1` `Value`.
3. Add configuration keys and fill in the `Value` column:

---

| Setting             | Value (example)                                  |
| ------------------- | ------------------------------------------------ |
| Credits Document    | _(optional)_ – single spreadsheet ID             |
| Credits Folder ID   | _(optional)_ – Drive folder ID for .xlsx imports |
| Credits Date        | `Date paid`                                      |
| Credits Amount      | `Amount paid`                                    |
| Credits Property    | `Property name`                                  |
| Credits Unit        | `Unit`                                           |
| Credits Category    | `Category`                                       |
| Credits Subcategory | `Sub-category`                                   |
| Credits Status      | `Payment status` _(or leave blank for default)_  |
| Credits Method      | `Payment method` _(or blank)_                    |
| Credits Payer       | `Payer / Payee` _(or blank)_                     |
| Add Credits Sheet   | `FALSE` (or `TRUE` to add audit sheets)          |
| Output Folder ID    | _(required)_ – ID from step 1                    |
| Reports Folder Name | `Reports` (or blank if not using subfolder)      |
| Exports Folder Name | `Exports` (or blank)                             |

---

4. Ensure `Add Credits Sheet` is a plain text `TRUE/FALSE` string (or use checkbox, both work).

### 3.6 Report Log

1. Add a sheet named Report Log.
2. Enter headers row 1 (A1–I1):

    - `Timestamp`
    - `Report Label`
    - `Version`
    - `Start Date`
    - `End Date`
    - `Spreadsheet ID`
    - `Report URL`
    - `Properties Included`

    (The script will append rows automatically.)

### 3.7 Import Log

1. Add a sheet named Import Log.
2. Enter headers row 1 (A1–F1):
    - `Timestamp`
    - `File ID`
    - `File Name`
    - `Last Modified`
    - `Rows Imported`
    - `Notes`

## 4. Create Template Sheets

Create three hidden template tabs. These must match the names expected by the script.

### 4.1 ReportBodyTemplate

1. Insert a new sheet named ReportBodyTemplate.
2. Design the layout, then include placeholder markers:
    - Place scalar tokens like `[PROPERTY NAME]`, `[REPORT LABEL]`, `[REPORT PERIOD]` where needed.
    - Add a transaction table with the following pattern:
        1. A row containing `[[ROW transactions]]` in the first cell (no additional text or spaces anywhere else on that row).
        2. One or more template rows that contain cell-level tokens such as `{unit}`, `{credits}`, `{fees}`, `{debits}`, `{securityDeposits}`, `{date}`, `{explanation}`, `{markupIncluded}`, `{markupRevenue}`, `{internalNotes}`.
        3. A row containing `[[ENDROW]]` in the first cell (again, with no other content on that row).
    - For Airbnb-only sections, wrap them in conditional markers:
        - Row with `[[IF has_airbnb]]`
        - Content rows referencing `[AIRBNB TOTAL]`, `[AIRBNB FEE]`
        - Row with `[[ENDIF]]`
3. Once configured, right-click > Hide sheet.

### 4.2 ReportTotalsTemplate

1. Create a sheet named ReportTotalsTemplate.
2. Include placeholders for summary values such as `[REPORT LABEL]`, `[REPORT PERIOD]`.
3. For the property summary table:
    - Row containing `[[ROW summary]]` (token must be the only content in that row)
    - Template row cells referencing `{property}`, `{dueToOwners}`, `{totalToPm}`, `{totalFees}`, `{newLeaseFees}`, `{renewalFees}`
    - Row containing `[[ENDROW]]`
4. Hide the sheet when finished.

### 4.3 ReportAirbnbTemplate

1. Create a sheet named ReportAirbnbTemplate.
2. Add scalar placeholders for `[REPORT LABEL]`, `[REPORT PERIOD]`.
3. Add Airbnb totals table:
    - `[[ROW airbnb]]` (only token in the row)
    - Template row referencing `{property}`, `{income}`, `{collectionFee}`
    - `[[ENDROW]]`
4. Hide the sheet when finished.

> Tip: If you want to base templates on CSV examples (`Docs/Report Templates.md`), paste the content, then replace static values with the tokens described above.

## 5. Launch the Apps Script Project

1. In the spreadsheet, click Extensions > Apps Script.
2. Delete any placeholder files (e.g., `Code.gs`) if present.
3. Create new script files with the exact filenames and paste the corresponding code from the repository:
    - `Constants.gs`
    - `Utils.gs`
    - `Configuration.gs`
    - `DriveUtils.gs`
    - `SheetUtils.gs`
    - `Logging.gs`
    - `TemplateEngine.gs`
    - `CreditsImport.gs`
    - `Transactions.gs`
    - `Reports.gs`
    - `Exports.gs`
    - `Menu.gs`
4. In Services (left sidebar) enable the Google Drive Advanced Service:
    - Click the `+` button near Services.
    - add Drive API (v2). The script uses `Drive.Files.copy` for Excel conversion.
5. Save the project, then click Run > onOpen once to populate menus (authorize when prompted).

## 6. Populate Initial Data (Optional)

1. If you have historical credit or transaction CSVs, import them manually into `Transactions (Master)` (ensuring all columns are filled). Alternatively, run the Apps Script import after setting configuration values.
2. Staging sheet will be blank until you click Reports > Load Staging Data.

## 7. Configure Credits Import Sources

1. Choose whether to import from a single Google Sheet or from `.xlsx` files in a Drive folder:
    - Single Sheet: paste the spreadsheet ID into `Configuration!B2` (`Credits Document`).
    - Folder: paste the folder ID into `Configuration!B3` (`Credits Folder ID`). The script will process all `.xlsx` files in that folder.
2. Leave the unused option blank. You can supply both; the script will process each unique file.
3. If `Add Credits Sheet` is `TRUE`, the master spreadsheet will include a copy of each import for auditing (sheet names prefixed with `Credits`).

## 8. Running the Workflow

1. Use the custom menu Reports:
    - Import Credits – Processes configured files, appends new transactions, and logs results.
    - Load Staging Data – Reads the controls on **Entry Controls**. Leave `Property` blank to load all properties, or select a single property to focus the staging grid.
    - Save Staging Data – Pushes changes back to the master sheet for the filtered scope (all properties or the single property).
    - Generate Report – Creates a versioned Google Sheets report using the same control values and logs the run.
    - Export Report (from log) – Prompts for a report label (or uses the most recent entry) and creates PDFs in the configured `Exports` folder.
2. In generated reports, you can use Google Sheets’ built-in **File > Download > PDF** action, or rely on the master workbook’s Export command to generate PDFs.

## 9. Logs & Maintenance

-   **Import Log**: Review processed files and re-run imports if you add new `.xlsx` files (modify or delete rows if you need to force re-import).
-   **Report Log**: Track generated reports, spreadsheet IDs, and admin fee decisions. Use the logged URLs to reopen reports quickly.
-   Periodically clear or archive old `Credits` audit sheets if `Add Credits Sheet` is enabled.
-   Keep template sheets up to date with any layout changes; update placeholders as needed (no code changes required for new tokens that map to existing data fields).

## 10. Final Checklist

-   [ ] Output folder ID set in `Configuration`.
-   [ ] Reports/Exports subfolder names (if used) match the Drive structure.
-   [ ] Properties sheet contains all active properties with correct rates, admin fees, and keywords.
-   [ ] Template sheets include the correct markers (`[[ROW ...]]`, `{field}`, `[PLACEHOLDER]`).
-   [ ] Drive Advanced Service enabled in Apps Script.
-   [ ] Custom menu appears (run `onOpen` if not).
-   [ ] Optional: add time-driven triggers for automatic imports or scheduled exports, if desired.

With these steps completed, the system is ready to ingest credit files, manage transactions with staging safeguards, and generate templated property reports and PDFs.
