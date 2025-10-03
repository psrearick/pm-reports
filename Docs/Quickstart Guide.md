# Quickstart Guide

Follow these steps to initialise a fresh copy of the Reports workbook and run your first reporting cycle. This guide assumes you received the spreadsheet with the Apps Script project already attached, but the data tabs are empty.

## 1. Prerequisites

-   Access to Google Drive and Google Sheets.
-   Output folders where the generated reports and exports will live.
-   (Optional) Historical transaction data and/or credit export files you plan to import.

## 2. Review the Workbook Layout

1. Locate the key sheets: **Transactions**, **Entry & Edit**, **Entry Controls**, **Properties**, **Configuration**, the hidden logs (**Import Log** and **Report Log**), and the hidden template sheets (**ReportBodyTemplate**, **ReportTotalsTemplate**, **ReportAirbnbTemplate**).
2. Familiarise yourself with the **Reports** custom menu. All automations are triggered from here.

## 3. Prepare Google Drive

1. Create (or identify) a parent Drive folder that will house your outputs.
2. Inside that folder, create two subfolders (names are up to you): one for generated report spreadsheets, one for exported PDFs.
3. If you will import credits automatically, gather the file IDs or place the source spreadsheets inside a dedicated Drive folder.

## 4. Configure the Workbook

Open the **Configuration** sheet and fill in the `Value` column for each relevant `Setting`:

-   **Output Folder ID:** Paste the ID of the parent Drive folder you created.
-   **Reports Folder Name / Exports Folder Name:** The subfolder names (leave blank to use the parent folder directly, but separate folders are recommended).
-   **Credits Document / Credits Folder ID:** Either point at a single source file or a folder that contains all credit exports. You can set both.
-   **Credits … headers:** Ensure every column label in the source credit files is listed exactly as it appears (Date, Amount, Property, Unit, Category, Subcategory, plus optional Status/Method/Payer).
-   **Add Credits Sheet:** Set to `TRUE` if you want the importer to create an audit tab for each file processed.

## 5. Update Property Settings

1. On the **Properties** sheet, verify every property you manage is listed with accurate settings:
    - **MAF** percentage (e.g. `0.08` for 8%).
    - **Markup** percentage for work orders where `Markup Included` will be TRUE.
    - **Airbnb** percentage and **Has Airbnb** flag if applicable.
    - **Admin Fee** amount and **Admin Fee Enabled** flag.
    - **Key** column keywords to recognise the property in imports (include common abbreviations and alternative names).
2. Remove properties you no longer manage or add new rows for new properties.

## 6. Seed the Transactions Table

1. If you have historical data, copy it into **Transactions** with the exact header order listed in row 1 (`Transaction ID`, `Date`, `Property`, …, `Deleted Timestamp`).
2. If you are starting from scratch, leave only the header row; you can add new items via staging.
3. Run **Reports → Clean Transactions Data**. This step will:
    - Generate UUIDs for blank Transaction IDs.
    - Normalise property names using the keywords you provided.
    - Standardise currency and boolean fields.
    - Log any issues that need manual attention (check _Extensions → Apps Script → Executions_ for details).

## 7. Optional: Import Latest Credits

If you maintain rent/credit exports in Drive, run **Reports → Import Credits**. The importer will:

-   Process each configured file or folder.
-   Append new rows to **Transactions**.
-   Skip files already logged in **Import Log**.
-   Create an audit sheet for each run if `Add Credits Sheet` is TRUE.

Run **Clean Transactions Data** afterwards to align formatting.

## 8. Configure Entry Controls for the Reporting Period

On the **Entry Controls** sheet:

1. Set **Start Date (B2)** and **End Date (B3)** for the period you are reviewing.
2. Leave **Property (B4)** blank to work on all properties, or enter a single property name to isolate it.
3. Enter a **Report Label (B5)** – this will become the name of the generated report.
4. Decide whether to **Show Deleted (B6)** transactions in staging.
5. (Optional) Set **Admin Fee Override (B7)** to TRUE or FALSE if you want to override the per-property default for the next report.

## 9. Edit Transactions via Staging

1. Choose **Reports → Load Staging Data**. The **Entry & Edit** sheet is refreshed with the filtered transactions (sorted by date).
2. Make any adjustments:
    - Update amounts, notes, units, explanations.
    - Toggle `Markup Included` and `Deleted` as needed.
    - Mark `Delete Permanently` TRUE to remove a transaction entirely.
    - Add new rows at the bottom (leave Transaction ID blank – the script will generate one).
3. When finished, run **Reports → Save Staging Data**. The script writes updates back to **Transactions**, handling new rows, soft deletes, hard deletes, and markup recalculation.
4. Reload staging if you want to verify the results.

## 10. Generate and Review the Report

1. Confirm Entry Controls still reflect the desired period and label.
2. Choose **Reports → Generate Report**. The script will create a new spreadsheet in the configured Reports folder, populate Summary/Airbnb/Property tabs, reorder them (Summary first, Airbnb second, property tabs after), and log the run.
3. Open the generated spreadsheet from the toast notification or by visiting the Reports folder / Report Log.
4. Inspect the tabs to confirm totals, formatting, and template output.

## 11. Export PDFs (Optional)

1. When ready to send statements, pick **Reports → Export Report**.
2. Enter the report label you want to export (leave blank for the most recent entry in **Report Log**).
3. PDF files for each tab are created in a new versioned subfolder inside your Exports folder. The script automatically spaces requests to avoid Google’s 429 rate limits.

## 12. Ongoing Maintenance Checklist

-   Update the **Properties** keywords whenever naming conventions change in source data.
-   Periodically run **Clean Transactions Data** to keep the master table tidy.
-   Monitor **Import Log** and **Report Log** for audit purposes.
-   Customise the template sheets (see “Template System Guide”) when the statement layout needs to change.
-   Add the reference page content (see “Reference Page Content”) to an in-workbook tab so daily operators have quick instructions.

Once these steps are complete you can repeat sections 8–11 for each reporting cycle with confidence that the automation and templates accurately reflect your configuration.
