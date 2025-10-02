# Property Management System Specification

## Overview

This document captures the agreed-upon behavior for the automated Google Sheets workflow that ingests credits, manages transactions, and produces property reports.

## Configuration Data

- Configuration entries live in `Transactions Configuration` (Google Sheet or CSV seeded from `Data/Transactions Configuration.csv`).
- Required keys (string match, case-sensitive):
  - `Credits Document` (optional if `Credits Folder ID` provided).
  - `Credits Folder ID` (optional, enables multi-file imports when supplied).
  - `Output Folder ID` (required; Drive folder where all generated assets reside).
  - `Reports Folder Name` (optional subfolder inside the Output folder for generated Google Sheets reports).
  - `Exports Folder Name` (optional subfolder inside the Output folder for PDF exports).
  - Additional header-mapping keys such as `Credits Date`, `Credits Amount`, `Credits Property`, etc., are mandatory for import parsing.
- Properties metadata is managed via the `Transactions Properties` sheet (seeded by `Data/Transactions Properties.csv`) with columns:
  - `Property`: canonical name used across the system.
  - `MAF`: decimal rate used in owner/PM calculations (0 for properties without % fee).
  - `Markup`: decimal vendor markup rate per property.
  - `Airbnb`: decimal percentage applied to Airbnb totals when `Has Airbnb` is true.
  - `Has Airbnb`: boolean flag enabling Airbnb reporting.
  - `Admin Fee`: flat-dollar fee applied in addition to MAF totals when enabled (new column).
  - `Admin Fee Enabled`: boolean default for including the Admin Fee during report generation (new column).
  - `Key`: comma-separated lowercase keywords for matching import rows to properties.

## Credits Import Workflow

- Imports can target a single file ID (`Credits Document`) or all spreadsheets stored in a Drive folder (`Credits Folder ID`).
- When a folder is specified, every `.xlsx` file found is processed; the script keeps a log of processed file IDs and modification timestamps to prevent duplicate ingestion.
- Each `.xlsx` is converted to a temporary Google Sheet for parsing. After successful import the temporary sheet is deleted, while the original `.xlsx` remains in place for manual housekeeping.
- Imported rows generate UUID `Transaction ID`s. A deterministic hash is not required because duplicate suppression relies on the processed-files log and existing transaction IDs in the master sheet.
- Header names in the credits file are mapped via configuration keys (e.g., `Credits Date`). If required keys are missing, the import aborts with a user-facing message.

## Transactions & Staging Workflow

- `Transactions (Master)` holds the canonical dataset. Every record carries:
  - `Transaction ID` (UUID), `Deleted` (boolean flag), `Deleted Timestamp`, etc., with additional hidden columns to support soft deletes.
- `Entry & Edit (Staging)` loads a filtered view based on property/date and includes the following behaviors:
  - Non-deleted rows appear by default.
  - A "Show Deleted" checkbox toggles visibility of soft-deleted rows; when visible they display as checked in the `Deleted` column.
  - The staging grid includes two checkbox columns: `Deleted` (soft delete) and `Delete Permanently` (hard delete on sync).
  - Removing a row from the staging sheet during editing toggles the `Deleted` flag when changes are saved.
  - Unchecking `Deleted` while saving reinstates the record in the master sheet.
- Airbnb detection uses the `Internal Notes` field; if a property has `Has Airbnb = TRUE` and the notes contain `airbnb` (case-insensitive), the amount contributes to the Airbnb total for that property.

## Reporting & Template System

- Template sheets (`ReportBodyTemplate`, `ReportTotalsTemplate`, `ReportAirbnbTemplate`) contain placeholder tokens that drive report generation.
- Supported syntax:
  - Scalar replacement: `[PLACEHOLDER]` substitutes single values, e.g., `[PROPERTY NAME]`.
  - Repeating row blocks: rows wrapped between `[[ROW block_name]]` and `[[ENDROW]]` are duplicated per dataset item. Within those rows, tokens use braces (e.g., `{unit}`, `{credits}`) that map to transaction fields.
  - Conditional blocks: `[[IF condition]]` ... `[[ENDIF]]` render only when the provided context flag evaluates to truthy (e.g., `[[IF has_airbnb]]`).
- Adding new placeholders requires only updating template sheets; the script will attempt to replace missing tokens gracefully and warn if unresolved placeholders remain.

## Report Generation & Export

- Report creation reads from `Entry & Edit (Staging)` after changes are saved to master.
- A new Google Sheets report is created under the Output folder (or inside the `Reports Folder Name` subfolder if configured).
- Soft-deleted records are excluded from report data by default; permanently deleted rows are removed entirely.
- Admin fees: during report generation a user-facing control determines whether to apply each property's `Admin Fee`. By default the checkbox state uses the `Admin Fee Enabled` value from the properties sheet.
- Versioned naming: when generating report files or export folders, if a file named `<Report Label>` already exists, the script appends `_2`. If `_2` already exists, it increments the highest suffix (`_3`, `_4`, etc.) without nesting.

## Export to PDF

- Export routines create (or reuse) a folder inside the Output directory based on the configured `Exports Folder Name` (or the Output folder directly if blank).
- For each report spreadsheet, an export subfolder named after the report label (with the same versioning rules) stores generated PDFs: one per property and a `Summary` file.
- Temporary Drive files created during export are cleaned up once PDFs are saved.

## Logging & Notifications

- Processed credit files (ID, name, timestamp, rows imported) are appended to an `Import Log` sheet in the master workbook to avoid double processing.
- Report generation events append to `Report Log` with additional metadata (version suffixes, admin fee state, row counts).
- Errors raise user-visible alerts (toast or modal). Success notifications are optional and can be added later without architectural changes.

