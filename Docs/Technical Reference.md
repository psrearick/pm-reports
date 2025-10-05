# Technical Reference

This guide documents the pm-reports Apps Script project from a developer perspective. It describes the workbook data model, script modules, execution workflows, shared utilities, and integration points so future changes can be planned with confidence.

## 1. Architectural Overview

- **Platform:** Google Sheets + Google Apps Script bound project. All automation runs inside the spreadsheet container and uses the standard Google service APIs.
- **Data Source of Record:** The `Transactions` sheet holds canonical ledger data. Every workflow reads from or writes to this table.
- **Supporting Tables:** `Entry Controls`, `Entry & Edit`, `Properties`, and `Configuration` drive filtering, staging edits, property metadata, and environment configuration respectively. Hidden sheets (`Import Log`, `Report Log`, and template sheets) provide audit history and templating.
- **Primary Workflows:**
  1. Import credits from Drive files (`Import Credits`).
  2. Stage edits via `Entry & Edit` (`Load Staging Data` / `Save Staging Data`).
  3. Normalise master data (`Clean Transactions Data`).
  4. Build a versioned report spreadsheet (`Generate Report`).
  5. Export per-sheet PDFs (`Export Report`).
- **Drive Output:** Reports and exports are versioned folders/files stored under a configured parent Drive folder.

## 2. Workbook Data Model

| Sheet | Visibility | Purpose | Key Columns |
| ----- | ---------- | ------- | ----------- |
| `Transactions` | Visible | Ledger transactions after staging save/import. | `Transaction ID`, `Date`, `Property`, `Credits`, `Fees`, `Debits`, `Security Deposits`, `Markup Included`, `Markup Revenue`, `Deleted`, `Deleted Timestamp`. |
| `Entry Controls` | Visible | User input for staging/report filters. | `B2` Start, `B3` End, `B4` Property filter, `B5` Report Label, `B6` Show Deleted. |
| `Entry & Edit` | Visible | Staging area refreshed by `loadStagingData()`; includes `Delete Permanently` helper column. | Same headers as `Transactions` plus `Delete Permanently`. |
| `Configuration` | Visible | Key/value pairs consumed by `Configuration.gs`. | `Setting`, `Value`. |
| `Properties` | Visible | Property metadata: financial settings, keywords, and optional report ordering. | `Property`, `MAF`, `Markup`, `Airbnb`, `Has Airbnb`, `Admin Fee`, `Admin Fee Enabled`, `Order`, `Key`. |
| `Import Log` | Hidden | Audit of processed credit files. | Timestamp, File metadata, Rows Imported, Notes. |
| `Report Log` | Hidden | Audit of generated reports. | Timestamp, Report Label, Version, Start/End, Spreadsheet ID/URL, Included Properties, Admin Fee decisions. |
| `ReportBodyTemplate` | Hidden | Property tab template consumed by the template engine. | Placeholder-driven rows/blocks. |
| `ReportTotalsTemplate` | Hidden | Summary tab template. | Placeholder-driven rows/blocks. |
| `ReportAirbnbTemplate` | Hidden | Airbnb tab template. | Placeholder-driven rows/blocks. |

## 3. Configuration and Caching

`Apps Script/Configuration.gs` centralises sheet lookups and caches results in module-level variables (`CONFIG_CACHE`, `PROPERTIES_CACHE`, etc.) to minimise repeated I/O. Callers clear caches implicitly by editing sheet data; reloading the spreadsheet resets the script runtime.

Key helpers:
- `getConfigMap()` – reads `Configuration` into a key/value map.
- `getConfigValue(key, options)` – type-aware fetch with required/default handling.
- `getPropertiesConfig()` – materialises property metadata, including optional numeric `order` used for report tab sequencing.
- `getPropertiesByName()` / `getPropertyByName()` / `resolvePropertyName()` – provide lookups and keyword based resolution for imports.

These functions underpin Drive folder discovery, credit header mapping, and report calculations.

## 4. Module Reference

### 4.1 Menu Entrypoints (`Apps Script/Menu.gs`)
- Registers the `Reports` custom menu via `onOpen()`.
- Each menu item invokes a `handle*` wrapper that calls the core workflow and surfaces toast/alert feedback through `runWithUiFeedback_()`.
- Ensures consistent user messaging, error propagation, and toast duration handling.

### 4.2 Transactions Workflow (`Apps Script/Transactions.gs`)
- `loadStagingData()` pulls filtered master rows (respecting `showDeleted`) into `Entry & Edit`, sorted by date.
- `saveStagingData()` diffs staging rows against master data. It handles new rows (UUID generation), updates (including deleted toggles and timestamp management), soft deletes (implicit when a master row disappears from staging), and permanent deletes.
- `cleanTransactionsData()` normalises IDs, booleans, currency, and dates after manual editing or imports. Issues are collected for logging/UI feedback.
- Utility helpers manage staging controls (`getStagingControls_()`), header index maps, date filtering, and applying markup calculations (`applyMarkupForTransaction_()`).

### 4.3 Reporting (`Apps Script/Reports.gs`)
- `generateReport()` orchestrates the entire report build:
  1. Reads report settings from `Entry Controls` (`getReportSettings_()`).
  2. Collects transactions grouped by property (`collectTransactionsForRange_()`), calculating per-property totals (`calculatePropertyTotals_()`).
  3. Loads property metadata (including `order`) and sorts properties using numeric order first, then alphabetically.
  4. Creates a versioned spreadsheet in the reports folder (`createVersionedSpreadsheet()` from `DriveUtils.gs`).
  5. Renders property, summary, and Airbnb sheets via `renderPropertySheets_()`, `renderSummarySheet_()`, `renderAirbnbSheet_()` which delegate to the template engine.
  6. Deletes the default `Sheet1`, reorders output tabs (Summary → Airbnb → property sheets respecting `order`), and logs metadata to `Report Log` (`logReportGeneration_()`).

- `buildReportData_()` exposes the canonical structure consumed by renderers (`{ startDate, endDate, properties: [...] }`).
- The module encapsulates all derived financial calculations (MAF, admin fees, Airbnb fee, due to owner, etc.).

### 4.4 Credit Imports (`Apps Script/CreditsImport.gs`)
- `importCredits()` resolves Google Sheets and Excel files from configured IDs/folders, skipping already processed files using signatures stored in `Import Log`.
- Converts Excel files to temporary Sheets via `DriveUtils.convertExcelFileToSheet()`, extracts credit records (`extractCreditsRecords_()`), and maps them to transaction rows (`buildTransactionRowFromCredit_()`), leveraging keyword-based property resolution.
- Appends new transactions, optionally emits audit sheets, and records results in `Import Log`.

### 4.5 Exporting (`Apps Script/Exports.gs`)
- `exportReportByLabel()` prompts the user for a label, resolves the latest matching report from `Report Log`, and exports each visible sheet to PDF.
- Uses Drive helpers to create versioned export folders (`createVersionedSubfolder()`), and `fetchWithRetry_()` to handle UrlFetchApp rate limits.

### 4.6 Drive Utilities (`Apps Script/DriveUtils.gs`)
- Wraps DriveApp operations with stricter error messaging and versioned naming (`VERSION_SUFFIX_SEPARATOR`).
- `ensureDestinationFolders()` resolves/creates the configured reports and exports subfolders under the output parent.

### 4.7 Template Engine (`Apps Script/TemplateEngine.gs`)
- Provides declarative templating for report sheets. Supports:
  - Row repeater blocks marked by `[[ROW block]] ... [[ENDROW]]` with optional transforms.
  - Conditional blocks gated by `[[IF flag]]` / `[[ENDIF]]` and inline `[text | if=flag]` snippets.
  - Scalar placeholders (`{TOKEN}`) and flag checks (`[if=flag]`).
- `renderTemplateSheet()` copies a hidden template into the destination spreadsheet, applies the provided `context` (`placeholders`, `rows`, `flags`), collects missing placeholder warnings, and returns the rendered sheet.

### 4.8 Logging Helpers (`Apps Script/Logging.gs`)
- `ensureLogSheet_()` creates headers when missing, preventing hard failures on first run.
- `appendLogRow_()` and `readLogRecords_()` abstract row-level operations against log sheets.

### 4.9 Sheet Utilities (`Apps Script/SheetUtils.gs`)
- Provides header indexing, object-based sheet reads/writes to simplify data transformations.
- Commonly used by staging and configuration routines.

### 4.10 General Utilities (`Apps Script/Utils.gs`)
- Normalisation helpers: `normalizeStringValue_()`, `toBool()`, `toNumber()`, `parseCurrency_()`, `isBlank_()`.
- Date/currency formatting, UUID/time helpers (`utcNow_()`), and spreadsheet access wrappers (`getActiveSpreadsheet_()`, `getSheetByName_()`).
- These utilities are intentionally side-effect free, easing reuse across modules.

## 5. Execution Workflows

### 5.1 Import Credits
1. Menu handler `handleImportCredits_()` → `importCredits()`.
2. Resolve configured files (`getCreditsSourceConfig()`), convert Excel files, locate headers, and build normalized records.
3. Skip already processed files using `(fileId, lastUpdated)` signatures drawn from `Import Log`.
4. Append transactions to `Transactions`, optionally create an audit sheet copy, and log the run.

### 5.2 Load / Save Staging
1. `handleLoadStaging_()` → `loadStagingData()` collects transactions inside the requested window/property and writes them to `Entry & Edit` with a `Delete Permanently` default of `FALSE`.
2. Users edit staging rows. `handleSaveStaging_()` → `saveStagingData()` reconciles back to master:
   - Inserts new transactions with generated IDs.
   - Updates existing rows, recalculating markup and managing delete flags.
   - Applies permanent deletes where requested and implicit soft deletes for removed rows.
   - Writes updates using batch operations (`applyMasterUpdates_()`, `appendTransactions_()`, `applyPermanentDeletes_()`).

### 5.3 Clean Transactions Data
1. `handleCleanTransactions_()` → `cleanTransactionsData()` scans the master sheet.
2. Repairs/normalises IDs, booleans, dates, currencies, and property names via keyword matching.
3. Returns a summary object used for toast messaging and collects issue details (for logging/searching the execution transcript).

### 5.4 Generate Report
1. `handleGenerateReport_()` → `generateReport()`.
2. Pulls active filters/label, fetches transactions grouped by property, and merges property metadata (including order, Airbnb flags, MAF/admin fees).
3. Calculates per-property totals plus aggregated summary metrics.
4. Creates a new Drive spreadsheet (versioned), renders sheets using the template engine, removes the default sheet, and orders visible tabs (Summary → Airbnb → properties ordered by numeric `order` with alphabetical fallback).
5. Logs metadata including properties included and admin fee application to `Report Log`.

### 5.5 Export Report
1. `handleExportReport_()` prompts for a label and resolves the target spreadsheet from `Report Log`.
2. `exportSpreadsheetToFolder_()` creates a versioned export folder inside the configured exports directory.
3. Iterates visible sheets (skipping templates), exporting each to PDF via UrlFetchApp. A retry/delay strategy protects against 429/5xx errors.

## 6. Logging, Auditing, and Error Handling

- Every major workflow writes to a log sheet (`Import Log`, `Report Log`) using helper functions, ensuring traceability.
- UI feedback: menu handlers wrap operations with `runWithUiFeedback_()` to show toast status and present alerts on failure.
- Missing template tokens raise warnings captured by the template engine to facilitate debugging.
- Configuration getters throw explicit errors if required keys/columns are missing, surfacing problems early.

## 7. Drive and Versioning Strategy

- All generated assets use `VERSION_SUFFIX_SEPARATOR` (underscore) to avoid overwrites. `createVersionedSpreadsheet()` and `createVersionedSubfolder()` inspect existing files/folders to determine the next suffix.
- `ensureDestinationFolders()` enforces a stable folder hierarchy: parent output folder → reports subfolder → exports subfolder (customisable via configuration).
- Excel credit files are temporarily copied to Sheets and trashed after import to prevent Drive clutter.

## 8. Template System Details

- Templates live in hidden sheets so designers can modify layout without changing code.
- Developers pass data via the `context` object (`{ placeholders, rows, flags }`).
- Row blocks support optional attributes processed by `transformBlockData_()` inside the engine.
- Inline `["Text" | if=flag]` helpers allow conditional text without removing entire rows. Missing placeholders add warnings to the return payload for logging if desired.

## 9. Extensibility Notes

- **Adding Columns:** Update `TRANSACTION_HEADERS`/`STAGING_HEADERS`, `SheetUtils` writers, and relevant normalization logic.
- **New Property Metadata:** Extend `PROPERTY_COLUMNS`, update `getPropertiesConfig()`, and adjust report rendering if the value needs to appear in templates.
- **Additional Workflows:** Add menu entries in `Menu.gs`, implement core logic in a new module, and follow the toast/error pattern.
- **Template Changes:** Modify hidden template sheets, ensuring placeholder names align with the data structures assembled in `Reports.gs`.
- **Deployment:** Because this is a container-bound Apps Script, deployment typically occurs by sharing the spreadsheet or pushing updates via the Apps Script editor/versioning.

## 10. Key Relationships

- **Menu** → orchestrates → **Transactions**, **Reports**, **Exports**, **CreditsImport**.
- **Transactions** ↔ **Reports** share the `Transactions` data and rely on **Configuration** for metadata.
- **CreditsImport** → uses → **Configuration**, **DriveUtils**, **Transactions** utilities, **Logging**.
- **Reports** → uses → **Configuration**, **Transactions** helpers, **TemplateEngine**, **DriveUtils**, **Logging**.
- **Exports** → uses → **DriveUtils**, `Report Log`, **Utils** for sanitisation.
- **TemplateEngine**, **Utils**, **SheetUtils** form the shared foundation across modules.

Understanding these dependencies helps plan changes without breaking other workflows. Start by updating shared utilities, then adjust dependent modules, and finally update the documentation and templates to match new behaviour.
