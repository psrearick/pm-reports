# PM Reports Apps Script

PM Reports is a Google Sheets + Apps Script solution for managing property management financials. It centralises transaction intake, staging, reporting, and PDF exports while maintaining a full audit trail of generated outputs.

## Features

- Import credit/receipt spreadsheets from Drive and normalise them into a master ledger.
- Stage, edit, and reconcile transactions with automatic ID management and markup recalculation.
- Clean the `Transactions` sheet to standardise dates, currency, booleans, and property naming.
- Generate versioned owner statement workbooks (Summary, Airbnb, per-property tabs) using template-driven rendering.
- Export report tabs to individual PDFs with retry-friendly batching.
- Log every import and report run for auditability.

## Repository Layout

```
Apps Script/       Core Apps Script source modules
Docs/              End-user and developer documentation
pm-reports.code-workspace  VS Code workspace definition
README.md          Project overview (this file)
```

### Key Script Modules

- `Menu.gs` – Registers the custom menu and routes user actions with consistent UI feedback.
- `Transactions.gs` – Handles staging load/save flows, transaction normalisation, and clean-up.
- `Reports.gs` – Builds report data, renders tabs via templates, and logs report runs.
- `CreditsImport.gs` – Imports Drive-based credit files, converts Excel to Sheets, and appends ledger rows.
- `Exports.gs` – Exports generated reports to PDFs and stores them in versioned Drive folders.
- `Configuration.gs`, `DriveUtils.gs`, `TemplateEngine.gs`, `Utils.gs` – Shared helpers for configuration, Drive operations, templating, and primitive utilities.

Full module descriptions are available in `Docs/Technical Reference.md`.

## Getting Started

1. Open the bound Google Sheet that hosts this Apps Script project.
2. Populate the `Configuration` sheet with required Drive IDs and header mappings.
3. Fill out the `Properties` sheet, including the optional `Order` column to control report tab sequencing.
4. Review the workflow in `Docs/Quickstart Guide.md` for operational setup.
5. Use the **Reports** custom menu to import credits, stage edits, clean data, generate reports, and export PDFs.

For detailed operator guidance, see `Docs/User Guide.md`. For a quick overview, refer to `Docs/Reference Page.md`.

## Development Notes

- Scripts rely on Google Apps Script services (`SpreadsheetApp`, `DriveApp`, `UrlFetchApp`, etc.) and run within the spreadsheet container. No external dependencies are required.
- Configuration and property metadata are cached per Apps Script execution context; editing the underlying sheets invalidates those caches automatically.
- Generated reports and exports use the underscore version suffix (e.g., `Report_2`) to prevent overwrites.
- Templates live in hidden sheets (`ReportBodyTemplate`, `ReportTotalsTemplate`, `ReportAirbnbTemplate`) and are rendered using the custom template engine (`TemplateEngine.gs`).

## Documentation

- `Docs/Quickstart Guide.md` – Initial setup and first-run checklist.
- `Docs/User Guide.md` – Day-to-day operational instructions.
- `Docs/Reference Page.md` – On-sheet cheat sheet for operators.
- `Docs/Technical Reference.md` – Architecture, module reference, and developer-centric notes.

Keep these documents updated alongside code changes to ensure both operators and developers stay aligned.
