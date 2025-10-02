# Property Management & Reporting System - Design Document

## System Overview

This document outlines the architecture of a custom Google Sheets-based system designed to automate property management accounting and reporting. The system's primary goals are to:

- Ensure data integrity through a single, master transaction database.
- Provide a simple, user-friendly interface for manual data entry and editing.
- Automate the creation of flexible, period-based financial reports for individual properties.
- Automate the generation of a summary report.
- Provide functionality to export final reports to PDF for easy sharing and archiving.
- Maintain a historical log of all generated reports.

The system is comprised of several interconnected Google Sheets and a Google Apps Script project that acts as the central controller.

## Core Components (The Sheets)

The system is built upon a set of specialized sheets within a single Google Sheets document:

- **`Transactions (Master)`**: This is the single source of truth. It is a permanent, ever-growing database of all transactions for all properties. It should only be modified by the script to ensure data integrity. A key feature is the `Transaction ID` column, which contains a unique identifier for every entry, preventing duplicates and enabling precise updates.
- **`Entry & Edit (Staging)`**: This is the user's primary workspace. It acts as a temporary, focused "window" into the master `Transactions` sheet. The user will load a specific subset of data (based on property and date range) into this sheet to perform manual additions, deletions, and edits without risk of corrupting the master data.
- **`Properties`**: A configuration sheet that holds all property-specific information, such as the official property name, MAF rate, Markup percentage, and keywords for matching.
- **`Configuration`**: A sheet for storing script settings and global variables. This includes column headers for credit imports and the names of template sheets, making the script easier to maintain.
- **`Report Log`**: A permanent, automated log of every report that is generated. It will store the report's name, the date it was created, the period it covers, and a direct link to the generated Google Sheet file.
- **`ReportBodyTemplate`, `ReportTotalsTemplate`, `ReportAirbnbTemplate`**: These are hidden sheets that contain the layout, formatting, and placeholder text for the final report documents.

## User Workflows

The system is designed around four primary workflows:

**Workflow 1: Importing New Credits**

1. The user places the exported `.xlsx` file from the payment processor into a specific Google Drive folder.
2. The user runs the "Import Credits" function from a custom menu.
3. The script reads the `.xlsx` file and, for each row, generates a unique key (e.g., combining date, amount, and property).
4. It checks this key against the existing `Transaction ID`s in the `Transactions (Master)` sheet to identify duplicates.
5. Only new, non-duplicate transactions are appended to the `Transactions (Master)` sheet, with a new, unique `Transaction ID` generated for each.

**Workflow 2: Manual Data Entry & Editing**

1. The user navigates to the `Entry & Edit (Staging)` sheet.
2. They enter a Start Date, End Date, and select a Property from a dropdown menu.
3. They click the "Load Data" button.
4. The script finds all matching records in `Transactions (Master)` and copies them to the staging sheet. It also secretly stores the master row's unique `Transaction ID` in a hidden column.
5. The user can now safely add new rows, delete rows, or modify existing rows in this focused view.
6. Once finished, the user clicks the "Save Changes to Master" button.
7. The script reads the staging sheet and syncs all changes back to the `Transactions (Master)` sheet, using the hidden `Transaction ID` to update existing records and adding new records as needed.

**Workflow 3: Generating a Final Report**

1. After the data in the `Entry & Edit (Staging)` sheet is finalized and saved, the user enters a descriptive name for the report in the "Report Label" field (e.g., "Q3 2025 Owner Payouts").
2. They click the "Generate Report" button.
3. The script reads the clean data _directly from the `Entry & Edit (Staging)` sheet_.
4. It creates a new, separate Google Sheets document in Google Drive, named with the report label.
5. It builds the report inside this new file, creating a formatted sheet for each property with relevant data, plus a "Summary Page" that pulls totals from each property sheet.
6. Finally, it adds a new entry to the `Report Log` sheet with the label and a link to the new file.

**Workflow 4: Exporting a Report to PDF**

1. The user opens a generated report spreadsheet (either immediately after creation or later via the link in the `Report Log`).
2. The script will have automatically added a custom menu to this new spreadsheet (e.g., "Export").
3. The user clicks Export > Export All to PDF.
4. The script, running from the master file but acting on the _active_ spreadsheet, creates a new folder in Google Drive named after the report.
5. It loops through every sheet in the report and saves a separate, formatted PDF of each one into that new folder.

# Spreadsheet Setup Instructions

The script will rely on these exact names and layouts.

## Group 1: Core Data Sheets

**`Transactions (Master)` Sheet**

- Create a sheet named `Transactions (Master)`.
- Set up the following headers in the first row, starting in cell A1:
    - `A1`: Transaction ID
    - `B1`: Date
    - `C1`: Property
    - `D1`: Unit
    - `E1`: Credits
    - `F1`: Fees
    - `G1`: Debits
    - `H1`: Security Deposits
    - `I1`: Debit/Credit Explanation
    - `J1`: Markup Included _(This cell should be formatted as a checkbox: Insert > Checkbox)_
    - `K1`: Markup Revenue
    - `L1`: Internal Notes

**`Properties` Sheet**

- Create a sheet named `Properties`.
- Set up the following headers in the first row:
    - `A1`: Property
    - `B1`: MAF
    - `C1`: Markup
    - `D1`: Airbnb
    - `E1`: Has Airbnb _(Format as checkbox)_
    - `F1`: Key
- Populate this sheet with property data.

## Group 2: Control & Interface Sheets

**`Entry & Edit (Staging)` Sheet**

- Create a sheet named `Entry & Edit (Staging)`.
- **Copy the headers** from `A1:L1` of the `Transactions (Master)` sheet and paste them into `A1:L1` of this sheet.
- **Set up the control panel:**
    - In cell `N2`, type the label: Start Date
    - In cell `N3`, type the label: End Date
    - In cell `N4`, type the label: Property
    - In cell `N5`, type the label: Report Label
    - In cell `O2`, enter the start date.
    - In cell `O3`, enter the end date.
    - In cell `O5`, enter the name for the final report file.
    - *For cell `O4` (Property Dropdown):*
        - Click on cell `O4`.
        - Go to the menu: Data > Data validation.
        - Click \+ Add rule.
        - Under "Criteria," choose Dropdown (from a range).
        - Click the select data range icon and choose the column containing the property names from the `Properties` sheet (e.g., `Properties!A2:A`).
        - Click Done.

**4`Configuration` Sheet**

- Create a sheet named `Configuration`.
- Set up two columns, `A` and `B`, for keys and values.
    - `A1`: Setting
    - `B1`: Value
    - `A2`: Credits Date
    - `B2`: _(Type the exact header name for Date from the credits XLSX file here)_
    - `A3`: Credits Amount
    - `B3`: _(Type the exact header name for Amount from the credits XLSX file here)_
    - ...continue this for all the column headers to map from the credits file.

**5`Report Log` Sheet**

- Create a sheet named `Report Log`.
- Set up the following headers in the first row:
    - `A1`: Timestamp
    - `B1`: Report Label
    - `C1`: Start Date
    - `D1`: End Date
    - `E1`: Link to Report
    - `F1`: Spreadsheet ID

## Group 3: Hidden Template Sheets

**6The Three Template Sheets**

- Create three new sheets with the exact names:
    1. `ReportBodyTemplate`
    2. `ReportTotalsTemplate`
    3. `ReportAirbnbTemplate`
- Copy and paste the content and formatting from the template CSVs into these corresponding sheets. Make sure to include the placeholder text like `[PROPERTY NAME]` and `[UNIT NUMBER]`.
- Once they are set up correctly, right-click on their tabs and select "Hide sheet" to keep your workspace clean. The script will still be able to find and use them.
