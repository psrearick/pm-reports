## Current Process

Month ends on 19th

1. Create `Transactions` spreadsheet for past month from template.
2. Around the 19th enter all transactions into the `Transactions` spreadsheet.
3. Download the `Credits` Payment Report spreadsheet from payment processor.
4. Type data from `Credits` into `Transactions` spreadsheet (Cannot copy/paste as formatting and column headers do not match).
5. Open `Income Statement` workbook for each property and duplicate the template worksheet and rename to the past month (ex. "Sep25").
6. Copy all transactions from `Transactions` into `Income Statement` for each property.
7. Export each `Income Statement` worksheet as a `.pdf`.
8. Create a `Summary Page` worksheet for the past month.
9. Copy the calculated values from each `Income Statement` worksheet into the `Summary Page`.
10. Export the `Summary Page` and a `.pdf`.

## Ideal Process

1. Create `Transactions` spreadsheet for past month from template.
2. Around the 19th enter all transactions into the `Transactions` spreadsheet.
3. Download the `Credits` Payment Report spreadsheet from payment processor.
4. Click button to generate reports, providing `Credits` file path as a target (or placing it in a specific directory with a specific filename to match a pattern).

## Input and Output

**Input:**

- Exported `.xlsx` file of all payments received from tenants: `Credits`
- Manually entered transaction data: `Transactions`

**Output:**

- A `.pdf` file for each property listing transactions and providing summary data: `Income Statement`
- A consolidation of summary data from each `Income Statement`: `Summary Page`

## Credits

- Mostly received from payment processor, exported as `.xlsx` from user defined report period
- Income from processed transaction paid by tenants

### Raw Data

- Date paid
- Amount paid
- Property name
- Unit
- Category
- Sub-category
- Payment status
- Payment method
- Payer/Payee

Category -> Sub-category examples:

    - Tenant charges & fees -> Late payment fee
    - Property General Expense -> Transaction Fee
    - Rent
    - Rent -> Prorated rent
    - Rent -> Pet rent
    - Rent -> Standard rent
    - Deposit

## Transactions

Manually entered data

### Raw

- Property
- Unit
- Fees
- Debit
- Security Deposits
- Date
- Debit/Credit Explanation
- Markup Included
- Internal Notes

### Calculated

- Markup Revenue: 10% if Markup Included

### Duplicated

- Credits: Payments received (Usually rent, copied from Credits sheet)
    - Debit/Credit Explanation value recorded as the Credits category value

## Income Statements

- 1 Workbook per Property, 1 Worksheet per Month

### Transactions

- Transactions copied from transactions spreadsheet into Income Statement for each property each month

### Additional Calculations

- Total Credits: Sum of all credits in worksheet
- Total Fees: Sum of all fees in worksheet
- Total Markup Revenue: Sum of all markups in worksheet (each entry calculated as 10% if markup included)
- Total MAF: ($5  count of units in property) + (0.06  Sum of credits)
- Total Security Deposits: Sum of all security deposits in worksheet
- Total Debits: (Sum of all debits in worksheet) + `Total Markup Revenue` + `Total MAF`
- Unlabeled field: `Total Credits` + `Total Security Deposits`
- Due to Owners: `Unlabeled field` - `Total Debits`
- Total to PM: `Total Markup Revenue` + `Total MAF`

### Legend/Details

- All security deposits are subject to a $12 Mailing Fee that includes tracking
- Late Fees = PM keeps all late fees
- All vendor charges subject to 10% markup (include with the expense, not a separate markup expense)
- CAM (logged like a unit) = Common Area Maintenance (for services like landscaping, cleaning common areas, snow removal, etc.)
- CAM Mark up 10%
- MAF = Management & Admin Fee (6% of collected rent plus $5/unit)
- MU = Multiple Units that received maintenance for the same thing
- Leasing Fee is 50% of base rent for one month for new leases (not pro-rated first month's rent)
- Renewal Fee is $125 flat fee

### Park and State Property Modifications

- MAF = Management & Admin Fee (8% of collected rent plus $5/unit)
- Renewal Fees is 25% fee
- New Lease fee is 100% fee
- Total MAF: ($5  count of units in property) + (0.08  Sum of credits)
- Airbnb Income: (number field)
- Airbnb-related unlabeled field
- Airbnb Total: `Airbnb Income` + `Airbnb-related unlabeled field`
- Airbnb PM Collection Fee: `Airbnb Total` * 0.04
- Unlabeled field (overridden): `Total Credits` + `Total Security Deposits` + `Airbnb Total`

## Summary Page

Calculated Data for Month

- Building: Property Name
- Due to Owners: `Due to Owners`
- Total to PM: `Total to PM`
- Total Fees: `Total Fees`
- New Lease Fees: Sum of `Debit` when `Debit/Credit Explanation` is "New Lease Fee"
- Renewal Fees: Sum of `Debit` when `Debit/Credit Explanation` is "Renewal Fee"

## MAF

Wilson, Park and State, and 1476 s high are 8%

schiller and patterson are 7%

Franklin Park, Bryden, Miller are 6%

Ann is 2%

Adams no MAF, no Markup, $15 admin

Park and State calculates airbnb income at 4% since the day to day of that is sub contracted out.
