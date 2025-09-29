// --- CONSTANTS ---
const SS = SpreadsheetApp.getActiveSpreadsheet();
const TRANSACTIONS_SHEET = SS.getSheetByName("Transactions");
const PROPERTIES_SHEET = SS.getSheetByName("Properties");
const CONFIG_SHEET = SS.getSheetByName("Configuration");

// --- MENU ---
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Automation')
    .addItem('1. Import Credits', 'importCredits')
    .addSeparator()
    .addItem('2. Generate Monthly Reports', 'generateMonthlyReports')
    .addToUi();
}

// --- MAIN REPORT GENERATION FUNCTION ---
function generateMonthlyReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Generate Reports', 'Enter the month and year for the reports (e.g., "Sep 2025"):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText() === '') {
    return; // User cancelled
  }

  const reportMonthStr = response.getResponseText();
  const reportDate = new Date(reportMonthStr);
  if (isNaN(reportDate.getTime())) {
    ui.alert('Invalid Date', 'Please enter a valid month and year, like "September 2025".', ui.ButtonSet.OK);
    return;
  }

  const startDate = new Date(reportDate.getFullYear(), reportDate.getMonth() - 1, 20);
  const endDate = new Date(reportDate.getFullYear(), reportDate.getMonth(), 20);

  const properties = getProperties();
  const transactions = getTransactionsData(startDate, endDate);
  const generatedSheetNames = [];

  const allUnitCounts = countUnitsPerProperty(TRANSACTIONS_SHEET.getDataRange().getValues());

  const reportSpreadsheetName = `Monthly Reports - ${Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "MMM yyyy")}`;
  const newSS = SpreadsheetApp.create(reportSpreadsheetName);
  const parentFolder = DriveApp.getFileById(SS.getId()).getParents().next();
  DriveApp.getFileById(newSS.getId()).moveTo(parentFolder);

  SS.getSheetByName("ReportBodyTemplate").copyTo(newSS).setName("ReportBodyTemplate");
  SS.getSheetByName("ReportTotalsTemplate").copyTo(newSS).setName("ReportTotalsTemplate");
  SS.getSheetByName("ReportAirbnbTemplate").copyTo(newSS).setName("ReportAirbnbTemplate");

  newSS.deleteSheet(newSS.getSheetByName('Sheet1'));
  SpreadsheetApp.setActiveSpreadsheet(newSS);

  for (const property of properties) {
    const propertyTransactions = transactions.filter(t => t.Property === property.Property);
    const unitCount = allUnitCounts[property.Property] || 0;

    const sheetName = createPropertyReport(newSS, property, propertyTransactions, unitCount);
    generatedSheetNames.push({ name: property.Property, sheet: sheetName });
  }

  if (generatedSheetNames.length > 0) {
    createSummaryReport(newSS, generatedSheetNames);
  }

  ui.alert('Success', `Reports generated in new spreadsheet: "${reportSpreadsheetName}"`, ui.ButtonSet.OK);
}


// --- HELPER FUNCTIONS FOR REPORTING ---

/**
 * Creates a single property report sheet by using template placeholders.
 */
function createPropertyReport(spreadsheet, property, transactions, unitCount) {
  const sheetName = property.Property;
  const reportSheet = spreadsheet.insertSheet(sheetName);

  // Get templates from the new spreadsheet
  const bodyTemplate = spreadsheet.getSheetByName("ReportBodyTemplate");
  const totalsTemplate = spreadsheet.getSheetByName("ReportTotalsTemplate");
  const airbnbTemplate = spreadsheet.getSheetByName("ReportAirbnbTemplate");

  let currentRow = 1;

  // 1. Copy the full structure (body, totals, etc.) into the new sheet first
  bodyTemplate.getDataRange().copyTo(reportSheet.getRange(currentRow, 1));
  currentRow += bodyTemplate.getLastRow();

  const totalsStartRow = currentRow + 1;
  totalsTemplate.getDataRange().copyTo(reportSheet.getRange(totalsStartRow, 1));
  currentRow += totalsTemplate.getLastRow() + 1; // Add one for spacing

  let airbnbStartRow = 0;
  if (property['Has Airbnb']) {
    airbnbStartRow = currentRow + 1;
    airbnbTemplate.getDataRange().copyTo(reportSheet.getRange(airbnbStartRow, 1));
  }

  // 2. Find the placeholder row for transaction data
  const transactionFinder = reportSheet.createTextFinder("\\[UNIT NUMBER\\]");
  const transactionPlaceholderCell = transactionFinder.findNext();

  if (!transactionPlaceholderCell) {
    // If no placeholder is found, skip this property.
    Logger.log(`Could not find transaction placeholder '[UNIT NUMBER]' for property: ${property.Property}`);
    return sheetName;
  }

  const transactionStartRow = transactionPlaceholderCell.getRow();
  const transactionTemplateRow = reportSheet.getRange(transactionStartRow, 1, 1, bodyTemplate.getLastColumn());

  // 3. Populate Transactions using the template row's format
  if (transactions.length > 0) {
    const transactionRows = transactions.map(t => {
       // Unit,Credits,Fees,Debits,Security Deposits,Date,Debit/Credit Explanation,Markup Included,Markup Revenue,Internal Notes
      return [
        t.Unit, t.Credits, t.Fees, t.Debits, t.SecurityDeposits, t.Date,
        t['Debit/Credit Explanation'], t['Markup Included'], t.MarkupRevenue, t['Internal Notes']
      ];
    });

    // If there is more than one transaction, insert new rows for them.
    if (transactions.length > 1) {
      reportSheet.insertRowsAfter(transactionStartRow, transactions.length - 1);
    }

    // Copy the formatting from the template row to all the new rows.
    const formatRange = reportSheet.getRange(transactionStartRow, 1, transactions.length, transactionTemplateRow.getNumColumns());
    transactionTemplateRow.copyTo(formatRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    // Now, set the values in the newly formatted range.
    formatRange.setValues(transactionRows);

  } else {
    // If there are no transactions, clear the placeholder row.
    transactionTemplateRow.clearContent();
  }

  const transactionDataEndRow = transactionStartRow + transactions.length -1;

  // 4. Add Formulas for Calculations
  addReportFormulas(reportSheet, property, transactionStartRow, transactionDataEndRow, totalsStartRow, airbnbStartRow, unitCount);

  // Auto-resize columns for readability
  reportSheet.getRange("A:J").autoResizeColumns();

  return sheetName;
}


/**
 * Adds all calculation formulas to a property report sheet.
 * (This is updated to handle dynamic transaction row ranges)
 */
function addReportFormulas(sheet, property, transStartRow, transEndRow, totalsStartRow, airbnbStartRow, unitCount) {
    // Define ranges for easy reference. Assumes standard template layout.
    // Handles the case where there are no transactions.
    const hasTransactions = transStartRow <= transEndRow;
    const creditsRange = hasTransactions ? `B${transStartRow}:B${transEndRow}` : "0";
    const feesRange = hasTransactions ? `C${transStartRow}:C${transEndRow}` : "0";
    const debitsRange = hasTransactions ? `D${transStartRow}:D${transEndRow}` : "0";
    const securityDepositsRange = hasTransactions ? `E${transStartRow}:E${transEndRow}` : "0";
    const markupRevenueRange = hasTransactions ? `I${transStartRow}:I${transEndRow}` : "0";

    // --- Totals Section Formulas ---
    // These offsets depend on your ReportTotalsTemplate layout.
    // E.g., if "Credits" is 2 rows below "Totals" and in column B, it's totalsStartRow + 2.
    const totalCreditsCell = `B${totalsStartRow + 2}`;
    const totalFeesCell = `B${totalsStartRow + 3}`;
    const totalDebitsCell = `C${totalsStartRow + 2}`;
    const totalMarkupRevenueCell = `C${totalsStartRow + 3}`;
    const totalMafCell = `C${totalsStartRow + 4}`;
    const toOwnersCell = `B${totalsStartRow + 7}`;
    const toStantCell = `B${totalsStartRow + 8}`;

    // Sums
    sheet.getRange(totalCreditsCell).setFormula(`=SUM(${creditsRange})`);
    sheet.getRange(totalFeesCell).setFormula(`=SUM(${feesRange})`);
    sheet.getRange(totalMarkupRevenueCell).setFormula(`=SUM(${markupRevenueRange})`);

    // MAF Calculation (handles property-specific rates)
    let mafFormula = `=(5 * ${unitCount}) + (${property.MAF} * ${totalCreditsCell})`;
    if (property.Property === "2536 Adams Ave") {
        mafFormula = '=15'; // Flat admin fee
    }
    sheet.getRange(totalMafCell).setFormula(mafFormula);

    // Total Debits Calculation
    sheet.getRange(totalDebitsCell).setFormula(`=SUM(${debitsRange}) + ${totalMarkupRevenueCell} + ${totalMafCell}`);

    // Final Payout Calculations
    let preOwnerTotalFormula = `=${totalCreditsCell} + SUM(${securityDepositsRange})`;
    if (property['Has Airbnb']) {
      const airbnbTotalCell = `B${airbnbStartRow + 2}`; // Assumes income is in B, 2 rows below "Airbnb"
      // Note: Airbnb income is likely manual entry, so we reference the empty cell
      preOwnerTotalFormula += `+${airbnbTotalCell}`;

      const airbnbFeeCell = `C${airbnbStartRow + 2}`; // Assumes fee is in C
      sheet.getRange(airbnbFeeCell).setFormula(`=${airbnbTotalCell} * ${property.Airbnb}`);
      sheet.getRange(toStantCell).setFormula(`=${totalMarkupRevenueCell} + ${totalMafCell} + ${totalFeesCell} + ${airbnbFeeCell}`);
    } else {
       sheet.getRange(toStantCell).setFormula(`=${totalMarkupRevenueCell} + ${totalMafCell} + ${totalFeesCell}`);
    }

    sheet.getRange(toOwnersCell).setFormula(`=(${preOwnerTotalFormula}) - ${totalDebitsCell}`);
}

/**
 * Creates the summary report page.
 */
function createSummaryReport(spreadsheet, generatedSheets) {
  const summarySheet = spreadsheet.insertSheet("Summary Page", 0);
  const headers = ["Building", "Due to Owners", "Total to PM", "Total Fees", "New Lease Fees", "Renewal Fees"];
  summarySheet.getRange("A1:F1").setValues([headers]).setFontWeight("bold");

  const rows = [];
  for (const item of generatedSheets) {
    const propertySheetName = item.sheet;
    // These cell references must match the layout of your Totals template
    const dueToOwnersFormula = `='${propertySheetName}'!B10`; // Assumes Due to Owners is in B10
    const totalToPMFormula = `='${propertySheetName}'!B11`;     // Assumes To Stant is in B11
    const totalFeesFormula = `='${propertySheetName}'!B6`;      // Assumes Fees is in B6
    const newLeaseFeesFormula = `=SUMIF('${propertySheetName}'!G:G, "New Lease Fee", '${propertySheetName}'!D:D)`;
    const renewalFeesFormula = `=SUMIF('${propertySheetName}'!G:G, "Renewal Fee", '${propertySheetName}'!D:D)`;

    rows.push([item.name, dueToOwnersFormula, totalToPMFormula, totalFeesFormula, newLeaseFeesFormula, renewalFeesFormula]);
  }

  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, headers.length).setFormulas(rows);
  }
  summarySheet.autoResizeColumns(1, headers.length);
}

/**
 * Gets and processes transaction data for a given date range.
 */
function getTransactionsData(startDate, endDate) {
  const allData = TRANSACTIONS_SHEET.getDataRange().getValues();
  const headers = allData.shift();
  const headerMap = {};
  headers.forEach((header, i) => headerMap[header.trim()] = i);

  const properties = getProperties();
  const propertyMap = properties.reduce((map, obj) => {
    map[obj.Property] = obj;
    return map;
  }, {});

  const transactions = allData.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h.trim()] = row[i]);
    return obj;
  }).filter(row => {
    const rowDate = row.Date;
    if (!row.Property || !(rowDate instanceof Date)) {
      return false;
    }
    const rowDateTime = rowDate.getTime();
    return rowDateTime >= startDate.getTime() && rowDateTime < endDate.getTime();
  });

  // Pre-calculate Markup Revenue
  transactions.forEach(t => {
      const property = propertyMap[t.Property];
      if (property && t['Markup Included'] === true && t.Debits > 0) {
          t.MarkupRevenue = t.Debits * property.Markup;
      } else {
          t.MarkupRevenue = 0;
      }
  });

  return transactions;
}

/**
 * Counts unique units for each property.
 */
function countUnitsPerProperty(data) {
    const headers = data.shift();
    const propertyIndex = headers.indexOf("Property");
    const unitIndex = headers.indexOf("Unit");
    const counts = {};
    const uniqueUnits = {};

    data.forEach(row => {
        const property = row[propertyIndex];
        const unit = row[unitIndex];
        if (property && unit) {
            if (!uniqueUnits[property]) {
                uniqueUnits[property] = new Set();
            }
            uniqueUnits[property].add(unit);
        }
    });

    for (const property in uniqueUnits) {
        counts[property] = uniqueUnits[property].size;
    }
    return counts;
}


// --- ORIGINAL SCRIPT FUNCTIONS (UNCHANGED OR SLIGHTLY MODIFIED) ---

function importCredits() {
  const creditsFile = getAttributeValue("Credits Document");
  const creditsTempFile = convertXLSXToGoogleSheet(creditsFile);
  var data = getDataFromSheet(creditsTempFile);
  const filteredData = filterCreditsData(data);

  if (getAttributeValue("Add Credits Sheet")) {
    const creditsSheet = SS.insertSheet(creditsTempFile);
    creditsSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  addCreditsToTransactionSheet(filteredData);

  Drive.Files.remove(creditsTempFile);
}

function addCreditsToTransactionSheet(creditsData) {
  if (!creditsData || creditsData.length <= 1) return;
  if (!TRANSACTIONS_SHEET) return;

  const lastRow = TRANSACTIONS_SHEET.getLastRow();
  const transactionHeaders = TRANSACTIONS_SHEET.getRange(1, 1, 1, TRANSACTIONS_SHEET.getLastColumn()).getValues()[0];

  const headerMap = {};
  transactionHeaders.forEach((header, i) => headerMap[header.trim()] = i);

  const requiredCols = ["Date", "Property", "Unit", "Debit/Credit Explanation", "Security Deposits", "Fees", "Credits"];
  for (const col of requiredCols) {
    if (headerMap[col] === undefined) return;
  }

  const rowsToAdd = [];
  const creditsDataOnly = creditsData.slice(1);

  creditsDataOnly.forEach(creditRow => {
    const date = creditRow[0]; // Is now a Date object
    const amount = parseCurrency(creditRow[1]);
    let property = creditRow[2]
    const propertyData = lookupProperty(property);
    if (propertyData) {
      property = propertyData.Property;
    }
    const unit = creditRow[3];
    const category = (creditRow[4] || "").toLowerCase().trim();
    const subcategory = creditRow[5] || "";

    if (category.includes("property general expense") || category.includes("owner distribution")) {
      return;
    }

    let newTransactionRow = new Array(transactionHeaders.length).fill('');
    newTransactionRow[headerMap["Date"]] = date;
    newTransactionRow[headerMap["Property"]] = property;
    newTransactionRow[headerMap["Unit"]] = unit;

    let explanation = creditRow[4] || "";
    if (subcategory && subcategory !== "–") {
      explanation += ` - ${subcategory}`;
    }
    newTransactionRow[headerMap["Debit/Credit Explanation"]] = explanation.trim();

    if (category.includes('deposit')) {
      newTransactionRow[headerMap["Security Deposits"]] = amount;
    } else if (category.includes('tenant charges') || category.includes('fees') || category.includes('late payment fee')) {
      newTransactionRow[headerMap["Fees"]] = amount;
    } else {
      newTransactionRow[headerMap["Credits"]] = amount;
    }
    rowsToAdd.push(newTransactionRow);
  });

  if (rowsToAdd.length > 0) {
    const startRow = lastRow + 1;
    const numRows = rowsToAdd.length;
    TRANSACTIONS_SHEET.getRange(startRow, 1, numRows, rowsToAdd[0].length).setValues(rowsToAdd);

    const dateColumnIndex = headerMap["Date"] + 1;
    if (dateColumnIndex > 0) {
      const dateRange = TRANSACTIONS_SHEET.getRange(startRow, dateColumnIndex, numRows, 1);
      dateRange.setNumberFormat("M/dd/yyyy");
    }
  }
}

function filterCreditsData(data) {
  const outputData = [];
  const fields = {
    "Date": { label: getAttributeValue("Credits Date"), index: -1 },
    "Amount": { label: getAttributeValue("Credits Amount"), index: -1 },
    "Property": { label: getAttributeValue("Credits Property"), index: -1 },
    "Unit": { label: getAttributeValue("Credits Unit"), index: -1 },
    "Category": { label: getAttributeValue("Credits Category"), index: -1 },
    "Subcategory": { label: getAttributeValue("Credits Subcategory"), index: -1 },
  };

  outputData.push(Object.keys(fields));

  let headerRowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    let foundCount = 0;
    for (const key in fields) {
      const fieldIndex = row.indexOf(fields[key].label);
      if (fieldIndex !== -1) {
        fields[key].index = fieldIndex;
        foundCount++;
      }
    }
    if (foundCount >= 4) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) return [];

  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    const dateField = row[fields.Date.index];
    if (!dateField || String(dateField).trim().startsWith("Total")) continue;
    if (!row[fields.Amount.index] || !row[fields.Property.index]) continue;

    const timeZone = Session.getScriptTimeZone();
    const format = "MMM dd, yyyy";
    const dateVal = Utilities.parseDate(String(dateField).trim(), timeZone, format);
    const unitVal = row[fields.Unit.index];

    const newRow = [
      dateVal,
      row[fields.Amount.index],
      row[fields.Property.index],
      (unitVal === "–") ? "" : unitVal,
      row[fields.Category.index] || "",
      row[fields.Subcategory.index] || ""
    ];
    outputData.push(newRow);
  }
  return outputData;
}

function getProperties() {
  const data = PROPERTIES_SHEET.getDataRange().getValues();
  const keys = data.shift();
  return data.map((propertyValues) => {
    const property = {};
    for (let i = 0; i < keys.length; i++) {
      property[keys[i].trim()] = propertyValues[i];
    }
    return property;
  });
}

function getAttributeValue(attributeName) {
  const attributes = CONFIG_SHEET.getRange("A2:B").getValues();
  const matches = attributes.filter(x => x[0] == attributeName);
  return matches.length > 0 ? matches[0][1] : "";
}

function convertXLSXToGoogleSheet(fileId) {
  const xlsxFile = DriveApp.getFileById(fileId);
  const fileName = xlsxFile.getName().replace(/\.xlsx?$/, "");
  const resource = {
    title: fileName,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: xlsxFile.getParents().next().getId() }]
  };
  const newSheetFile = Drive.Files.copy(resource, fileId);
  return newSheetFile.id;
}

function getDataFromSheet(sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId);
  return sheet.getSheets()[0].getDataRange().getValues();
}

function parseCurrency(value) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return 0;
  return parseFloat(value.replace(/[$,]/g, '')) || 0;
}

function lookupProperty(searchValue) {
  const properties = getProperties();
  for (const property of properties) {
    const searchValueLower = searchValue.toLowerCase();
    const pattern = property.Key.toLowerCase().split(",").join(".*");
    const regex = new RegExp(pattern, "i");
    if (regex.test(searchValueLower)) {
      return property;
    }
  }
  return null;
}
