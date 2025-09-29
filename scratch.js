const SS = SpreadsheetApp.getActiveSpreadsheet();
const TRANSACTIONS = SS.getSheetByName("Transactions");
const PROPERTIES = getProperties();

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
  if (!creditsData || creditsData.length <= 1) {
    return;
  }

  if (!TRANSACTIONS) {
    return;
  }

  const lastRow = TRANSACTIONS.getLastRow();
  const transactionHeaders = TRANSACTIONS.getRange(1, 1, 1, TRANSACTIONS.getLastColumn()).getValues()[0];

  const headerMap = {};
  transactionHeaders.forEach((header, i) => {
    headerMap[header.trim()] = i;
  });

  const requiredCols = ["Date", "Property", "Unit", "Debit/Credit Explanation", "Security Deposits", "Fees", "Credits"];
  for (const col of requiredCols) {
    if (headerMap[col] === undefined) {
      return;
    }
  }

  const rowsToAdd = [];
  const creditsDataOnly = creditsData.slice(1);

  creditsDataOnly.forEach(creditRow => {
    const date = creditRow[0];
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
    TRANSACTIONS.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
  }
}

function filterCreditsData(data) {
  const outputData = [];
  const fields = {
    "Date": { label: getAttributeValue("Credits Date"), index: -1 },
    "Amount":  { label: getAttributeValue("Credits Amount"), index: -1 },
    "Property":  { label: getAttributeValue("Credits Property"), index: -1 },
    "Unit":  { label: getAttributeValue("Credits Unit"), index: -1 },
    "Category":  { label: getAttributeValue("Credits Category"), index: -1 },
    "Subcategory":  { label: getAttributeValue("Credits Subcategory"), index: -1 },
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

  if (headerRowIndex === -1) {
    return [];
  }

  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];

    const dateField = row[fields.Date.index];
    if (!dateField || String(dateField).trim().startsWith("Total")) {
      continue;
    }

    if (!row[fields.Amount.index] || !row[fields.Property.index]) {
      continue;
    }

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

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Import Credits', 'importCredits');
  menu.addToUi();
}

function lookupProperty(searchValue) {
  for (const property of PROPERTIES) {
    searchValueLower = searchValue.toLowerCase();
    const pattern = property.Key.toLowerCase().split(",").join(".*");
    regex = new RegExp(pattern, "i")
    const is_match = regex.test(searchValueLower);

    if (is_match) {
      return property;
    }
  }

  return null;
}


function getProperties() {
  const ws = SS.getSheetByName("Properties");
  const data = ws.getDataRange().getValues();
  const keys = data.shift();

  return data.map((propertyValues) => {
    const property = {};

    for (let i = 0; i < keys.length; i++) {
      property[keys[i].replace(" ", "")] = propertyValues[i];
    }

    return property;
  });
}

function getAttributeValue(attributeName) {
  const ws = SS.getSheetByName("Configuration");
  const attributes = ws.getRange("A2:B").getValues();
  const matches = attributes.filter(x => x[0]==attributeName);

  if (matches.length == 0) {
    return "";
  }

  return matches[0][1];
}

function getDataFromSheet(sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId);
  const sourceSheet = sheet.getSheets()[0];

  return sourceSheet.getDataRange().getValues();
}

function parseCurrency(value) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return 0;
  return parseFloat(value.replace(/[$,]/g, '')) || 0;
}
