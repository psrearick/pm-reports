function importCredits() {
  const creditsFile = getAttributeValue("Credits Document");
  const creditsTempFile = convertXLSXToGoogleSheet(creditsFile);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getDataFromSheet(creditsTempFile);
  const filteredData = filterCreditsData(data);

  if (getAttributeValue("Add Credits Sheet")) {
  const creditsSheet = ss.insertSheet(creditsTempFile);
  creditsSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  addCreditsToTransactionSheet(data);

  Drive.Files.remove(creditsTempFile);
}

function addCreditsToTransactionSheet(creditsData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Transactions");
}

function filterCreditsData(data) {
  const outputData = [];

  const fields = {
    "Date": { "label": "", "index": 0 },
    "Amount":  { "label": "", "index": 0 },
    "Property":  { "label": "", "index": 0 },
    "Unit":  { "label": "", "index": 0 },
    "Category":  { "label": "", "index": 0 },
    "Subcategory":  { "label": "", "index": 0 },
  }

  const labelsRow = [];
  for (const key in fields) {
    fields[key]["label"] = getAttributeValue("Credits " + key);
    labelsRow.push(key);
  }
  outputData.push(labelsRow);

  let allFieldsDefined = false;


  data.forEach((row) => {
    let includesAllFields = true;

    for (const key in fields) {
      const fieldIndex = row.indexOf(fields[key]["label"]);
      if (fieldIndex == -1) {
        includesAllFields = false;
        break;
      }

      fields[key]["index"] = fieldIndex;
    }

    if (includesAllFields) {
      allFieldsDefined = true;
      return;
    }

    if (!allFieldsDefined) {
      return;
    }

    const newRow = []

    if (!row[fields["Date"]["index"]] || row[fields["Date"]["index"]].startsWith("Total")) {
      return;
    }

    if (!row[fields["Amount"]["index"]]) {
      return;
    }

    if (!row[fields["Property"]["index"]]) {
      return;
    }

    if (!row[fields["Unit"]["index"]] || row[fields["Unit"]["index"]] == "–") {
      return;
    }

    if (!row[fields["Category"]["index"]]) {
      return;
    }

    for (const field in fields) {
      newRow.push(row[fields[field]["index"]]);
    }

    outputData.push(newRow);
  });

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

// HELPERS

function getAttributeValue(attributeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Configuration");
  const attributes = ws.getRange("a2:B").getValues();
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
