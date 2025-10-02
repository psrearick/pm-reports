function ensureLogSheet_(sheetName, headers) {
  let sheet = getActiveSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    sheet = getActiveSpreadsheet_().insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function appendLogRow_(sheet, values) {
  sheet.appendRow(values);
}

function readLogRecords_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const rows = range.getValues();
  const records = [];
  rows.forEach(function (row) {
    if (row.every(isBlank_)) {
      return;
    }
    const record = {};
    headers.forEach(function (header, index) {
      record[header] = row[index];
    });
    records.push(record);
  });
  return records;
}
