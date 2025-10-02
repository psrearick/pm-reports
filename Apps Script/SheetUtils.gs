function getHeaderIndexMap_(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const key = normalizeKey_(headerRow[i]);
    if (key) {
      map[key] = i;
    }
  }
  return map;
}

function readSheetAsObjects_(sheet, headerRowIndex) {
  headerRowIndex = headerRowIndex || 1;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < headerRowIndex) {
    return [];
  }
  const range = sheet.getRange(headerRowIndex, 1, lastRow - headerRowIndex + 1, lastColumn);
  const values = range.getValues();
  if (!values.length) {
    return [];
  }
  const headers = values[0];
  const headerMap = getHeaderIndexMap_(headers);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const rowValues = values[i];
    if (rowValues.every(isBlank_)) {
      continue;
    }
    const obj = {};
    Object.keys(headerMap).forEach(function (key) {
      obj[key] = rowValues[headerMap[key]];
    });
    obj._rowNumber = headerRowIndex + i;
    rows.push(obj);
  }
  return rows;
}

function writeObjectsToSheet_(sheet, headers, rows, startRow) {
  startRow = startRow || 1;
  if (!headers || !headers.length) {
    throw new Error('Headers are required when writing objects to sheet: ' + sheet.getName());
  }
  const output = [headers];
  rows.forEach(function (row) {
    const line = [];
    headers.forEach(function (key) {
      line.push(row[key] !== undefined ? row[key] : '');
    });
    output.push(line);
  });
  const range = sheet.getRange(startRow, 1, output.length, headers.length);
  range.clearContent();
  range.setValues(output);
}
