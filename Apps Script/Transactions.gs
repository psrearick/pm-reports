function loadStagingData() {
  const controls = getStagingControls_();
  const propertyFilter = controls.property ? [controls.property] : null;
  const masterRows = fetchMasterTransactions_(propertyFilter, controls.startDate, controls.endDate, true);
  const sortedRows = masterRows.slice().sort(function (a, b) {
    return compareDatesAscending_(a.data['Date'], b.data['Date']);
  });
  const rowsToDisplay = controls.showDeleted ? sortedRows : sortedRows.filter(function (record) {
    return !toBool(record.data['Deleted']);
  });
  writeRowsToStaging_(rowsToDisplay.map(function (record) { return record.data; }));
}

function saveStagingData() {
  const controls = getStagingControls_();
  const stagingRows = readStagingRows_();
  const propertyFilter = controls.property ? [controls.property] : null;
  const masterRecords = fetchMasterTransactions_(propertyFilter, controls.startDate, controls.endDate, true);
  const masterById = {};
  masterRecords.forEach(function (record) {
    if (record.data['Transaction ID']) {
      masterById[record.data['Transaction ID']] = record;
    }
  });
  const processedIds = new Set();
  const permanentDeletes = [];
  const updates = [];
  const newTransactions = [];
  const now = utcNow_();

  stagingRows.forEach(function (row) {
    const id = row['Transaction ID'];
    const deletePermanently = toBool(row['Delete Permanently']);
    if (!id) {
      if (deletePermanently || isRowEffectivelyEmpty_(row)) {
        return;
      }
      const transaction = normalizeStagingRow_(row, controls.property);
      transaction['Transaction ID'] = Utilities.getUuid();
      if (transaction['Deleted']) {
        transaction['Deleted Timestamp'] = now;
      }
      applyMarkupForTransaction_(transaction);
      newTransactions.push(transaction);
      return;
    }
    const masterRecord = masterById[id];
    if (!masterRecord) {
      if (deletePermanently || isRowEffectivelyEmpty_(row)) {
        return;
      }
      const transaction = normalizeStagingRow_(row, controls.property);
      transaction['Transaction ID'] = id;
      if (transaction['Deleted']) {
        transaction['Deleted Timestamp'] = now;
      }
      applyMarkupForTransaction_(transaction);
      newTransactions.push(transaction);
      return;
    }
    processedIds.add(id);
    if (deletePermanently) {
      permanentDeletes.push(masterRecord);
      return;
    }
    const transaction = normalizeStagingRow_(row, controls.property);
    transaction['Transaction ID'] = id;
    const wasDeleted = toBool(masterRecord.data['Deleted']);
    if (transaction['Deleted']) {
      transaction['Deleted Timestamp'] = wasDeleted && masterRecord.data['Deleted Timestamp'] ? masterRecord.data['Deleted Timestamp'] : now;
    } else {
      transaction['Deleted Timestamp'] = '';
    }
    applyMarkupForTransaction_(transaction);
    updates.push({ rowIndex: masterRecord.rowIndex, transaction: transaction });
  });

  masterRecords.forEach(function (record) {
    const id = record.data['Transaction ID'];
    if (!id || processedIds.has(id)) {
      return;
    }
    if (permanentDeletes.some(function (item) { return item.data['Transaction ID'] === id; })) {
      return;
    }
    if (toBool(record.data['Deleted'])) {
      return;
    }
    const transaction = Object.assign({}, record.data);
    transaction['Deleted'] = true;
    transaction['Deleted Timestamp'] = now;
    updates.push({ rowIndex: record.rowIndex, transaction: transaction });
  });

  applyMasterUpdates_(updates);
  appendTransactions_(getSheetByName_(SHEET_NAMES.TRANSACTIONS), newTransactions);
  applyPermanentDeletes_(permanentDeletes);
}

function getStagingControls_() {
  const sheet = getSheetByName_(STAGING_CONTROL.SHEET);
  const property = normalizeStringValue_(sheet.getRange(STAGING_CONTROL.PROPERTY_CELL).getValue());
  const startDate = parseOptionalDate_(sheet.getRange(STAGING_CONTROL.START_DATE_CELL).getValue());
  const endDate = parseOptionalDate_(sheet.getRange(STAGING_CONTROL.END_DATE_CELL).getValue());
  const showDeleted = toBool(sheet.getRange(STAGING_CONTROL.SHOW_DELETED_CELL).getValue());
  return {
    property: property,
    startDate: startDate,
    endDate: endDate,
    showDeleted: showDeleted
  };
}

function parseOptionalDate_(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function fetchMasterTransactions_(propertyFilter, startDate, endDate, includeDeleted) {
  const sheet = getSheetByName_(SHEET_NAMES.TRANSACTIONS);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2) {
    return [];
  }
  const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = range.getValues();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headerIndex = getHeaderIndexMap_(headers);
  const properties = propertyFilter ? ensureArray_(propertyFilter).filter(function (item) { return !!normalizeStringValue_(item); }) : [];
  const matchAllProperties = properties.length === 0;
  const results = [];
  for (let i = 0; i < values.length; i++) {
    const rowValues = values[i];
    if (rowValues.every(isBlank_)) {
      continue;
    }
    const data = {};
    TRANSACTION_HEADERS.forEach(function (header) {
      const index = headerIndex[header];
      data[header] = index !== undefined ? rowValues[index] : '';
    });
    if (!data['Property']) {
      continue;
    }
    if (!matchAllProperties && properties.indexOf(data['Property']) === -1) {
      continue;
    }
    if (!withinDateRange_(data['Date'], startDate, endDate)) {
      continue;
    }
    if (!includeDeleted && toBool(data['Deleted'])) {
      continue;
    }
    results.push({ rowIndex: i + 2, data: data });
  }
  return results;
}

function withinDateRange_(value, startDate, endDate) {
  if (!startDate && !endDate) {
    return true;
  }
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) {
    return false;
  }
  if (startDate && date < startDate) {
    return false;
  }
  if (endDate && date > endDate) {
    return false;
  }
  return true;
}

function writeRowsToStaging_(rows) {
  const sheet = getSheetByName_(SHEET_NAMES.STAGING);
  sheet.getRange(1, 1, 1, STAGING_HEADERS.length).setValues([STAGING_HEADERS]);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, STAGING_HEADERS.length).clearContent();
  }
  if (!rows.length) {
    return;
  }
  const data = rows.map(function (row) {
    const output = [];
    STAGING_HEADERS.forEach(function (header) {
      if (header === 'Delete Permanently') {
        output.push(false);
      } else if (header === 'Deleted' || header === 'Markup Included') {
        output.push(toBool(row[header]));
      } else {
        output.push(row[header] !== undefined ? row[header] : '');
      }
    });
    return output;
  });
  sheet.getRange(2, 1, data.length, STAGING_HEADERS.length).setValues(data);
}

function readStagingRows_() {
  const sheet = getSheetByName_(SHEET_NAMES.STAGING);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const headers = sheet.getRange(1, 1, 1, STAGING_HEADERS.length).getValues()[0];
  const headerIndex = getHeaderIndexMap_(headers);
  const range = sheet.getRange(2, 1, lastRow - 1, STAGING_HEADERS.length);
  const values = range.getValues();
  const rows = [];
  values.forEach(function (valueRow) {
    if (valueRow.every(isBlank_)) {
      return;
    }
    const row = {};
    STAGING_HEADERS.forEach(function (header) {
      const index = headerIndex[header];
      row[header] = index !== undefined ? valueRow[index] : '';
    });
    rows.push(row);
  });
  return rows;
}

function normalizeStagingRow_(row, fallbackProperty) {
  const transaction = {};
  TRANSACTION_HEADERS.forEach(function (header) {
    transaction[header] = row[header] !== undefined ? row[header] : '';
  });
  if (!transaction['Property'] && fallbackProperty) {
    transaction['Property'] = fallbackProperty;
  }
  transaction['Property'] = normalizeStringValue_(transaction['Property']);
  if (!transaction['Property']) {
    throw new Error('Each row must specify a Property.');
  }
  transaction['Date'] = normalizeDateValue_(transaction['Date']);
  transaction['Credits'] = isBlank_(row['Credits']) ? '' : parseCurrency_(row['Credits']);
  transaction['Fees'] = isBlank_(row['Fees']) ? '' : parseCurrency_(row['Fees']);
  transaction['Debits'] = isBlank_(row['Debits']) ? '' : parseCurrency_(row['Debits']);
  transaction['Security Deposits'] = isBlank_(row['Security Deposits']) ? '' : parseCurrency_(row['Security Deposits']);
  transaction['Markup Included'] = toBool(transaction['Markup Included']);
  transaction['Deleted'] = toBool(transaction['Deleted']);
  transaction['Markup Revenue'] = isBlank_(row['Markup Revenue']) ? '' : parseCurrency_(row['Markup Revenue']);
  transaction['Deleted Timestamp'] = normalizeDateValue_(transaction['Deleted Timestamp']);
  return transaction;
}

function compareDatesAscending_(a, b) {
  const dateA = normalizeDateForSort_(a);
  const dateB = normalizeDateForSort_(b);
  return dateA - dateB;
}

function normalizeDateForSort_(value) {
  if (value instanceof Date) {
    return value.getTime();
  }
  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return parsed.getTime();
  }
  return 0;
}

function normalizeDateValue_(value) {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' : parsed;
}

function applyMarkupForTransaction_(transaction) {
  if (!transaction['Markup Included']) {
    transaction['Markup Revenue'] = 0;
    return;
  }
  const property = getPropertyByName(transaction['Property']);
  if (!property) {
    transaction['Markup Revenue'] = transaction['Markup Revenue'] || 0;
    return;
  }
  const basis = transaction['Debits'] || 0;
  const markup = basis * (property.markup || 0);
  transaction['Markup Revenue'] = Math.round(markup * 100) / 100;
}

function applyMasterUpdates_(updates) {
  if (!updates.length) {
    return;
  }
  const sheet = getSheetByName_(SHEET_NAMES.TRANSACTIONS);
  const headers = TRANSACTION_HEADERS;
  updates.sort(function (a, b) { return a.rowIndex - b.rowIndex; });
  updates.forEach(function (update) {
    const rowValues = headers.map(function (header) {
      return update.transaction[header] !== undefined ? update.transaction[header] : '';
    });
    sheet.getRange(update.rowIndex, 1, 1, headers.length).setValues([rowValues]);
  });
}

function applyPermanentDeletes_(permanentDeletes) {
  if (!permanentDeletes.length) {
    return;
  }
  const sheet = getSheetByName_(SHEET_NAMES.TRANSACTIONS);
  const rows = permanentDeletes.map(function (record) { return record.rowIndex; }).sort(function (a, b) { return b - a; });
  rows.forEach(function (rowIndex) {
    sheet.deleteRow(rowIndex);
  });
}

function isRowEffectivelyEmpty_(row) {
  const fields = ['Date', 'Credits', 'Fees', 'Debits', 'Security Deposits', 'Debit/Credit Explanation', 'Internal Notes'];
  return fields.every(function (field) {
    return !row[field];
  });
}
