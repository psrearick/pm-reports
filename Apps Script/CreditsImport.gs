function importCredits() {
  const headerMap = getCreditsHeaderMapping();
  const sourceConfig = getCreditsSourceConfig();
  const filesToProcess = resolveCreditsFiles_(sourceConfig);
  if (!filesToProcess.length) {
    throw new Error('No credit files found to process.');
  }
  const importLogSheet = ensureLogSheet_(SHEET_NAMES.IMPORT_LOG, LOG_HEADERS.IMPORT);
  const processedSignatures = getProcessedCreditSignatures_(importLogSheet);
  const transactionsSheet = getSheetByName_(SHEET_NAMES.TRANSACTIONS);
  const results = [];
  filesToProcess.forEach(function (file) {
    const fileId = file.getId();
    const lastUpdated = file.getLastUpdated().getTime();
    if (isCreditFileAlreadyProcessed_(fileId, lastUpdated, processedSignatures)) {
      results.push({ file: file, skipped: true, reason: 'Already processed' });
      return;
    }
    const openResult = openSpreadsheetFromFile_(file);
    try {
      const records = extractCreditsRecords_(openResult.spreadsheet, headerMap);
      if (!records.length) {
        results.push({ file: file, skipped: true, reason: 'No data rows detected' });
        logCreditImport_(importLogSheet, file, lastUpdated, 0, 'No data rows detected');
        return;
      }
      const transactions = records.map(function (record) {
        return buildTransactionRowFromCredit_(record);
      });
      appendTransactions_(transactionsSheet, transactions);
      logCreditImport_(importLogSheet, file, lastUpdated, transactions.length, 'Imported');
      if (sourceConfig.addSheet) {
        addCreditsAuditSheet_(records, file.getName());
      }
      results.push({ file: file, skipped: false, imported: transactions.length });
    } finally {
      if (openResult.isTemporary) {
        deleteSpreadsheetById(openResult.spreadsheet.getId());
      }
    }
  });
  return results;
}

function resolveCreditsFiles_(sourceConfig) {
  const filesMap = {};
  if (sourceConfig.fileId) {
    const trimmed = sourceConfig.fileId.toString().trim();
    if (trimmed) {
      const file = DriveApp.getFileById(trimmed);
      filesMap[file.getId()] = file;
    }
  }
  if (sourceConfig.folderId) {
    const folder = DriveApp.getFolderById(sourceConfig.folderId.toString().trim());
    const iterator = folder.getFiles();
    while (iterator.hasNext()) {
      const file = iterator.next();
      if (isProcessableCreditsMimeType_(file.getMimeType())) {
        filesMap[file.getId()] = file;
      }
    }
  }
  return Object.keys(filesMap).map(function (id) { return filesMap[id]; });
}

function isProcessableCreditsMimeType_(mimeType) {
  if (!mimeType) {
    return false;
  }
  return mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.MICROSOFT_EXCEL;
}

function openSpreadsheetFromFile_(file) {
  const mimeType = file.getMimeType();
  if (mimeType === MimeType.GOOGLE_SHEETS) {
    return { spreadsheet: SpreadsheetApp.openById(file.getId()), isTemporary: false };
  }
  const spreadsheet = convertExcelFileToSheet(file);
  return { spreadsheet: spreadsheet, isTemporary: true };
}

function extractCreditsRecords_(spreadsheet, headerMap) {
  const sheet = spreadsheet.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return [];
  }
  const headerInfo = locateCreditsHeaderRow_(values, headerMap);
  const dataRows = [];
  for (let r = headerInfo.rowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const rawDate = row[headerInfo.columns.date];
    const rawAmount = row[headerInfo.columns.amount];
    if (isBlank_(rawDate) && isBlank_(rawAmount)) {
      continue;
    }
    if (typeof rawDate === 'string' && rawDate.indexOf('Total') === 0) {
      continue;
    }
    if (typeof rawAmount === 'string' && rawAmount.indexOf('Total') === 0) {
      continue;
    }
    if (rawDate === '' || rawDate === null) {
      continue;
    }
    const amount = parseCurrency_(rawAmount);
    if (!amount) {
      continue;
    }
    const record = {
      date: parseDateValue_(rawDate),
      amount: amount,
      property: normalizeStringValue_(row[headerInfo.columns.property]),
      unit: normalizeUnitValue_(row[headerInfo.columns.unit]),
      category: normalizeStringValue_(row[headerInfo.columns.category]),
      subcategory: normalizeStringValue_(row[headerInfo.columns.subcategory]),
      status: headerInfo.columns.status >= 0 ? normalizeStringValue_(row[headerInfo.columns.status]) : '',
      method: headerInfo.columns.method >= 0 ? normalizeStringValue_(row[headerInfo.columns.method]) : '',
      counterparty: headerInfo.columns.payer >= 0 ? normalizeStringValue_(row[headerInfo.columns.payer]) : ''
    };
    dataRows.push(record);
  }
  return dataRows;
}

function locateCreditsHeaderRow_(values, headerMap) {
  const keys = Object.keys(headerMap);
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const columns = {};
    let matches = 0;
    for (let c = 0; c < row.length; c++) {
      const cellValue = normalizeStringValue_(row[c]);
      keys.forEach(function (key) {
        if (columns[key] === undefined && cellValue === headerMap[key]) {
          columns[key] = c;
          matches++;
        }
      });
    }
    if (matches === keys.length) {
      // locate optional headers
      columns.status = findColumnIndex_(row, headerMap.status || 'Payment status');
      columns.method = findColumnIndex_(row, headerMap.method || 'Payment method');
      columns.payer = findColumnIndex_(row, headerMap.payer || 'Payer / Payee');
      return { rowIndex: r, columns: columns };
    }
  }
  throw new Error('Unable to locate credits header row. Check configuration.');
}

function findColumnIndex_(row, headerName) {
  if (!headerName) {
    return -1;
  }
  for (let c = 0; c < row.length; c++) {
    if (normalizeStringValue_(row[c]) === headerName) {
      return c;
    }
  }
  return -1;
}

function normalizeUnitValue_(value) {
  const normalized = normalizeStringValue_(value);
  if (normalized === '–' || normalized === '-') {
    return '';
  }
  return normalized;
}

function parseDateValue_(value) {
  if (value instanceof Date) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }
  throw new Error('Unable to parse date value: ' + value);
}

function buildTransactionRowFromCredit_(record) {
  const explanation = record.subcategory || record.category;
  const notesParts = [];
  if (record.method) {
    notesParts.push('Method: ' + record.method);
  }
  if (record.status) {
    notesParts.push('Status: ' + record.status);
  }
  if (record.counterparty) {
    notesParts.push('Payer: ' + record.counterparty);
  }
  const notes = notesParts.join(' | ');
  return {
    'Transaction ID': Utilities.getUuid(),
    'Date': record.date,
    'Property': record.property,
    'Unit': record.unit,
    'Credits': record.amount,
    'Fees': '',
    'Debits': '',
    'Security Deposits': '',
    'Debit/Credit Explanation': explanation,
    'Markup Included': false,
    'Markup Revenue': '',
    'Internal Notes': notes,
    'Deleted': false,
    'Deleted Timestamp': ''
  };
}

function appendTransactions_(sheet, transactions) {
  if (!transactions.length) {
    return;
  }
  const headers = TRANSACTION_HEADERS;
  const rows = transactions.map(function (transaction) {
    return headers.map(function (header) {
      return transaction[header] !== undefined ? transaction[header] : '';
    });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function getProcessedCreditSignatures_(importLogSheet) {
  const records = readLogRecords_(importLogSheet, LOG_HEADERS.IMPORT);
  const signatures = {};
  records.forEach(function (record) {
    const fileId = record[LOG_HEADERS.IMPORT[1]] || record['File ID'];
    const lastModified = record[LOG_HEADERS.IMPORT[3]] || record['Last Modified'];
    if (!fileId || !lastModified) {
      return;
    }
    const dateValue = lastModified instanceof Date ? lastModified : new Date(lastModified);
    if (isNaN(dateValue.getTime())) {
      return;
    }
    if (!signatures[fileId]) {
      signatures[fileId] = {};
    }
    signatures[fileId][dateValue.getTime()] = true;
  });
  return signatures;
}

function isCreditFileAlreadyProcessed_(fileId, lastUpdatedMillis, signatures) {
  return signatures[fileId] && signatures[fileId][lastUpdatedMillis];
}

function logCreditImport_(sheet, file, lastUpdatedMillis, rowsImported, notes) {
  const timestamp = utcNow_();
  const lastModifiedDate = new Date(lastUpdatedMillis);
  appendLogRow_(sheet, [timestamp, file.getId(), file.getName(), lastModifiedDate, rowsImported, notes]);
}

function addCreditsAuditSheet_(records, fileName) {
  const headers = ['Date', 'Amount', 'Property', 'Unit', 'Category', 'Subcategory', 'Status', 'Method', 'Payer'];
  const rows = records.map(function (record) {
    return [record.date, record.amount, record.property, record.unit, record.category, record.subcategory, record.status, record.method, record.counterparty];
  });
  const sheetNameBase = 'Credits ' + truncateName_(sanitizeSheetName_(fileName), 20);
  const sheet = createVersionedSheet_(sheetNameBase);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function createVersionedSheet_(baseName) {
  const spreadsheet = getActiveSpreadsheet_();
  let name = baseName;
  let counter = 1;
  while (spreadsheet.getSheetByName(name)) {
    counter++;
    name = baseName + VERSION_SUFFIX_SEPARATOR + counter;
  }
  return spreadsheet.insertSheet(name);
}
