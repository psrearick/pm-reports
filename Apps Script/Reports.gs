function generateReport() {
  const settings = getReportSettings_();
  if (!settings.reportLabel) {
    throw new Error('Report Label is required.');
  }
  const folders = ensureDestinationFolders();
  const reportData = buildReportData_(settings.startDate, settings.endDate, settings.property);
  if (!reportData.properties.length) {
    throw new Error('No transactions found for the selected period.');
  }
  const created = createVersionedSpreadsheet(settings.reportLabel, folders.reportsFolder);
  const spreadsheet = created.spreadsheet;
  renderPropertySheets_(spreadsheet, reportData, settings);
  renderSummarySheet_(spreadsheet, reportData, settings);
  renderAirbnbSheet_(spreadsheet, reportData, settings);
  removeDefaultSheet_(spreadsheet);
  reorderReportSheets_(spreadsheet, reportData);
  logReportGeneration_(reportData, settings, created, spreadsheet);
  return {
    spreadsheetId: spreadsheet.getId(),
    name: created.name
  };
}

function getReportSettings_() {
  const sheet = getSheetByName_(STAGING_CONTROL.SHEET);
  const startDate = parseOptionalDate_(sheet.getRange(STAGING_CONTROL.START_DATE_CELL).getValue());
  const endDate = parseOptionalDate_(sheet.getRange(STAGING_CONTROL.END_DATE_CELL).getValue());
  const reportLabel = normalizeStringValue_(sheet.getRange(STAGING_CONTROL.REPORT_LABEL_CELL).getValue());
  const property = normalizeStringValue_(sheet.getRange(STAGING_CONTROL.PROPERTY_CELL).getValue());
  return {
    startDate: startDate,
    endDate: endDate,
    reportLabel: reportLabel,
    property: property
  };
}

function buildReportData_(startDate, endDate, propertyFilter) {
  const propertyConfigs = getPropertiesConfig();
  const propertyMap = {};
  propertyConfigs.forEach(function (config) {
    propertyMap[config.name] = config;
  });
  const transactionsByProperty = collectTransactionsForRange_(startDate, endDate, propertyMap, propertyFilter);
  const properties = [];
  Object.keys(transactionsByProperty).forEach(function (propertyName) {
    const propertyTransactions = transactionsByProperty[propertyName];
    const propertyConfig = propertyMap[propertyName] || {
      name: propertyName,
      maf: 0,
      markup: 0,
      hasAirbnb: false,
      airbnbPercent: 0,
      adminFee: 0,
      adminFeeEnabled: false,
      order: null
    };
    const totals = calculatePropertyTotals_(propertyTransactions, propertyConfig);
    properties.push({
      name: propertyName,
      property: propertyConfig,
      transactions: propertyTransactions,
      totals: totals
    });
  });
  properties.sort(function (a, b) {
    const orderA = (a.property && typeof a.property.order === 'number') ? a.property.order : null;
    const orderB = (b.property && typeof b.property.order === 'number') ? b.property.order : null;
    if (orderA !== null && orderB !== null && orderA !== orderB) {
      return orderA - orderB;
    }
    if (orderA !== null && orderB === null) {
      return -1;
    }
    if (orderA === null && orderB !== null) {
      return 1;
    }
    return a.name.localeCompare(b.name);
  });
  return {
    startDate: startDate,
    endDate: endDate,
    properties: properties
  };
}

function collectTransactionsForRange_(startDate, endDate, propertyMap, propertyFilter) {
  const sheet = getSheetByName_(SHEET_NAMES.TRANSACTIONS);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const results = {};
  if (lastRow < 2) {
    return results;
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headerIndex = getHeaderIndexMap_(headers);
  const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = range.getValues();
  const properties = propertyFilter ? ensureArray_(propertyFilter).filter(function (item) { return !!normalizeStringValue_(item); }) : [];
  const matchAllProperties = properties.length === 0;
  values.forEach(function (row) {
    if (row.every(isBlank_)) {
      return;
    }
    const deleted = toBool(row[headerIndex['Deleted']]);
    if (deleted) {
      return;
    }
    const dateValue = row[headerIndex['Date']];
    if (!withinDateRange_(dateValue, startDate, endDate)) {
      return;
    }
    const propertyName = row[headerIndex['Property']];
    if (!propertyName) {
      return;
    }
    if (!matchAllProperties && properties.indexOf(propertyName) === -1) {
      return;
    }
    const transaction = {
      id: row[headerIndex['Transaction ID']],
      date: normalizeDateValue_(dateValue),
      property: propertyName,
      unit: row[headerIndex['Unit']] || '',
      credits: parseCurrency_(row[headerIndex['Credits']]),
      fees: parseCurrency_(row[headerIndex['Fees']]),
      debits: parseCurrency_(row[headerIndex['Debits']]),
      securityDeposits: parseCurrency_(row[headerIndex['Security Deposits']]),
      explanation: row[headerIndex['Debit/Credit Explanation']] || '',
      markupIncluded: toBool(row[headerIndex['Markup Included']]),
      markupRevenue: parseCurrency_(row[headerIndex['Markup Revenue']]),
      internalNotes: row[headerIndex['Internal Notes']] || ''
    };
    const propertyConfig = propertyMap[propertyName];
    if (propertyConfig && propertyConfig.hasAirbnb) {
      const internal = (transaction.internalNotes || '').toLowerCase();
      transaction.isAirbnb = internal.indexOf('airbnb') !== -1;
    } else {
      transaction.isAirbnb = false;
    }
    if (!results[propertyName]) {
      results[propertyName] = [];
    }
    results[propertyName].push(transaction);
  });
  return results;
}

function calculatePropertyTotals_(transactions, propertyConfig) {
  const totals = {
    credits: 0,
    fees: 0,
    debits: 0,
    markupRevenue: 0,
    securityDeposits: 0,
    securityDepositMailFees: 0,
    unitCount: 0,
    adminFeeApplied: false,
    adminFeeAmount: 0,
    airbnbIncome: 0,
    airbnbTotal: 0,
    airbnbFee: 0,
    newLeaseFees: 0,
    renewalFees: 0
  };
  const unitSet = new Set();
  transactions.forEach(function (transaction) {
    totals.fees += transaction.fees || 0;
    totals.debits += transaction.debits || 0;
    totals.markupRevenue += transaction.markupRevenue || 0;
    totals.securityDeposits += transaction.securityDeposits || 0;
    if (transaction.unit && transaction.unit.toString().trim() && transaction.unit.toString().toUpperCase() !== 'CAM') {
      unitSet.add(transaction.unit.toString().trim());
    }
    if (transaction.isAirbnb) {
      totals.airbnbIncome += transaction.credits || 0;
    } else {
      totals.credits += transaction.credits || 0;
    }
    if (transaction.explanation === 'New Lease Fee') {
      totals.newLeaseFees += transaction.debits || 0;
    }
    if (transaction.explanation === 'Renewal Fee') {
      totals.renewalFees += transaction.debits || 0;
    }
    if (normalize("Security Deposit Return Mail Fee") === normalize(transaction.explanation)) {
        totals.securityDepositMailFees += transaction.debits || 0;
    }
  });
  totals.unitCount = unitSet.size;
  const mafRate = propertyConfig.maf || 0;
  const mafFromCredits = totals.credits * mafRate;
  let totalMaf = mafFromCredits;
  const adminFeeApplied = toBool(propertyConfig.adminFeeEnabled);
  if (adminFeeApplied && propertyConfig.adminFee) {
    totals.adminFeeApplied = true;
    totals.adminFeeAmount = propertyConfig.adminFee;
  }
  let totalDebits = totals.debits + totals.markupRevenue + totalMaf + totals.adminFeeAmount;
  let combinedCredits = totals.credits + totals.securityDeposits;
  if (propertyConfig.hasAirbnb) {
    totals.airbnbTotal = totals.airbnbIncome;
    totals.airbnbFee = Math.round(totals.airbnbTotal * (propertyConfig.airbnbPercent || 0) * 100) / 100;
    combinedCredits += totals.airbnbTotal;
    totalDebits += totals.airbnbFee;
  }
  const dueToOwners = combinedCredits - totalDebits;
  const totalToPm = totals.markupRevenue + totalMaf + totals.airbnbFee + totals.adminFeeAmount;
  totals.totalMaf = Math.round(totalMaf * 100) / 100;
  totals.totalDebits = Math.round(totalDebits * 100) / 100;
  totals.combinedCredits = Math.round(combinedCredits * 100) / 100;
  totals.dueToOwners = Math.round(dueToOwners * 100) / 100;
  totals.totalToPm = Math.round(totalToPm * 100) / 100;
  totals.totalCredits = Math.round(totals.credits * 100) / 100;
  totals.totalFees = Math.round(totals.fees * 100) / 100;
  totals.totalMarkupRevenue = Math.round(totals.markupRevenue * 100) / 100;
  totals.totalSecurityDeposits = Math.round(totals.securityDeposits * 100) / 100;
  totals.totalSecurityDepositMailFees = Math.round(totals.securityDepositMailFees * 100) / 100;
  return totals;
}

function renderPropertySheets_(spreadsheet, reportData, settings) {
  const dateRangeLabel = formatDateRange_(reportData.startDate, reportData.endDate);
  reportData.properties.forEach(function (property) {
    const context = {
      placeholders: {
        'PROPERTY NAME': property.name,
        'REPORT LABEL': settings.reportLabel,
        'REPORT PERIOD': dateRangeLabel,
        'TOTAL CREDITS': formatCurrency_(property.totals.totalCredits),
        'TOTAL FEES': formatCurrency_(property.totals.totalFees),
        'TOTAL MARKUP': formatCurrency_(property.totals.totalMarkupRevenue),
        'TOTAL MAF': formatCurrency_(property.totals.totalMaf),
        'TOTAL SECURITY DEPOSITS': formatCurrency_(property.totals.totalSecurityDeposits),
        'TOTAL SECURITY DEPOSIT RETURN MAIL FEES': formatCurrency_(property.totals.totalSecurityDepositMailFees),
        'TOTAL DEBITS': formatCurrency_(property.totals.totalDebits),
        'COMBINED CREDITS': formatCurrency_(property.totals.combinedCredits),
        'TOTAL RECEIPTS': formatCurrency_(property.totals.combinedCredits),
        'DUE TO OWNERS': formatCurrency_(property.totals.dueToOwners),
        'TOTAL TO PM': formatCurrency_(property.totals.totalToPm),
        'UNIT COUNT': property.totals.unitCount,
        'ADMIN FEE APPLIED': property.totals.adminFeeApplied ? 'Yes' : 'No',
        'ADMIN FEE AMOUNT': formatCurrency_(property.totals.adminFeeAmount)
      },
      rows: {
        transactions: property.transactions.map(function (transaction) {
          return {
            unit: transaction.unit,
            credits: formatCurrency_(transaction.credits),
            fees: formatCurrency_(transaction.fees),
            debits: formatCurrency_(transaction.debits),
            securityDeposits: formatCurrency_(transaction.securityDeposits),
            date: formatDateDisplay_(transaction.date),
            explanation: transaction.explanation,
            markupIncluded: transaction.markupIncluded ? 'Yes' : 'No',
            markupRevenue: formatCurrency_(transaction.markupRevenue),
            internalNotes: transaction.internalNotes || ''
          };
        })
      },
      flags: {
        has_airbnb: property.property.hasAirbnb,
        has_admin_fee: property.totals.adminFeeApplied
      }
    };
    if (property.property.hasAirbnb) {
      context.placeholders['AIRBNB TOTAL'] = formatCurrency_(property.totals.airbnbTotal);
      context.placeholders['AIRBNB FEE'] = formatCurrency_(property.totals.airbnbFee);
    }
    const sheetName = sanitizeSheetName_(property.name);
    renderTemplateSheet(SHEET_NAMES.TEMPLATE_BODY, spreadsheet, sheetName, context);
  });
}

function renderSummarySheet_(spreadsheet, reportData, settings) {
  const context = {
    placeholders: {
      'REPORT LABEL': settings.reportLabel,
      'REPORT PERIOD': formatDateRange_(reportData.startDate, reportData.endDate)
    },
    rows: {
      summary: reportData.properties.map(function (property) {
        return {
          property: property.name,
          dueToOwners: formatCurrency_(property.totals.dueToOwners),
          totalToPm: formatCurrency_(property.totals.totalToPm),
          totalFees: formatCurrency_(property.totals.totalFees),
          totalMaf: formatCurrency_(property.totals.totalMaf),
          totalMarkup: formatCurrency_(property.totals.totalMarkupRevenue),
          combinedCredits: formatCurrency_(property.totals.combinedCredits),
          newLeaseFees: formatCurrency_(property.totals.newLeaseFees),
          renewalFees: formatCurrency_(property.totals.renewalFees)
        };
      })
    },
    flags: {}
  };
  const aggregated = reportData.properties.reduce(function (acc, property) {
    const totals = property.totals || {};
    acc.dueToOwners += totals.dueToOwners || 0;
    acc.totalToPm += totals.totalToPm || 0;
    acc.totalFees += totals.totalFees || 0;
    acc.totalMaf += totals.totalMaf || 0;
    acc.totalMarkup += totals.totalMarkupRevenue || 0;
    acc.totalDebits += totals.totalDebits || 0;
    acc.totalCredits += totals.totalCredits || 0;
    acc.totalSecurityDeposits += totals.totalSecurityDeposits || 0;
    acc.totalSecurityDepositMailFees += totals.totalSecurityDepositMailFees || 0;
    acc.combinedCredits += totals.combinedCredits || 0;
    acc.newLeaseFees += totals.newLeaseFees || 0;
    acc.renewalFees += totals.renewalFees || 0;
    acc.airbnbTotal += totals.airbnbTotal || 0;
    acc.airbnbFee += totals.airbnbFee || 0;
    acc.propertyCount += 1;
    return acc;
  }, {
    dueToOwners: 0,
    totalToPm: 0,
    totalFees: 0,
    totalMaf: 0,
    totalMarkup: 0,
    totalDebits: 0,
    totalCredits: 0,
    totalSecurityDeposits: 0,
    totalSecurityDepositMailFees: 0,
    combinedCredits: 0,
    newLeaseFees: 0,
    renewalFees: 0,
    airbnbTotal: 0,
    airbnbFee: 0,
    propertyCount: 0
  });
  context.placeholders['SUMMARY PROPERTY COUNT'] = aggregated.propertyCount;
  context.placeholders['SUMMARY TOTAL DUE TO OWNERS'] = formatCurrency_(aggregated.dueToOwners);
  context.placeholders['SUMMARY TOTAL TO PM'] = formatCurrency_(aggregated.totalToPm);
  context.placeholders['SUMMARY TOTAL FEES'] = formatCurrency_(aggregated.totalFees);
  context.placeholders['SUMMARY TOTAL MAF'] = formatCurrency_(aggregated.totalMaf);
  context.placeholders['SUMMARY TOTAL MARKUP'] = formatCurrency_(aggregated.totalMarkup);
  context.placeholders['SUMMARY TOTAL DEBITS'] = formatCurrency_(aggregated.totalDebits);
  context.placeholders['SUMMARY TOTAL CREDITS'] = formatCurrency_(aggregated.totalCredits);
  context.placeholders['SUMMARY TOTAL SECURITY DEPOSITS'] = formatCurrency_(aggregated.totalSecurityDeposits);
  context.placeholders['SUMMARY TOTAL SECURITY DEPOSIT RETURN MAIL FEES'] = formatCurrency_(aggregated.totalSecurityDepositMailFees);
  context.placeholders['SUMMARY COMBINED CREDITS'] = formatCurrency_(aggregated.combinedCredits);
  context.placeholders['SUMMARY TOTAL NEW LEASE FEES'] = formatCurrency_(aggregated.newLeaseFees);
  context.placeholders['SUMMARY TOTAL RENEWAL FEES'] = formatCurrency_(aggregated.renewalFees);
  context.placeholders['SUMMARY AIRBNB TOTAL'] = formatCurrency_(aggregated.airbnbTotal);
  context.placeholders['SUMMARY AIRBNB FEE'] = formatCurrency_(aggregated.airbnbFee);
  renderTemplateSheet(SHEET_NAMES.TEMPLATE_TOTALS, spreadsheet, 'Summary', context);
}

function renderAirbnbSheet_(spreadsheet, reportData, settings) {
  const airbnbRows = reportData.properties.filter(function (property) {
    return property.property.hasAirbnb && property.totals.airbnbTotal > 0;
  }).map(function (property) {
    return {
      property: property.name,
      income: formatCurrency_(property.totals.airbnbTotal),
      collectionFee: formatCurrency_(property.totals.airbnbFee)
    };
  });
  if (!airbnbRows.length) {
    return;
  }
  const context = {
    placeholders: {
      'REPORT LABEL': settings.reportLabel,
      'REPORT PERIOD': formatDateRange_(reportData.startDate, reportData.endDate)
    },
    rows: {
      airbnb: airbnbRows
    },
    flags: {}
  };
  renderTemplateSheet(SHEET_NAMES.TEMPLATE_AIRBNB, spreadsheet, 'Airbnb', context);
}

function formatDateRange_(startDate, endDate) {
  if (!startDate && !endDate) {
    return '';
  }
  const start = startDate ? formatDateDisplay_(startDate) : '';
  const end = endDate ? formatDateDisplay_(endDate) : '';
  if (start && end) {
    return start + ' - ' + end;
  }
  return start || end;
}

function removeDefaultSheet_(spreadsheet) {
  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(defaultSheet);
  }
}

function reorderReportSheets_(spreadsheet, reportData) {
  const visibleSheets = spreadsheet.getSheets().filter(function (sheet) {
    return !sheet.isSheetHidden();
  });
  if (visibleSheets.length < 2) {
    // Nothing to reorder; avoid triggering "remove all visible sheets" errors.
    return;
  }
  const desiredOrder = [];
  const summarySheet = spreadsheet.getSheetByName('Summary');
  if (summarySheet) {
    desiredOrder.push(summarySheet);
  }
  const airbnbSheet = spreadsheet.getSheetByName('Airbnb');
  if (airbnbSheet) {
    desiredOrder.push(airbnbSheet);
  }
  reportData.properties.forEach(function (property) {
    const sheetName = sanitizeSheetName_(property.name);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet && desiredOrder.indexOf(sheet) === -1) {
      desiredOrder.push(sheet);
    }
  });
  desiredOrder.forEach(function (sheet, index) {
    sheet.activate();
    spreadsheet.moveActiveSheet(index + 1);
  });
  if (summarySheet) {
    summarySheet.activate();
  }
}

function logReportGeneration_(reportData, settings, created, spreadsheet) {
  const logSheet = ensureLogSheet_(SHEET_NAMES.REPORT_LOG, LOG_HEADERS.REPORT);
  const version = deriveReportVersion_(settings.reportLabel, created.name);
  const propertiesIncluded = reportData.properties.map(function (property) { return property.name; }).join(', ');
  const adminDetails = reportData.properties.map(function (property) {
    if (property.totals.adminFeeApplied) {
      return property.name + '=Applied ' + formatCurrency_(property.totals.adminFeeAmount);
    }
    return property.name + '=Default';
  }).join('; ');
  appendLogRow_(logSheet, [
    utcNow_(),
    settings.reportLabel,
    version,
    settings.startDate || '',
    settings.endDate || '',
    spreadsheet.getId(),
    spreadsheet.getUrl(),
    propertiesIncluded,
    adminDetails
  ]);
}

function deriveReportVersion_(label, actualName) {
  if (!label) {
    return actualName;
  }
  if (actualName === label) {
    return '1';
  }
  const pattern = new RegExp('^' + escapeForRegExp_(label) + '_' + '(\\d+)$');
  const match = actualName.match(pattern);
  if (match) {
    return match[1];
  }
  return actualName;
}
