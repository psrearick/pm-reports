# Master Transaction Script

```js
// ADD CREDITS FUNCTION - Process TenantCloud Report
// Add this to your existing Property Management Google Apps Script

/* =========================
   UTILITY FUNCTIONS (if not already in your script)
   ========================= */

function notify_(msg) {
  try {
    SpreadsheetApp.getActive().toast(msg);
  } catch(_) {}
  Logger.log(msg);
}

function canonProp_(s) {
  var x = String(s || '').toLowerCase();
  x = x.replace(/&/g, ' and ');
  x = x.replace(/[^\w ]+/g, ' ');
  x = x.replace(/\b(and)\b/g, ' ');
  x = x.replace(/\s+/g, ' ').trim();
  return x;
}

function ensureMinRows_(sh, minIndex) {
  var mr = sh.getMaxRows();
  if(minIndex > mr) sh.insertRowsAfter(mr, minIndex - mr);
}

function ensureMinColumns_(sh, minIndex) {
  var mc = sh.getMaxColumns();
  if(minIndex > mc) sh.insertColumnsAfter(mc, minIndex - mc);
}

/* =========================
   CREDITS PROCESSING FUNCTIONS
   ========================= */

/**
 * Main function to add credits from TenantCloud report
 * Call this manually or set up as a trigger
 */
function addCreditsFromReport() {
  var master = SpreadsheetApp.getActive();
  var reportData = getReportData_(); // You'll need to implement this based on your data source

  if (!reportData || !reportData.length) {
    notify_('No report data found. Please check your data source.');
    return;
  }

  var processed = processCreditsData_(master, reportData);
  var tabs = getPropertyTabs_(master);

  // Refresh cover sheet and reorder tabs after adding credits
  buildCoverAboveLabelsOnly_(master, tabs);
  reorderTabsPreferred_(master);

  notify_('Credits processed: ' + processed.total + ' transactions added');
  Logger.log('Credits Summary: ' + JSON.stringify(processed.summary));
}

/**
 * Process credits data and add to appropriate property sheets
 */
function processCreditsData_(master, reportData) {
  var total = 0;
  var summary = {};

  for (var i = 0; i < reportData.length; i++) {
    var transaction = reportData[i];

    // Skip filtered categories
    if (shouldSkipTransaction_(transaction)) {
      continue;
    }

    // Find the appropriate property sheet
    var propertySheet = findPropertySheet_(master, transaction.property);
    if (!propertySheet) {
      Logger.log('No sheet found for property: ' + transaction.property);
      continue;
    }

    // Add the credit transaction
    if (addCreditToSheet_(propertySheet, transaction)) {
      total++;

      // Track summary
      var propName = propertySheet.getName();
      if (!summary[propName]) summary[propName] = 0;
      summary[propName]++;
    }

    // Periodic flush for performance
    if (total % (CONFIG.FLUSH_INTERVAL || 10) === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    }
  }

  return { total: total, summary: summary };
}

/**
 * Check if transaction should be skipped
 */
function shouldSkipTransaction_(transaction) {
  var category = String(transaction.category || '').toLowerCase();

  // Skip these categories
  var skipCategories = [
    'property general expense',
    'owner distribution'
  ];

  for (var i = 0; i < skipCategories.length; i++) {
    if (category.includes(skipCategories[i])) {
      return true;
    }
  }

  return false;
}

/**
 * Find property sheet using fuzzy matching
 */
function findPropertySheet_(master, propertyName) {
  if (!propertyName) return null;

  var canonicalProperty = fuzzyMatchProperty_(propertyName);
  var sheet = master.getSheetByName(canonicalProperty);

  if (!sheet) {
    // Try direct name match
    sheet = master.getSheetByName(propertyName);
  }

  if (!sheet) {
    // Try searching all sheets for best match
    var sheets = master.getSheets();
    var bestMatch = null;
    var bestScore = 0;

    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName === CONFIG.COVER_SHEET_NAME) continue;

      var score = calculateMatchScore_(propertyName, sheetName);
      if (score > bestScore && score > 0.5) {
        bestScore = score;
        bestMatch = sheets[i];
      }
    }

    sheet = bestMatch;
  }

  return sheet;
}

/**
 * Calculate match score between property names (0-1)
 */
function calculateMatchScore_(name1, name2) {
  var canon1 = canonProp_(name1);
  var canon2 = canonProp_(name2);

  if (canon1 === canon2) return 1.0;

  // Check for substring matches
  if (canon1.includes(canon2) || canon2.includes(canon1)) {
    return 0.8;
  }

  // Check word overlap
  var words1 = canon1.split(' ');
  var words2 = canon2.split(' ');
  var overlap = 0;

  for (var i = 0; i < words1.length; i++) {
    for (var j = 0; j < words2.length; j++) {
      if (words1[i] === words2[j] && words1[i].length > 2) {
        overlap++;
        break;
      }
    }
  }

  var maxWords = Math.max(words1.length, words2.length);
  return maxWords > 0 ? overlap / maxWords : 0;
}

/**
 * Add credit transaction to property sheet with enhanced error handling
 */
function addCreditToSheet_(sheet, transaction) {
  try {
    var headerRow = findHeaderRow_(sheet);
    var cols = detectTargetColumnsAtRowEnhanced_(sheet, headerRow);

    // Insert new row at the configured position
    var INSERT_AT = CONFIG.WRITABLE_START_ROW;
    ensureMinRows_(sheet, INSERT_AT);
    sheet.insertRowsBefore(INSERT_AT, 1);

    var r = INSERT_AT;

    // Extract unit from property or use provided unit
    var unit = extractUnit_(transaction);

    // Add basic transaction data
    if (cols.unitCol && unit) {
      sheet.getRange(r, cols.unitCol).setValue(String(unit));
    }

    if (cols.dateCol && transaction.date) {
      var date = parseReportDate_(transaction.date);
      sheet.getRange(r, cols.dateCol).setValue(date);
    }

    if (cols.explCol && transaction.category) {
      // Use category as explanation, add sub-category if available
      var explanation = String(transaction.category);
      if (transaction.subCategory && transaction.subCategory !== '-' &&
          transaction.subCategory.trim() !== '') {
        explanation += ' - ' + String(transaction.subCategory);
      }
      sheet.getRange(r, cols.explCol).setValue(explanation);
    }

    // Determine which column to use based on category
    var amount = parseAmount_(transaction.amount);
    if (amount && !isNaN(amount) && amount > 0) {
      var targetColumn = determineAmountColumn_(transaction.category, cols);
      if (targetColumn) {
        sheet.getRange(r, targetColumn).setValue(Math.abs(amount)); // Use absolute value for credits
      } else {
        Logger.log('No target column found for category: ' + transaction.category + ' in sheet: ' + sheet.getName());
      }
    }

    return true;

  } catch (e) {
    Logger.log('Error adding credit to ' + sheet.getName() + ': ' + e.message);
    Logger.log('Transaction data: ' + JSON.stringify(transaction));
    return false;
  }
}

/**
 * Extract unit number from various sources
 */
function extractUnit_(transaction) {
  var unit = '';

  // First try the unit field directly
  if (transaction.unit && transaction.unit !== '-' && String(transaction.unit).trim() !== '') {
    unit = String(transaction.unit).trim();
  }

  // If no unit field, try to extract from property name
  if (!unit && transaction.property) {
    var propStr = String(transaction.property);

    // Look for patterns like "Ohio & Bryden, Unit 224-A"
    var unitMatch = propStr.match(/, Unit (.+)$/);
    if (unitMatch) {
      unit = unitMatch[1].trim();
    } else {
      // Look for patterns like "Unit A-1" in the property string
      unitMatch = propStr.match(/Unit ([A-Za-z0-9-]+)/);
      if (unitMatch) {
        unit = unitMatch[1].trim();
      } else {
        // Look for unit numbers at the end like "1476" or "224-A"
        unitMatch = propStr.match(/\s+([A-Za-z]?[0-9]+[A-Za-z]?(?:-[A-Za-z0-9]+)?)$/);
        if (unitMatch) {
          unit = unitMatch[1].trim();
        }
      }
    }
  }

  return unit;
}

/**
 * Enhanced column detection to include Credit and Security Deposit columns
 */
function detectTargetColumnsAtRowEnhanced_(sh, headerRow) {
  var raw = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  var H = raw.map(function(h) { return String(h || '').replace(/\s+/g, ' ').trim(); });

  function find(re) {
    for (var i = 0; i < H.length; i++)
      if (re.test(H[i])) return i + 1;
    return null;
  }

  var unitCol   = find(/^unit$/i) || find(/unit\s*#?|unit\s*number/i) || 1;
  var debitsCol = find(/^debits?$/i) || find(/amount|price|total/i) || CONFIG.AMOUNT_COL_INDEX;
  var creditCol = find(/^credits?$/i) || find(/credit|income|revenue/i);
  var feeCol    = find(/^fees?$/i) || find(/fee/i);
  var secDepCol = find(/^security\s*deposits?$/i) || find(/security.*deposit|deposit/i);

  var dateCol   = find(/^date$/i) || find(/trans.*date|transaction.*date|date\s*paid/i);
  if (!dateCol) {
    var H2 = H.map(function(x) { return x.replace(/[^\w ]/g, ''); });
    for (var i = 0; i < H2.length; i++)
      if (/\bdate\b/i.test(H2[i])) { dateCol = i + 1; break; }
  }

  var explCol   = find(/^debit\/?credit explanation$/i) || find(/debit.*credit.*explan|explan|descr|description|memo/i);
  if (!explCol) {
    var H3 = H.map(function(x) { return x.replace(/[^\w ]/g, ''); });
    for (var j = 0; j < H3.length; j++)
      if (/(debit.*credit.*explan|explan|descr|description|memo)/i.test(H3[j])) { explCol = j + 1; break; }
  }

  var markupCol = find(/^markup included$/i) || find(/markup.*included|markup/i) || CONFIG.MARKUP_FLAG_COL_INDEX;
  var mrevCol   = find(/^markup revenue$/i)  || find(/markup.*revenue|markup.*rev/i) || CONFIG.MARKUP_REV_COL_INDEX;

  var maxCol = Math.max(unitCol, debitsCol, dateCol || 0, explCol || 0, markupCol, mrevCol,
                       creditCol || 0, feeCol || 0, secDepCol || 0);
  ensureMinColumns_(sh, maxCol);

  return {
    unitCol: unitCol,
    debitsCol: debitsCol,
    creditCol: creditCol,
    feeCol: feeCol,
    secDepCol: secDepCol,
    dateCol: dateCol,
    explCol: explCol,
    markupCol: markupCol,
    mrevCol: mrevCol
  };
}

/**
 * Determine which column to place the amount in based on category
 */
function determineAmountColumn_(category, cols) {
  var cat = String(category || '').toLowerCase();

  // Security Deposits
  if (cat.includes('deposit')) {
    return cols.secDepCol;
  }

  // Fees (tenant charges & fees)
  if (cat.includes('tenant charges') || cat.includes('fees') || cat.includes('late payment fee')) {
    return cols.feeCol;
  }

  // Regular rent and other income goes to Credits
  if (cat.includes('rent') || cat === '' || cat.includes('income')) {
    return cols.creditCol;
  }

  // Default to credit column for unspecified categories
  return cols.creditCol;
}

/**
 * Parse amount from report (handles currency formatting)
 */
function parseAmount_(amountStr) {
  if (!amountStr) return 0;

  var cleanAmount = String(amountStr).replace(/[\$,\s]/g, '');
  return Number(cleanAmount);
}

/**
 * Parse date from report format
 */
function parseReportDate_(dateStr) {
  if (!dateStr) return new Date();

  try {
    // Handle various date formats from the report
    var dateString = String(dateStr).trim();

    // Try parsing common formats
    var parsedDate = new Date(dateString);

    if (isNaN(parsedDate.getTime())) {
      // Fallback to manual parsing if needed
      // Format: "Sep 1, 2025" or similar
      var parts = dateString.match(/(\w+)\s+(\d+),\s+(\d+)/);
      if (parts) {
        var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        var monthIndex = monthNames.indexOf(parts[1]);
        if (monthIndex !== -1) {
          parsedDate = new Date(parseInt(parts[3]), monthIndex, parseInt(parts[2]));
        }
      }
    }

    return isNaN(parsedDate.getTime()) ? new Date() : parsedDate;

  } catch (e) {
    Logger.log('Date parsing error: ' + e.message + ' for date: ' + dateStr);
    return new Date();
  }
}

/**
 * Get report data from Google Sheets XLSX file
 */
function getReportData_() {
  // Your TenantCloud report spreadsheet
  var REPORT_SPREADSHEET_ID = '1uueJ5ttzVZe4FT9t0FbyuLEmTC-smnjV';

  try {
    var reportSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    var sheet = reportSpreadsheet.getSheets()[0]; // Use first sheet

    return parseReportFromTenantCloudSheet_(sheet);

  } catch (e) {
    Logger.log('Error accessing report spreadsheet: ' + e.message);
    notify_('Error accessing TenantCloud report. Check permissions and spreadsheet ID.');
    return [];
  }
}

/**
 * Parse TenantCloud report from Google Sheets
 */
function parseReportFromTenantCloudSheet_(sheet) {
  var data = sheet.getDataRange().getValues();
  var transactions = [];

  // Find the header row that contains "Date paid", "Amount paid", etc.
  var headerRowIndex = -1;
  var columnMap = {};

  for (var i = 0; i < Math.min(20, data.length); i++) {
    var row = data[i];
    var rowText = row.map(function(cell) { return String(cell || '').toLowerCase(); }).join(' ');

    if (rowText.includes('date paid') && rowText.includes('amount paid')) {
      headerRowIndex = i;

      // Map column positions
      for (var j = 0; j < row.length; j++) {
        var header = String(row[j] || '').toLowerCase().trim();
        if (header.includes('date paid')) columnMap.dateCol = j;
        else if (header.includes('amount paid')) columnMap.amountCol = j;
        else if (header.includes('property name')) columnMap.propertyCol = j;
        else if (header.includes('unit')) columnMap.unitCol = j;
        else if (header.includes('category') && !header.includes('sub')) columnMap.categoryCol = j;
        else if (header.includes('sub-category')) columnMap.subCategoryCol = j;
        else if (header.includes('payment status')) columnMap.statusCol = j;
        else if (header.includes('payment method')) columnMap.methodCol = j;
        else if (header.includes('payer') || header.includes('payee')) columnMap.payerCol = j;
      }
      break;
    }
  }

  if (headerRowIndex === -1) {
    Logger.log('Could not find header row in TenantCloud report');
    notify_('Could not find proper headers in TenantCloud report. Please check the file format.');
    return [];
  }

  // Process data rows
  for (var i = headerRowIndex + 1; i < data.length; i++) {
    var row = data[i];

    // Skip empty rows or total rows
    var firstCell = String(row[0] || '').trim();
    if (!firstCell || firstCell.toLowerCase().includes('total') ||
        firstCell.toLowerCase().includes('grand total')) {
      continue;
    }

    // Extract transaction data
    var transaction = {
      date: columnMap.dateCol !== undefined ? row[columnMap.dateCol] : '',
      amount: columnMap.amountCol !== undefined ? row[columnMap.amountCol] : '',
      property: columnMap.propertyCol !== undefined ? row[columnMap.propertyCol] : '',
      unit: columnMap.unitCol !== undefined ? row[columnMap.unitCol] : '',
      category: columnMap.categoryCol !== undefined ? row[columnMap.categoryCol] : '',
      subCategory: columnMap.subCategoryCol !== undefined ? row[columnMap.subCategoryCol] : '',
      status: columnMap.statusCol !== undefined ? row[columnMap.statusCol] : '',
      method: columnMap.methodCol !== undefined ? row[columnMap.methodCol] : '',
      payer: columnMap.payerCol !== undefined ? row[columnMap.payerCol] : ''
    };

    // Clean up property name (remove extra text after property name)
    if (transaction.property) {
      transaction.property = cleanPropertyName_(String(transaction.property));
    }

    // Only add if it has required data and valid amount
    if (transaction.date && transaction.amount && transaction.property &&
        parseAmount_(transaction.amount) > 0) {
      transactions.push(transaction);
    }
  }

  Logger.log('Found ' + transactions.length + ' transactions in TenantCloud report');
  return transactions;
}

/**
 * Clean property name from TenantCloud format
 */
function cleanPropertyName_(propertyStr) {
  var cleaned = String(propertyStr).trim();

  // Handle cases like "Ohio & Bryden, Unit 224-A" -> "Ohio & Bryden"
  if (cleaned.includes(', Unit ')) {
    cleaned = cleaned.split(', Unit ')[0];
  }

  // Handle cases like "Park and State St, Unit 16-1 W Park St" -> "Park and State St"
  if (cleaned.includes(', Unit ')) {
    cleaned = cleaned.split(', Unit ')[0];
  }

  // Remove common suffixes that might interfere with matching
  var suffixesToRemove = [', Single-family', ' Ave.', ' St.', ' S.'];
  for (var i = 0; i < suffixesToRemove.length; i++) {
    if (cleaned.endsWith(suffixesToRemove[i])) {
      cleaned = cleaned.replace(suffixesToRemove[i], '').trim();
    }
  }

  // Handle specific property name variations
  var nameMap = {
    '1505-1515 Franklin Park S.': '1505-1515 Franklin Park',
    '189 W. Patterson Ave': '189 Patterson',
    '22 Wilson Ave': '22 Wilson',
    'Park and State St': 'Park and State',
    '1476 High St': '1476 S High St',
    '196 Miller Ave': '196 Miller',
    '2536 Adams Ave.': '2536 Adams',
    '705 Ann': '705 Ann'
  };

  return nameMap[cleaned] || cleaned;
}

/**
 * Sample data for testing - replace with actual implementation
 */
function getSampleReportData_() {
  return [
    {
      date: 'Sep 1, 2025',
      amount: '$1000.00',
      property: 'Ohio & Bryden',
      unit: '224-A',
      category: 'Rent',
      subCategory: '',
      status: 'Success',
      method: 'ACH',
      payer: 'Jacob Blount'
    },
    {
      date: 'Sep 4, 2025',
      amount: '$1095.00',
      property: 'Park and State St',
      unit: '16-2 W Park St',
      category: 'Deposit',
      subCategory: '',
      status: 'Success',
      method: 'ACH',
      payer: 'Adam Muschott'
    }
  ];
}

/**
 * Enhanced column detection to include Credit and Security Deposit columns
 */
function detectTargetColumnsAtRowEnhanced_(sh, headerRow) {
  var cols = detectTargetColumnsAtRow_(sh, headerRow);

  // Add credit column detection if not already found
  if (!cols.creditCol) {
    var raw = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
    var H = raw.map(function(h) { return String(h || '').replace(/\s+/g, ' ').trim(); });

    function find(re) {
      for (var i = 0; i < H.length; i++)
        if (re.test(H[i])) return i + 1;
      return null;
    }

    cols.creditCol = find(/^credits?$/i) || find(/credit|income|revenue/i);

    if (!cols.secDepCol) {
      cols.secDepCol = find(/^security\s*deposits?$/i) || find(/security.*deposit|deposits/i);
    }
  }

  return cols;
}

/**
 * Test function to check if we can access the report and see its structure
 */
function testReportAccess() {
  var REPORT_SPREADSHEET_ID = '1uueJ5ttzVZe4FT9t0FbyuLEmTC-smnjV';

  try {
    var reportSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    var sheet = reportSpreadsheet.getSheets()[0];

    Logger.log('Successfully accessed report spreadsheet: ' + reportSpreadsheet.getName());
    Logger.log('Sheet name: ' + sheet.getName());
    Logger.log('Sheet dimensions: ' + sheet.getLastRow() + ' rows, ' + sheet.getLastColumn() + ' columns');

    // Show first few rows to understand structure
    var sampleData = sheet.getRange(1, 1, Math.min(10, sheet.getLastRow()), sheet.getLastColumn()).getValues();
    Logger.log('Sample data from first 10 rows:');
    for (var i = 0; i < sampleData.length; i++) {
      Logger.log('Row ' + (i + 1) + ': ' + JSON.stringify(sampleData[i]));
    }

    notify_('Report access successful! Check logs for details.');

  } catch (e) {
    Logger.log('Error accessing report: ' + e.message);
    notify_('Error accessing report: ' + e.message);
  }
}

/**
 * Main function to add credits from TenantCloud report with enhanced logging
 */
function addCreditsFromReport() {
  var master = SpreadsheetApp.getActive();

  Logger.log('Starting credit processing...');
  notify_('Starting to process TenantCloud report...');

  var reportData = getReportData_();

  if (!reportData || !reportData.length) {
    notify_('No report data found. Please check your data source.');
    return;
  }

  Logger.log('Found ' + reportData.length + ' total transactions in report');

  var processed = processCreditsData_(master, reportData);
  var tabs = getPropertyTabs_(master);

  // Refresh cover sheet and reorder tabs after adding credits
  buildCoverAboveLabelsOnly_(master, tabs);
  reorderTabsPreferred_(master);

  var message = 'Credits processed: ' + processed.total + ' transactions added';
  notify_(message);
  Logger.log(message);
  Logger.log('Processing summary by property: ' + JSON.stringify(processed.summary));
}

/* =========================
   MENU ADDITIONS
   ========================= */

// Enhanced menu with test function
function onOpenEnhanced() {
  SpreadsheetApp.getUi()
    .createMenu('Property Management')
    .addItem('Import Current Month Only', 'importOnlyForCurrentMonth')
    .addItem('Post Debits & Clean Empty Rows', 'postDebitsMonthlyNow')
    .addSeparator()
    .addItem('Add Credits from Report', 'addCreditsFromReport')
    .addItem('Test Report Access', 'testReportAccess')
    .addSeparator()
    .addItem('Clean Empty Rows (Manual)', 'cleanAllEmptyRows')
    .addSeparator()
    .addItem('Create Import Trigger (1st, 6:15 AM)', 'createMonthlyImportTrigger1')
    .addItem('Create Post Trigger (19th, 9:15 PM)', 'createMonthlyPostTrigger19')
    .addItem('Remove All Triggers', 'removeAllMonthlyTriggers')
    .addItem('List Current Triggers', 'listCurrentTriggers')
    .addToUi();
}
```
