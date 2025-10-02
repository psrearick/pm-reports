function exportActiveReportToPdf() {
  const spreadsheet = SpreadsheetApp.getActive();
  exportSpreadsheetToFolder_(spreadsheet, spreadsheet.getName());
}

function exportReportByLabel() {
  const sheet = getSheetByName_(SHEET_NAMES.REPORT_LOG);
  const ui = SpreadsheetApp.getUi();
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('Export Report', 'No entries found in Report Log.', ui.ButtonSet.OK);
    return false;
  }
  const response = ui.prompt('Export Report', 'Enter the report label to export (leave blank for the most recent entry):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  const requestedLabel = normalizeStringValue_(response.getResponseText());
  const entry = findReportLogEntry_(sheet, requestedLabel);
  if (!entry) {
    ui.alert('Export Report', 'Report not found in log.', ui.ButtonSet.OK);
    return false;
  }
  const spreadsheet = SpreadsheetApp.openById(entry.spreadsheetId);
  exportSpreadsheetToFolder_(spreadsheet, spreadsheet.getName());
  ui.alert('Export Report', 'PDF export created for "' + spreadsheet.getName() + '".', ui.ButtonSet.OK);
  return true;
}

function exportSpreadsheetToFolder_(spreadsheet, reportLabel) {
  const folders = ensureDestinationFolders();
  const folderBaseName = truncateName_(sanitizeSheetName_(reportLabel), 80);
  const exportFolderInfo = createVersionedSubfolder(folders.exportsFolder, folderBaseName);
  const sheets = spreadsheet.getSheets().filter(function (sheet) {
    return sheet.getSheetName() !== SHEET_NAMES.TEMPLATE_BODY &&
      sheet.getSheetName() !== SHEET_NAMES.TEMPLATE_TOTALS &&
      sheet.getSheetName() !== SHEET_NAMES.TEMPLATE_AIRBNB &&
      !sheet.isSheetHidden();
  });
  sheets.forEach(function (sheet) {
    const pdfBlob = exportSheetToPdf_(spreadsheet, sheet);
    const fileName = sanitizeSheetName_(sheet.getName()) + '.pdf';
    pdfBlob.setName(fileName);
    exportFolderInfo.folder.createFile(pdfBlob);
  });
  return exportFolderInfo;
}

function exportSheetToPdf_(spreadsheet, sheet) {
  const exportUrl = buildSheetExportUrl_(spreadsheet.getId(), sheet.getSheetId());
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  return response.getBlob();
}

function findReportLogEntry_(sheet, requestedLabel) {
  const headerCount = LOG_HEADERS.REPORT.length;
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, headerCount).getValues();
  const normalizedLabel = normalizeStringValue_(requestedLabel);
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const label = normalizeStringValue_(row[1]);
    if (!normalizedLabel || label.toLowerCase() === normalizedLabel.toLowerCase()) {
      return {
        label: label,
        spreadsheetId: row[5],
        version: row[2]
      };
    }
  }
  return null;
}

function buildSheetExportUrl_(spreadsheetId, sheetId) {
  const baseUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export';
  const params = {
    format: 'pdf',
    portrait: false,
    size: 'letter',
    fitw: true,
    top_margin: 0.5,
    bottom_margin: 0.5,
    left_margin: 0.5,
    right_margin: 0.5,
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: false,
    fzr: false,
    gid: sheetId
  };
  const query = Object.keys(params).map(function (key) {
    return key + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return baseUrl + '?' + query;
}
