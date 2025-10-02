function exportActiveReportToPdf() {
  const spreadsheet = SpreadsheetApp.getActive();
  const reportLabel = spreadsheet.getName();
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
}

function exportSheetToPdf_(spreadsheet, sheet) {
  const exportUrl = buildSheetExportUrl_(spreadsheet.getId(), sheet.getSheetId());
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  return response.getBlob();
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
