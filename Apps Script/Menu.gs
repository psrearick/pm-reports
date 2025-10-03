function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PM Reports')
    .addItem('Import Credits', 'handleImportCredits_')
    .addSeparator()
    .addItem('Load Staging Data', 'handleLoadStaging_')
    .addItem('Save Staging Data', 'handleSaveStaging_')
    .addItem('Generate Report', 'handleGenerateReport_')
    .addItem('Export Report', 'handleExportReport_')
    .addToUi();
}

function handleImportCredits_() {
  runWithUiFeedback_('Importing credits...', function () {
    const results = importCredits();
    const imported = results.filter(function (result) { return !result.skipped; }).length;
    const skipped = results.length - imported;
    return 'Imported ' + imported + ' file(s), skipped ' + skipped + '.';
  });
}

function handleLoadStaging_() {
  runWithUiFeedback_('Loading staging data...', function () {
    loadStagingData();
    return 'Staging data refreshed.';
  });
}

function handleSaveStaging_() {
  runWithUiFeedback_('Saving staging changes...', function () {
    saveStagingData();
    return 'Master transactions updated.';
  });
}

function handleGenerateReport_() {
  runWithUiFeedback_('Generating report...', function () {
    const result = generateReport();
    return 'Report created: ' + result.name;
  });
}

function handleExportReport_() {
  runWithUiFeedback_('Preparing export...', function () {
    const exported = exportReportByLabel();
    return exported ? 'Export completed.' : 'Export cancelled or not found.';
  });
}

function runWithUiFeedback_(activityMessage, action) {
  const spreadsheet = SpreadsheetApp.getActive();
  showToast_(spreadsheet, activityMessage, 5);
  try {
    const resultMessage = action();
    showToast_(spreadsheet, resultMessage, 5);
  } catch (err) {
    showToast_(spreadsheet, 'Error: ' + err.message, 10);
    SpreadsheetApp.getUi().alert('PM Reports', err.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw err;
  }
}

function showToast_(spreadsheet, message, seconds) {
  spreadsheet.toast(message, 'PM Reports', seconds || 5);
}
