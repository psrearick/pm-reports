function renderTemplateSheet(templateSheetName, destinationSpreadsheet, outputSheetName, context) {
  context = context || {};
  context.placeholders = context.placeholders || {};
  context.rows = context.rows || {};
  context.flags = context.flags || {};
  const templateSheet = getSheetByName_(templateSheetName);
  const targetSheet = templateSheet.copyTo(destinationSpreadsheet);
  targetSheet.setName(outputSheetName);
  targetSheet.activate();
  const warnings = applyTemplateContext_(targetSheet, context);
  return { sheet: targetSheet, warnings: warnings };
}

function applyTemplateContext_(sheet, context) {
  const warnings = [];
  processRowBlocks_(sheet, context, warnings);
  processConditionalBlocks_(sheet, context);
  replaceScalarPlaceholders_(sheet, context.placeholders, warnings);
  return warnings;
}

function processRowBlocks_(sheet, context, warnings) {
  for (;;) {
    const matches = sheet.createTextFinder('[[ROW').findAll();
    if (!matches || !matches.length) {
      break;
    }
    const markerCell = matches[matches.length - 1];
    const markerValue = markerCell.getValue();
    const match = markerValue.match(/^\[\[ROW\s+([A-Za-z0-9_\-]+)\]\]$/);
    if (!match) {
      markerCell.clear();
      continue;
    }
    const blockName = match[1];
    renderRowBlockInstance_(sheet, markerCell, blockName, context, warnings);
  }
}

function renderRowBlockInstance_(sheet, markerCell, blockName, context, warnings) {
  const startRow = markerCell.getRow();
  const lastColumn = sheet.getLastColumn() || 1;
  const endRow = findMarkerRow_(sheet, startRow + 1, '[[ENDROW]]');
  if (!endRow) {
    throw new Error('Missing [[ENDROW]] marker for block ' + blockName + ' in sheet ' + sheet.getName());
  }
  const templateStartRow = startRow + 1;
  const templateRowCount = endRow - templateStartRow;
  if (templateRowCount <= 0) {
    sheet.deleteRow(endRow);
    sheet.deleteRow(startRow);
    return;
  }
  const templateRange = sheet.getRange(templateStartRow, 1, templateRowCount, lastColumn);
  const templateValues = templateRange.getValues();
  const blockData = context.rows[blockName] || [];
  if (!blockData.length) {
    sheet.deleteRows(templateStartRow, templateRowCount);
  } else {
    const totalRowsNeeded = templateRowCount * blockData.length;
    const extraRowsNeeded = totalRowsNeeded - templateRowCount;
    if (extraRowsNeeded > 0) {
      sheet.insertRowsAfter(templateStartRow + templateRowCount - 1, extraRowsNeeded);
    }
    for (let i = 0; i < blockData.length; i++) {
      const offset = templateStartRow + i * templateRowCount;
      const targetRange = sheet.getRange(offset, 1, templateRowCount, lastColumn);
      if (i > 0) {
        templateRange.copyTo(targetRange, { contentsOnly: false });
      }
      const renderedValues = renderTemplateBlockValues_(templateValues, blockData[i], context.placeholders, warnings);
      targetRange.setValues(renderedValues);
    }
  }
  const newEndRow = findMarkerRow_(sheet, startRow + 1, '[[ENDROW]]');
  if (newEndRow) {
    sheet.deleteRow(newEndRow);
  }
  sheet.deleteRow(startRow);
}

function renderTemplateBlockValues_(templateValues, rowData, placeholders, warnings) {
  const output = [];
  for (let r = 0; r < templateValues.length; r++) {
    const row = templateValues[r];
    const renderedRow = [];
    for (let c = 0; c < row.length; c++) {
      renderedRow.push(replaceValuePlaceholders_(row[c], rowData, placeholders, warnings));
    }
    output.push(renderedRow);
  }
  return output;
}

function processConditionalBlocks_(sheet, context) {
  for (;;) {
    const matches = sheet.createTextFinder('[[IF').findAll();
    if (!matches || !matches.length) {
      break;
    }
    const markerCell = matches[matches.length - 1];
    const markerValue = markerCell.getValue();
    const match = markerValue.match(/^\[\[IF\s+([A-Za-z0-9_\-]+)\]\]$/);
    if (!match) {
      markerCell.clear();
      continue;
    }
    const flagName = match[1];
    const endRow = findMarkerRow_(sheet, markerCell.getRow() + 1, '[[ENDIF]]');
    if (!endRow) {
      throw new Error('Missing [[ENDIF]] marker for flag ' + flagName + ' in sheet ' + sheet.getName());
    }
    const include = toBool(context.flags[flagName]);
    const startRow = markerCell.getRow();
    if (!include) {
      sheet.deleteRows(startRow, endRow - startRow + 1);
    } else {
      sheet.deleteRow(endRow);
      sheet.deleteRow(startRow);
    }
  }
}

function replaceScalarPlaceholders_(sheet, placeholders, warnings) {
  const range = sheet.getDataRange();
  if (!range) {
    return;
  }
  const values = range.getValues();
  let needsUpdate = false;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const original = values[r][c];
      const rendered = replaceValuePlaceholders_(original, null, placeholders, warnings);
      if (rendered !== original) {
        needsUpdate = true;
        values[r][c] = rendered;
      }
    }
  }
  if (needsUpdate) {
    range.setValues(values);
  }
}

function replaceValuePlaceholders_(value, rowData, placeholders, warnings) {
  if (value === null || value === undefined) {
    return value;
  }
  if (typeof value !== 'string') {
    return value;
  }
  if (/^\[\[(ROW|ENDROW|IF|ENDIF)/.test(value)) {
    return value;
  }
  let output = value;
  output = output.replace(/\{([^}]+)\}/g, function (_, token) {
    const key = token.trim();
    if (rowData && Object.prototype.hasOwnProperty.call(rowData, key)) {
      return rowData[key];
    }
    if (placeholders && Object.prototype.hasOwnProperty.call(placeholders, key)) {
      return placeholders[key];
    }
    warnings.push('Missing value for {' + key + '}');
    return '';
  });
  output = output.replace(/\[([^\]]+)\]/g, function (_, token) {
    const key = token.trim();
    if (placeholders && Object.prototype.hasOwnProperty.call(placeholders, key)) {
      return placeholders[key];
    }
    warnings.push('Missing value for [' + key + ']');
    return '';
  });
  return output;
}

function findMarkerRow_(sheet, startRow, markerValue) {
  const lastRow = sheet.getLastRow();
  for (let r = startRow; r <= lastRow; r++) {
    const rowValues = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (let c = 0; c < rowValues.length; c++) {
      if (rowValues[c] === markerValue) {
        return r;
      }
    }
  }
  return null;
}
