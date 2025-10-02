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
  replaceScalarPlaceholders_(sheet, context.placeholders, context.flags, warnings);
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
    const match = markerValue.match(/^\[\[ROW\s+([A-Za-z0-9_\-]+)(.*?)\]\]$/);
    if (!match) {
      markerCell.clear();
      continue;
    }
    const blockName = match[1];
    const attributes = parseRowAttributes_(match[2]);
    renderRowBlockInstance_(sheet, markerCell, blockName, context, warnings, attributes);
  }
}

function renderRowBlockInstance_(sheet, markerCell, blockName, context, warnings, attributes) {
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
  let blockData = context.rows[blockName] || [];
  blockData = transformBlockData_(blockData, attributes || {});
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
      const renderedValues = renderTemplateBlockValues_(templateValues, blockData[i], context.placeholders, context.flags, warnings);
      targetRange.setValues(renderedValues);
    }
  }
  const newEndRow = findMarkerRow_(sheet, startRow + 1, '[[ENDROW]]');
  if (newEndRow) {
    sheet.deleteRow(newEndRow);
  }
  sheet.deleteRow(startRow);
}

function renderTemplateBlockValues_(templateValues, rowData, placeholders, flags, warnings) {
  const output = [];
  for (let r = 0; r < templateValues.length; r++) {
    const row = templateValues[r];
    const renderedRow = [];
    for (let c = 0; c < row.length; c++) {
      renderedRow.push(replaceValuePlaceholders_(row[c], rowData, placeholders, flags, warnings));
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

function replaceScalarPlaceholders_(sheet, placeholders, flags, warnings) {
  const range = sheet.getDataRange();
  if (!range) {
    return;
  }
  const values = range.getValues();
  let needsUpdate = false;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const original = values[r][c];
      const rendered = replaceValuePlaceholders_(original, null, placeholders, flags, warnings);
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

function replaceValuePlaceholders_(value, rowData, placeholders, flags, warnings) {
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
    const parsed = parsePlaceholderToken_(token);
    if (parsed.condition) {
      const flagValue = flags && Object.prototype.hasOwnProperty.call(flags, parsed.condition) ? flags[parsed.condition] : false;
      if (!toBool(flagValue)) {
        if (parsed.isBlock) {
          return '';
        }
        return '';
      }
    }
    const key = parsed.key;
    if (rowData && Object.prototype.hasOwnProperty.call(rowData, key)) {
      return rowData[key];
    }
    if (placeholders && Object.prototype.hasOwnProperty.call(placeholders, key)) {
      return placeholders[key];
    }
    warnings.push('Missing value for [' + key + ']');
    return '';
  });
  if (parsedAttachmentRegex_.test(output)) {
    output = processInlineConditionalBlocks_(output, flags);
  }
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

function parseRowAttributes_(attributeString) {
  const attributes = {};
  if (!attributeString) {
    return attributes;
  }
  const parts = attributeString.trim().split(/\s+/);
  parts.forEach(function (part) {
    if (!part) {
      return;
    }
    const match = part.match(/^([A-Za-z0-9_\-]+)=(.+)$/);
    if (match) {
      attributes[match[1]] = match[2];
    }
  });
  return attributes;
}

function transformBlockData_(blockData, attributes) {
  let data = Array.isArray(blockData) ? blockData.slice() : [];
  if (!data.length) {
    return data;
  }
  data = sortBlockData_(data, attributes);
  data = applyGroupingForDisplay_(data, attributes.group);
  return data;
}

function sortBlockData_(data, attributes) {
  const sortOrder = [];
  if (attributes.group) {
    sortOrder.push({ field: attributes.group, secondary: false });
  }
  if (attributes.sort) {
    sortOrder.push({ field: attributes.sort, secondary: true });
  }
  if (!sortOrder.length) {
    return data.slice();
  }
  return data.slice().sort(function (a, b) {
    for (let i = 0; i < sortOrder.length; i++) {
      const descriptor = sortOrder[i];
      const field = descriptor.field;
      const valueA = getComparableValue_(a[field]);
      const valueB = getComparableValue_(b[field]);
      if (valueA < valueB) {
        return -1;
      }
      if (valueA > valueB) {
        return 1;
      }
    }
    return 0;
  });
}

function applyGroupingForDisplay_(data, groupField) {
  let lastGroupKey = null;
  return data.map(function (item) {
    const clone = Object.assign({}, item);
    const rawValue = item[groupField];
    const groupKey = rawValue === null || rawValue === undefined ? '' : rawValue.toString();
    if (lastGroupKey !== null && groupKey === lastGroupKey) {
      clone[groupField] = '';
    } else {
      lastGroupKey = groupKey;
    }
    return clone;
  });
}

function getComparableValue_(value) {
  if (value instanceof Date) {
    return value.getTime();
  }
  if (typeof value === 'number') {
    return value;
  }
  if (value === null || value === undefined) {
    return Number.NEGATIVE_INFINITY;
  }
  const parsedDate = new Date(value);
  if (!isNaN(parsedDate.getTime())) {
    return parsedDate.getTime();
  }
  const numeric = Number(value);
  if (!isNaN(numeric)) {
    return numeric;
  }
  return value.toString().toLowerCase();
}

function parsePlaceholderToken_(token) {
  const blockMatch = token.match(/^\s*\/if\s*$/);
  if (blockMatch) {
    return { key: '', condition: null, isBlockEnd: true, isBlock: false };
  }
  const inlineBlockMatch = token.match(/^\s*if\s*=\s*([A-Za-z0-9_\-]+)\s*$/);
  if (inlineBlockMatch) {
    return { key: '', condition: inlineBlockMatch[1], isBlockStart: true, isBlock: true };
  }
  const inlineContentMatch = token.match(/^\s*(.+?)\s*\|\s*if\s*=\s*([A-Za-z0-9_\-]+)\s*$/);
  if (inlineContentMatch) {
    return { key: inlineContentMatch[1].trim(), condition: inlineContentMatch[2].trim(), isBlock: true, inlineContent: inlineContentMatch[1].trim() };
  }
  const match = token.match(/^\s*(.+?)(?:\s+if\s*=\s*([A-Za-z0-9_\-]+))?\s*$/);
  if (!match) {
    return { key: token.trim(), condition: null, isBlock: false };
  }
  return {
    key: match[1].trim(),
    condition: match[2] ? match[2].trim() : null,
    isBlock: false
  };
}

function processInlineConditionalBlocks_(text, flags) {
  let output = text;
  output = output.replace(/\[if=([A-Za-z0-9_\-]+)\]\s*([\s\S]*?)\s*\[\/if\]/g, function (_, condition, content) {
    const flagValue = flags && Object.prototype.hasOwnProperty.call(flags, condition) ? flags[condition] : false;
    return toBool(flagValue) ? content : '';
  });
  output = output.replace(/\[([^\]]+)\|\s*if=([A-Za-z0-9_\-]+)\]/g, function (_, content, condition) {
    const flagValue = flags && Object.prototype.hasOwnProperty.call(flags, condition) ? flags[condition] : false;
    return toBool(flagValue) ? content : '';
  });
  return output;
}

const parsedAttachmentRegex_ = /\[if=|\|\s*if=/;
