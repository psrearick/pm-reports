function getActiveSpreadsheet_() {
  return SpreadsheetApp.getActive();
}

function getSheetByName_(name) {
  const sheet = getActiveSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet not found: ' + name);
  }
  return sheet;
}

function normalizeStringValue_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value.toString().trim();
}

function toBool(value) {
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true' || normalized === 'yes' || normalized === '1') {
      return true;
    }
    if (normalized === 'false' || normalized === 'no' || normalized === '0' || normalized === '') {
      return false;
    }
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  return false;
}

function toNumber(value, defaultValue) {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'boolean') {
    return value ? 1 : 0;
  }
  if (typeof value === 'string') {
    const num = Number(value.replace(/[^0-9.\-]/g, ''));
    if (!isNaN(num)) {
      return num;
    }
  }
  return defaultValue;
}

function parseCurrency_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string') {
    const cleaned = value.replace(/[^0-9.\-]/g, '');
    const parsed = Number(cleaned);
    if (!isNaN(parsed)) {
      return parsed;
    }
  }
  return 0;
}

function isBlank_(value) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function normalizeKey_(value) {
  return (value || '').toString().trim();
}

function utcNow_() {
  return new Date();
}

function formatDateForLog_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatDateDisplay_(date) {
  if (!date) {
    return '';
  }
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  if (isNaN(date.getTime())) {
    return '';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function formatCurrency_(value) {
  const number = typeof value === 'number' ? value : parseCurrency_(value);
  if (!number) {
    return '$0.00';
  }
  return '$' + number.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function ensureArray_(value) {
  if (Array.isArray(value)) {
    return value;
  }
  if (value === undefined || value === null) {
    return [];
  }
  return [value];
}

function sanitizeSheetName_(name) {
  const cleaned = (name || '').toString().replace(/[\\\[\]\?\*\/]/g, ' ').replace(/\s+/g, ' ').trim();
  return cleaned || 'Sheet';
}

function truncateName_(name, maxLength) {
  if (!name) {
    return '';
  }
  if (name.length <= maxLength) {
    return name;
  }
  return name.substring(0, maxLength);
}

function normalize(str) {
  return (str || "")
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, "")
    .trim()
    .replace(/\s+/, " ");
}
