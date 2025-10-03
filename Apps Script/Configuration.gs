let CONFIG_CACHE = null;
let PROPERTIES_CACHE = null;
let PROPERTIES_BY_NAME_CACHE = null;
let PROPERTIES_BY_KEYWORD_CACHE = null;

function getConfigMap() {
  if (CONFIG_CACHE) {
    return CONFIG_CACHE;
  }
  const sheet = getSheetByName_(SHEET_NAMES.CONFIG);
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const key = normalizeKey_(values[i][0]);
    if (!key) {
      continue;
    }
    map[key] = values[i][1];
  }
  CONFIG_CACHE = map;
  return map;
}

function getConfigValue(key, options) {
  options = options || {};
  const map = getConfigMap();
  const value = Object.prototype.hasOwnProperty.call(map, key) ? map[key] : undefined;
  if (options.required && (value === undefined || value === '')) {
    throw new Error('Missing configuration value for "' + key + '"');
  }
  if (value === undefined || value === '') {
    return options.defaultValue;
  }
  if (options.parse === 'boolean') {
    return toBool(value);
  }
  if (options.parse === 'number') {
    const num = toNumber(value, options.defaultValue);
    if (num === undefined || num === null || isNaN(num)) {
      throw new Error('Configuration value for "' + key + '" must be numeric.');
    }
    return num;
  }
  return value;
}

function getCreditsHeaderMapping() {
  const statusHeader = getConfigValue(CONFIG_KEYS.CREDITS_STATUS, { defaultValue: '' });
  const methodHeader = getConfigValue(CONFIG_KEYS.CREDITS_METHOD, { defaultValue: '' });
  const payerHeader = getConfigValue(CONFIG_KEYS.CREDITS_PAYER, { defaultValue: '' });
  return {
    date: getConfigValue(CONFIG_KEYS.CREDITS_DATE, { required: true }),
    amount: getConfigValue(CONFIG_KEYS.CREDITS_AMOUNT, { required: true }),
    property: getConfigValue(CONFIG_KEYS.CREDITS_PROPERTY, { required: true }),
    unit: getConfigValue(CONFIG_KEYS.CREDITS_UNIT, { required: true }),
    category: getConfigValue(CONFIG_KEYS.CREDITS_CATEGORY, { required: true }),
    subcategory: getConfigValue(CONFIG_KEYS.CREDITS_SUBCATEGORY, { required: true }),
    status: statusHeader || 'Payment status',
    method: methodHeader || 'Payment method',
    payer: payerHeader || 'Payer / Payee'
  };
}

function getFolderConfig() {
  return {
    outputFolderId: getConfigValue(CONFIG_KEYS.OUTPUT_FOLDER_ID, { required: true }),
    reportsFolderName: getConfigValue(CONFIG_KEYS.REPORTS_FOLDER_NAME, { defaultValue: '' }),
    exportsFolderName: getConfigValue(CONFIG_KEYS.EXPORTS_FOLDER_NAME, { defaultValue: '' })
  };
}

function getCreditsSourceConfig() {
  return {
    fileId: getConfigValue(CONFIG_KEYS.CREDITS_FILE_ID, { defaultValue: '' }),
    folderId: getConfigValue(CONFIG_KEYS.CREDITS_FOLDER_ID, { defaultValue: '' }),
    addSheet: getConfigValue(CONFIG_KEYS.ADD_CREDITS_SHEET, { parse: 'boolean', defaultValue: false })
  };
}

function clearConfigCacheForTesting() {
  CONFIG_CACHE = null;
}

function getPropertiesConfig() {
  if (PROPERTIES_CACHE) {
    return PROPERTIES_CACHE;
  }
  const sheet = getSheetByName_(SHEET_NAMES.PROPERTIES);
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error('Properties sheet is empty.');
  }
  const headerRow = values[0];
  const headerIndex = {};
  for (let i = 0; i < headerRow.length; i++) {
    const key = normalizeKey_(headerRow[i]);
    if (key) {
      headerIndex[key] = i;
    }
  }
  const requiredHeaders = [
    PROPERTY_COLUMNS.PROPERTY,
    PROPERTY_COLUMNS.MAF,
    PROPERTY_COLUMNS.MARKUP,
    PROPERTY_COLUMNS.AIRBNB_PERCENT,
    PROPERTY_COLUMNS.HAS_AIRBNB,
    PROPERTY_COLUMNS.ADMIN_FEE,
    PROPERTY_COLUMNS.ADMIN_FEE_ENABLED,
    PROPERTY_COLUMNS.KEYWORDS
  ];
  requiredHeaders.forEach(function (name) {
    if (!Object.prototype.hasOwnProperty.call(headerIndex, name)) {
      throw new Error('Missing column "' + name + '" in Properties sheet.');
    }
  });
  const properties = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const propertyName = row[headerIndex[PROPERTY_COLUMNS.PROPERTY]];
    if (!propertyName) {
      continue;
    }
    const property = {
      name: propertyName,
      maf: toNumber(row[headerIndex[PROPERTY_COLUMNS.MAF]], 0) || 0,
      markup: toNumber(row[headerIndex[PROPERTY_COLUMNS.MARKUP]], 0) || 0,
      airbnbPercent: toNumber(row[headerIndex[PROPERTY_COLUMNS.AIRBNB_PERCENT]], 0) || 0,
      hasAirbnb: toBool(row[headerIndex[PROPERTY_COLUMNS.HAS_AIRBNB]]),
      adminFee: toNumber(row[headerIndex[PROPERTY_COLUMNS.ADMIN_FEE]], 0) || 0,
      adminFeeEnabled: toBool(row[headerIndex[PROPERTY_COLUMNS.ADMIN_FEE_ENABLED]]),
      keywords: parseKeywords_(row[headerIndex[PROPERTY_COLUMNS.KEYWORDS]])
    };
    properties.push(property);
  }
  PROPERTIES_CACHE = properties;
  buildPropertyIndexes_(properties);
  return properties;
}

function getPropertiesByName() {
  if (PROPERTIES_BY_NAME_CACHE) {
    return PROPERTIES_BY_NAME_CACHE;
  }
  getPropertiesConfig();
  return PROPERTIES_BY_NAME_CACHE;
}

function getPropertyByName(name) {
  if (!name) {
    return null;
  }
  const map = getPropertiesByName();
  return map[name] || null;
}

function resolvePropertyName(rawProperty) {
  const normalized = normalizeStringValue_(rawProperty);
  if (!normalized) {
    return '';
  }
  const directMatch = getPropertyByName(normalized);
  if (directMatch) {
    return directMatch.name;
  }
  const properties = getPropertiesConfig();
  const haystack = normalized.toLowerCase();
  let candidate = null;
  properties.forEach(function (property) {
    if (!property.keywords || !property.keywords.length) {
      return;
    }
    const matchesAll = property.keywords.every(function (keyword) {
      return haystack.indexOf(keyword.toLowerCase()) !== -1;
    });
    if (matchesAll) {
      if (!candidate || property.keywords.length > candidate.keywords.length) {
        candidate = property;
      }
    }
  });
  if (candidate) {
    return candidate.name;
  }
  return normalized;
}

function buildPropertyIndexes_(properties) {
  const byName = {};
  const byKeyword = {};
  properties.forEach(function (property) {
    byName[property.name] = property;
    property.keywords.forEach(function (keyword) {
      const key = keyword.toLowerCase();
      if (!byKeyword[key]) {
        byKeyword[key] = [];
      }
      byKeyword[key].push(property);
    });
  });
  PROPERTIES_BY_NAME_CACHE = byName;
  PROPERTIES_BY_KEYWORD_CACHE = byKeyword;
}

function lookupPropertyByKeyword(keyword) {
  if (!keyword) {
    return [];
  }
  if (!PROPERTIES_BY_KEYWORD_CACHE) {
    getPropertiesConfig();
  }
  const key = keyword.toLowerCase();
  return PROPERTIES_BY_KEYWORD_CACHE[key] || [];
}

function parseKeywords_(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.filter(function (item) { return item; }).map(function (item) { return item.toString().trim(); });
  }
  return value.toString().split(',').map(function (item) { return item.trim(); }).filter(function (item) { return item; });
}

function clearPropertiesCacheForTesting() {
  PROPERTIES_CACHE = null;
  PROPERTIES_BY_NAME_CACHE = null;
  PROPERTIES_BY_KEYWORD_CACHE = null;
}
