const SHEET_NAMES = {
  CONFIG: 'Configuration',
  PROPERTIES: 'Properties',
  TRANSACTIONS: 'Transactions',
  STAGING: 'Entry & Edit',
  CONTROL: 'Entry Controls',
  IMPORT_LOG: 'Import Log',
  REPORT_LOG: 'Report Log',
  TEMPLATE_BODY: 'ReportBodyTemplate',
  TEMPLATE_TOTALS: 'ReportTotalsTemplate',
  TEMPLATE_AIRBNB: 'ReportAirbnbTemplate'
};

const CONFIG_KEYS = {
  CREDITS_FILE_ID: 'Credits Document',
  CREDITS_FOLDER_ID: 'Credits Folder ID',
  OUTPUT_FOLDER_ID: 'Output Folder ID',
  REPORTS_FOLDER_NAME: 'Reports Folder Name',
  EXPORTS_FOLDER_NAME: 'Exports Folder Name',
  CREDITS_DATE: 'Credits Date',
  CREDITS_AMOUNT: 'Credits Amount',
  CREDITS_PROPERTY: 'Credits Property',
  CREDITS_UNIT: 'Credits Unit',
  CREDITS_CATEGORY: 'Credits Category',
  CREDITS_SUBCATEGORY: 'Credits Subcategory',
  CREDITS_STATUS: 'Credits Status',
  CREDITS_METHOD: 'Credits Method',
  CREDITS_PAYER: 'Credits Payer',
  ADD_CREDITS_SHEET: 'Add Credits Sheet'
};

const PROPERTY_COLUMNS = {
  PROPERTY: 'Property',
  MAF: 'MAF',
  MARKUP: 'Markup',
  AIRBNB_PERCENT: 'Airbnb',
  HAS_AIRBNB: 'Has Airbnb',
  ADMIN_FEE: 'Admin Fee',
  ADMIN_FEE_ENABLED: 'Admin Fee Enabled',
  ORDER: 'Order',
  KEYWORDS: 'Key'
};

const TRANSACTION_HEADERS = [
  'Transaction ID',
  'Date',
  'Property',
  'Unit',
  'Credits',
  'Fees',
  'Debits',
  'Security Deposits',
  'Debit/Credit Explanation',
  'Markup Included',
  'Markup Revenue',
  'Internal Notes',
  'Deleted',
  'Deleted Timestamp'
];

const STAGING_HEADERS = TRANSACTION_HEADERS.concat(['Delete Permanently']);

const STAGING_CONTROL = {
  SHEET: SHEET_NAMES.CONTROL,
  START_DATE_CELL: 'B2',
  END_DATE_CELL: 'B3',
  PROPERTY_CELL: 'B4',
  REPORT_LABEL_CELL: 'B5',
  SHOW_DELETED_CELL: 'B6',
};

const TEMPLATE_TOKENS = {
  ROW_START: '\\[\\[ROW ',
  ROW_END: '\\[\\[ENDROW\\]\\]',
  IF_START: '\\[\\[IF ',
  IF_END: '\\[\\[ENDIF\\]\\]'
};

const VERSION_SUFFIX_SEPARATOR = '_';

const LOG_HEADERS = {
  IMPORT: ['Timestamp', 'File ID', 'File Name', 'Last Modified', 'Rows Imported', 'Notes'],
  REPORT: ['Timestamp', 'Report Label', 'Version', 'Start Date', 'End Date', 'Spreadsheet ID', 'Report URL', 'Properties Included']
};
