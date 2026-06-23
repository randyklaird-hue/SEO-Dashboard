// ── Config.gs ──────────────────────────────────────────────────────────────
// Reads / writes user settings from a hidden "Config" sheet.

const CONFIG_SHEET = '_Config';
const SETTINGS = {
  SITE_URL:       'siteUrl',
  GA4_PROPERTY:   'ga4PropertyId',
  DATE_RANGE_DAYS: 'dateRangeDays',
};

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) sheet = _initConfigSheet(ss);

  const data = sheet.getDataRange().getValues();
  const cfg = {};
  data.forEach(([key, value]) => { if (key) cfg[key] = value; });
  return cfg;
}

function saveConfig(siteUrl, ga4PropertyId, dateRangeDays) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) sheet = _initConfigSheet(ss);

  const map = {
    [SETTINGS.SITE_URL]:        siteUrl,
    [SETTINGS.GA4_PROPERTY]:    ga4PropertyId,
    [SETTINGS.DATE_RANGE_DAYS]: dateRangeDays || 90,
  };

  sheet.clearContents();
  Object.entries(map).forEach(([k, v], i) => sheet.getRange(i + 1, 1, 1, 2).setValues([[k, v]]));
}

function _initConfigSheet(ss) {
  const sheet = ss.insertSheet(CONFIG_SHEET);
  sheet.hideSheet();
  saveConfig('', '', 90);
  return sheet;
}

function getDateRange(days) {
  const end = new Date();
  const start = new Date();
  start.setDate(end.getDate() - (days || 90));
  return {
    startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    endDate:   Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd'),
  };
}
