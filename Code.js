// ── Code.gs ────────────────────────────────────────────────────────────────
// Entry point: menu, sidebar dialog, and main report orchestration.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 SEO Report')
    .addItem('⚙️  Configure…',           'showConfigDialog')
    .addSeparator()
    .addItem('▶  Run Full Report',        'runReport')
    .addItem('🔍 Search Console Only',    'runSearchConsoleOnly')
    .addItem('📈 Google Analytics Only',  'runAnalyticsOnly')
    .addSeparator()
    .addItem('🕒 Schedule Daily (6 AM)',  'scheduleDailyReport')
    .addItem('🗑  Remove Schedule',        'removeSchedule')
    .addToUi();
}

// ── Configuration dialog ───────────────────────────────────────────────────

function showConfigDialog() {
  const cfg = getConfig();
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; font-size: 14px; color: #202124; }
    label { display: block; font-weight: 500; margin-top: 16px; margin-bottom: 4px; font-size: 13px; }
    input, select { width: 100%; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 14px; box-sizing: border-box; }
    input:focus { outline: none; border-color: #1a73e8; }
    .hint { font-size: 11px; color: #5f6368; margin-top: 3px; }
    button { margin-top: 20px; padding: 10px 24px; background: #1a73e8; color: white; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; width: 100%; }
    button:hover { background: #1557b0; }
    .cancel { background: white; color: #1a73e8; border: 1px solid #dadce0; margin-top: 8px; }
    .cancel:hover { background: #e8f0fe; }
  </style>
</head>
<body>
  <h2 style="margin-top:0;font-size:18px">SEO Report Configuration</h2>

  <label for="siteUrl">Search Console Site URL</label>
  <input id="siteUrl" type="text" placeholder="sc-domain:example.com" value="${cfg[SETTINGS.SITE_URL] || ''}">
  <div class="hint">Use <code>sc-domain:example.com</code> for domain property or <code>https://example.com/</code> for URL prefix.</div>

  <label for="ga4Property">GA4 Property ID</label>
  <input id="ga4Property" type="text" placeholder="123456789" value="${cfg[SETTINGS.GA4_PROPERTY] || ''}">
  <div class="hint">Find it in Google Analytics → Admin → Property Settings.</div>

  <label for="days">Date Range</label>
  <select id="days">
    <option value="30"  ${(cfg[SETTINGS.DATE_RANGE_DAYS] == 30)  ? 'selected' : ''}>Last 30 days</option>
    <option value="60"  ${(cfg[SETTINGS.DATE_RANGE_DAYS] == 60)  ? 'selected' : ''}>Last 60 days</option>
    <option value="90"  ${(!cfg[SETTINGS.DATE_RANGE_DAYS] || cfg[SETTINGS.DATE_RANGE_DAYS] == 90) ? 'selected' : ''}>Last 90 days</option>
    <option value="180" ${(cfg[SETTINGS.DATE_RANGE_DAYS] == 180) ? 'selected' : ''}>Last 6 months</option>
  </select>

  <button onclick="save()">Save &amp; Close</button>
  <button class="cancel" onclick="google.script.host.close()">Cancel</button>

  <script>
    function save() {
      const siteUrl  = document.getElementById('siteUrl').value.trim();
      const ga4Prop  = document.getElementById('ga4Property').value.trim();
      const days     = parseInt(document.getElementById('days').value);
      google.script.run
        .withSuccessHandler(() => google.script.host.close())
        .withFailureHandler(e => alert('Error: ' + e.message))
        .saveConfig(siteUrl, ga4Prop, days);
    }
  </script>
</body>
</html>
  `)
  .setWidth(420)
  .setHeight(420)
  .setTitle('Configure SEO Report');

  SpreadsheetApp.getUi().showModalDialog(html, 'Configure SEO Report');
}

// ── Report runners ─────────────────────────────────────────────────────────

function runReport() {
  const cfg = getConfig();
  _validateConfig(cfg);

  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dateRange = getDateRange(parseInt(cfg[SETTINGS.DATE_RANGE_DAYS]) || 90);

    _toast('Fetching Search Console data…');
    const scData = fetchSearchConsoleData(cfg[SETTINGS.SITE_URL], dateRange.startDate, dateRange.endDate);

    _toast('Fetching GA4 data…');
    const ga4Data = fetchGA4Data(cfg[SETTINGS.GA4_PROPERTY], dateRange.startDate, dateRange.endDate);

    _toast('Writing report…');
    writeDashboard(ss, scData, ga4Data, cfg, dateRange);
    writeSearchConsoleSheet(ss, scData);
    writeGA4Sheets(ss, ga4Data);

    // Switch to the dashboard tab
    ss.getSheetByName('📊 Dashboard').activate();

    _toast('✅ Report complete!', 5);
  } catch (e) {
    ui.alert('Report Error', e.message, ui.ButtonSet.OK);
    Logger.log(e.stack || e.message);
  }
}

function runSearchConsoleOnly() {
  const cfg = getConfig();
  if (!cfg[SETTINGS.SITE_URL]) {
    SpreadsheetApp.getUi().alert('Please configure the Search Console site URL first (SEO Report → Configure…).');
    return;
  }
  const dateRange = getDateRange(parseInt(cfg[SETTINGS.DATE_RANGE_DAYS]) || 90);
  _toast('Fetching Search Console data…');
  const scData = fetchSearchConsoleData(cfg[SETTINGS.SITE_URL], dateRange.startDate, dateRange.endDate);
  writeSearchConsoleSheet(SpreadsheetApp.getActiveSpreadsheet(), scData);
  _toast('✅ Search Console data updated!', 5);
}

function runAnalyticsOnly() {
  const cfg = getConfig();
  if (!cfg[SETTINGS.GA4_PROPERTY]) {
    SpreadsheetApp.getUi().alert('Please configure the GA4 Property ID first (SEO Report → Configure…).');
    return;
  }
  const dateRange = getDateRange(parseInt(cfg[SETTINGS.DATE_RANGE_DAYS]) || 90);
  _toast('Fetching GA4 data…');
  const ga4Data = fetchGA4Data(cfg[SETTINGS.GA4_PROPERTY], dateRange.startDate, dateRange.endDate);
  writeGA4Sheets(SpreadsheetApp.getActiveSpreadsheet(), ga4Data);
  _toast('✅ GA4 data updated!', 5);
}

// ── Scheduler ─────────────────────────────────────────────────────────────

function scheduleDailyReport() {
  // Remove existing triggers first to avoid duplicates
  removeSchedule();
  ScriptApp.newTrigger('runReport')
           .timeBased()
           .everyDays(1)
           .atHour(6)
           .create();
  SpreadsheetApp.getUi().alert('Scheduled! The report will refresh daily at ~6 AM in the script timezone.');
}

function removeSchedule() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ── Helpers ────────────────────────────────────────────────────────────────

function _validateConfig(cfg) {
  const missing = [];
  if (!cfg[SETTINGS.SITE_URL])      missing.push('Search Console site URL');
  if (!cfg[SETTINGS.GA4_PROPERTY])  missing.push('GA4 Property ID');
  if (missing.length) {
    throw new Error(`Missing configuration: ${missing.join(', ')}.\n\nOpen SEO Report → Configure… to set these values.`);
  }
}

function _toast(msg, timeout) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'SEO Report', timeout || 3);
}
