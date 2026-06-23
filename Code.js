// ── Code.gs ────────────────────────────────────────────────────────────────
// Entry point: menu, sidebar dialog, and main report orchestration.

// ── Licensing ──────────────────────────────────────────────────────────────
// Free tier: Search Console only, 30-day max date range.
// Pro tier:  Full GA4 data, 90/180-day ranges, daily scheduling.

function isPro() {
  try {
    const url = 'https://appsmarket.googleapis.com/appsmarket/v2/userLicense/' + _getAppId();
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) return false;
    const data = JSON.parse(resp.getContentText());
    return data.state === 'ACTIVE' || data.state === 'ACTIVE_TRIAL';
  } catch (e) {
    return false;
  }
}

function _getAppId() {
  // Replace with your actual Marketplace App ID once published.
  // Found in GCP Console → Marketplace SDK → App ID.
  return PropertiesService.getScriptProperties().getProperty('MARKETPLACE_APP_ID') || '';
}

function _requirePro(featureName) {
  if (!isPro()) {
    SpreadsheetApp.getUi().alert(
      '⭐ Pro Feature',
      `${featureName} is available on the Pro plan ($29/month).\n\nUpgrade at: https://randyklaird-hue.github.io/SEO-Dashboard/`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
  return true;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 SEO Report')
    .addItem('⚙️  Configure…',              'showConfigDialog')
    .addSeparator()
    .addItem('▶  Run Search Console Report', 'runSearchConsoleOnly')
    .addItem('📈 Run Full Report (Pro)',      'runReport')
    .addSeparator()
    .addItem('🕒 Schedule Daily — Pro',      'scheduleDailyReport')
    .addItem('🗑  Remove Schedule',           'removeSchedule')
    .addSeparator()
    .addItem('⭐ Upgrade to Pro',            'showUpgradeDialog')
    .addToUi();
}

// ── Configuration dialog ───────────────────────────────────────────────────

function showConfigDialog() {
  const cfg = getConfig();
  const pro = isPro();
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; font-size: 14px; color: #202124; }
    label { display: block; font-weight: 500; margin-top: 16px; margin-bottom: 4px; font-size: 13px; }
    input, select { width: 100%; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 14px; box-sizing: border-box; }
    input:focus { outline: none; border-color: #1a73e8; }
    input:disabled, select:disabled { background: #f8f9fa; color: #9aa0a6; cursor: not-allowed; }
    .hint { font-size: 11px; color: #5f6368; margin-top: 3px; }
    .pro-badge { display: inline-block; background: #fbbc04; color: #202124; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-left: 6px; vertical-align: middle; }
    .pro-lock { font-size: 11px; color: #ea4335; margin-top: 3px; }
    button { margin-top: 20px; padding: 10px 24px; background: #1a73e8; color: white; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; width: 100%; }
    button:hover { background: #1557b0; }
    .cancel { background: white; color: #1a73e8; border: 1px solid #dadce0; margin-top: 8px; }
    .cancel:hover { background: #e8f0fe; }
    .upgrade-bar { background: #fce8e6; border: 1px solid #f5c6c0; border-radius: 6px; padding: 10px 14px; margin-bottom: 16px; font-size: 13px; color: #c5221f; }
    .upgrade-bar a { color: #c5221f; font-weight: 600; }
  </style>
</head>
<body>
  <h2 style="margin-top:0;font-size:18px">SEO Report Configuration</h2>

  ${!pro ? `<div class="upgrade-bar">Free plan — Search Console only, 30-day max. <a href="https://randyklaird-hue.github.io/SEO-Dashboard/" target="_blank">Upgrade to Pro →</a></div>` : ''}

  <label for="siteUrl">Search Console Site URL</label>
  <input id="siteUrl" type="text" placeholder="sc-domain:example.com" value="${cfg[SETTINGS.SITE_URL] || ''}">
  <div class="hint">Use <code>sc-domain:example.com</code> for domain property or <code>https://example.com/</code> for URL prefix.</div>

  <label for="ga4Property">GA4 Property ID <span class="pro-badge">PRO</span></label>
  <input id="ga4Property" type="text" placeholder="123456789" value="${cfg[SETTINGS.GA4_PROPERTY] || ''}" ${!pro ? 'disabled' : ''}>
  ${!pro ? '<div class="pro-lock">🔒 Upgrade to Pro to enable GA4 data</div>' : '<div class="hint">Find it in Google Analytics → Admin → Property Settings.</div>'}

  <label for="days">Date Range ${!pro ? '<span class="pro-badge">30-DAY MAX</span>' : ''}</label>
  <select id="days" ${!pro ? 'disabled' : ''}>
    <option value="30"  ${(cfg[SETTINGS.DATE_RANGE_DAYS] == 30 || !pro)  ? 'selected' : ''}>Last 30 days</option>
    ${pro ? `
    <option value="60"  ${cfg[SETTINGS.DATE_RANGE_DAYS] == 60  ? 'selected' : ''}>Last 60 days</option>
    <option value="90"  ${(!cfg[SETTINGS.DATE_RANGE_DAYS] || cfg[SETTINGS.DATE_RANGE_DAYS] == 90) ? 'selected' : ''}>Last 90 days</option>
    <option value="180" ${cfg[SETTINGS.DATE_RANGE_DAYS] == 180 ? 'selected' : ''}>Last 6 months</option>` : ''}
  </select>

  <button onclick="save()">Save &amp; Close</button>
  <button class="cancel" onclick="google.script.host.close()">Cancel</button>

  <script>
    const isPro = ${pro};
    function save() {
      const siteUrl = document.getElementById('siteUrl').value.trim();
      const ga4Prop = isPro ? document.getElementById('ga4Property').value.trim() : '';
      const days    = isPro ? parseInt(document.getElementById('days').value) : 30;
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
  .setHeight(pro ? 420 : 460)
  .setTitle('Configure SEO Report');

  SpreadsheetApp.getUi().showModalDialog(html, 'Configure SEO Report');
}

function showUpgradeDialog() {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Google Sans', Arial, sans-serif; padding: 24px; font-size: 14px; color: #202124; text-align: center; }
    h2 { font-size: 20px; margin-bottom: 8px; }
    .price { font-size: 36px; font-weight: 700; color: #1a73e8; margin: 16px 0 4px; }
    .price span { font-size: 16px; font-weight: 400; color: #5f6368; }
    ul { text-align: left; display: inline-block; margin: 16px 0; padding-left: 20px; color: #3c4043; }
    ul li { margin-bottom: 8px; }
    a.btn { display: inline-block; background: #1a73e8; color: white; padding: 12px 32px; border-radius: 6px; font-size: 15px; font-weight: 600; text-decoration: none; margin-top: 8px; }
    a.btn:hover { background: #1557b0; }
    .close { margin-top: 12px; font-size: 13px; color: #5f6368; cursor: pointer; }
  </style>
</head>
<body>
  <div style="font-size:32px">⭐</div>
  <h2>Upgrade to Pro</h2>
  <div class="price">$29<span>/month</span></div>
  <ul>
    <li>✅ Google Analytics 4 data</li>
    <li>✅ 90-day &amp; 6-month date ranges</li>
    <li>✅ Daily auto-refresh scheduling</li>
    <li>✅ Landing pages &amp; countries tabs</li>
    <li>✅ Daily trends chart</li>
  </ul>
  <a class="btn" href="https://randyklaird-hue.github.io/SEO-Dashboard/" target="_blank">Upgrade on Marketplace</a>
  <div class="close" onclick="google.script.host.close()">Maybe later</div>
</body>
</html>
  `)
  .setWidth(340)
  .setHeight(380)
  .setTitle('Upgrade to Pro');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upgrade to Pro');
}

// ── Report runners ─────────────────────────────────────────────────────────

function runReport() {
  if (!_requirePro('Full Report (GA4 + all tabs)')) return;

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
  // Free tier: cap at 30 days
  const days = isPro() ? (parseInt(cfg[SETTINGS.DATE_RANGE_DAYS]) || 90) : 30;
  const dateRange = getDateRange(days);
  _toast('Fetching Search Console data…');
  const scData = fetchSearchConsoleData(cfg[SETTINGS.SITE_URL], dateRange.startDate, dateRange.endDate);
  writeSearchConsoleSheet(SpreadsheetApp.getActiveSpreadsheet(), scData);
  _toast('✅ Search Console data updated!', 5);
}

// ── Scheduler ─────────────────────────────────────────────────────────────

function scheduleDailyReport() {
  if (!_requirePro('Daily scheduling')) return;
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
