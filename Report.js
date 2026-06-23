// ── Report.gs ──────────────────────────────────────────────────────────────
// Writes data to sheets and builds charts.

const COLORS = {
  header:    '#1a73e8',
  headerText:'#ffffff',
  accent:    '#e8f0fe',
  green:     '#34a853',
  yellow:    '#fbbc04',
  red:       '#ea4335',
  gray:      '#f8f9fa',
};

// ── Sheet helpers ──────────────────────────────────────────────────────────

function getOrCreateSheet(ss, name, position) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name, position);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function styleHeader(sheet, row, cols) {
  const range = sheet.getRange(row, 1, 1, cols);
  range.setBackground(COLORS.header)
       .setFontColor(COLORS.headerText)
       .setFontWeight('bold')
       .setFontSize(10);
}

function autoResizeCols(sheet, numCols) {
  for (let c = 1; c <= numCols; c++) sheet.autoResizeColumn(c);
}

function pctFormat(sheet, startRow, col, numRows) {
  sheet.getRange(startRow, col, numRows, 1).setNumberFormat('0.00%');
}

function numFormat(sheet, startRow, col, numRows, fmt) {
  sheet.getRange(startRow, col, numRows, 1).setNumberFormat(fmt || '#,##0');
}

// ── Dashboard ──────────────────────────────────────────────────────────────

function writeDashboard(ss, scData, ga4Data, cfg, dateRange) {
  const sheet = getOrCreateSheet(ss, '📊 Dashboard', 0);
  sheet.setTabColor(COLORS.header);

  const { siteUrl, ga4PropertyId, dateRangeDays } = cfg;

  // Title block
  sheet.getRange('A1').setValue('SEO & Analytics Dashboard').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A2').setValue(`Site: ${siteUrl || '(not set)'}  |  GA4: ${ga4PropertyId || '(not set)'}  |  ${dateRange.startDate} → ${dateRange.endDate} (${dateRangeDays} days)`).setFontColor('#5f6368');
  sheet.getRange('A3').setValue(`Generated: ${new Date().toLocaleString()}`).setFontColor('#9aa0a6').setFontSize(9);

  // ── KPI summary cards (row 5) ─────────────────────────────────────────────
  const totalClicks      = scData.keywords.reduce((s, r) => s + r.clicks, 0);
  const totalImpressions = scData.keywords.reduce((s, r) => s + r.impressions, 0);
  const avgPosition      = scData.keywords.length
    ? scData.keywords.reduce((s, r) => s + r.position, 0) / scData.keywords.length
    : 0;
  const totalSessions    = ga4Data.channels.reduce((s, r) => s + r.sessions, 0);
  const totalUsers       = ga4Data.channels.reduce((s, r) => s + r.users, 0);
  const totalConversions = ga4Data.channels.reduce((s, r) => s + r.conversions, 0);

  const kpis = [
    ['Total Clicks',       totalClicks,      '#,##0'],
    ['Impressions',        totalImpressions, '#,##0'],
    ['Avg. Position',      avgPosition,      '0.0'],
    ['Overall CTR',        totalClicks / (totalImpressions || 1), '0.00%'],
    ['Sessions',           totalSessions,    '#,##0'],
    ['Users',              totalUsers,       '#,##0'],
    ['Conversions',        totalConversions, '#,##0'],
  ];

  sheet.getRange(5, 1, 1, kpis.length * 2)
       .setBackground(COLORS.accent)
       .setFontWeight('bold');

  kpis.forEach(([label, value, fmt], i) => {
    const col = i * 2 + 1;
    sheet.getRange(5, col).setValue(label).setFontColor('#5f6368').setFontSize(9);
    sheet.getRange(6, col).setValue(value).setNumberFormat(fmt).setFontSize(14).setFontWeight('bold');
    sheet.setColumnWidth(col,     120);
    sheet.setColumnWidth(col + 1, 12);
  });

  // ── Top 10 keywords table ─────────────────────────────────────────────────
  const kwRow = 9;
  sheet.getRange(kwRow, 1).setValue('Top Keywords by Clicks').setFontWeight('bold').setFontSize(11);

  const kwHeaders = ['Keyword', 'Clicks', 'Impressions', 'CTR', 'Avg Position'];
  sheet.getRange(kwRow + 1, 1, 1, kwHeaders.length).setValues([kwHeaders]);
  styleHeader(sheet, kwRow + 1, kwHeaders.length);

  const top10 = scData.keywords.slice(0, 10);
  if (top10.length) {
    const kwData = top10.map(r => [r.keyword, r.clicks, r.impressions, r.ctr, r.position]);
    sheet.getRange(kwRow + 2, 1, kwData.length, 5).setValues(kwData);
    pctFormat(sheet, kwRow + 2, 4, kwData.length);
    numFormat(sheet, kwRow + 2, 5, kwData.length, '0.0');
    sheet.getRange(kwRow + 2, 1, kwData.length, 5)
         .setBackground(COLORS.gray);
  }

  // ── Top 10 channels table ─────────────────────────────────────────────────
  const chRow = kwRow + 14;
  sheet.getRange(chRow, 1).setValue('Traffic Channels').setFontWeight('bold').setFontSize(11);

  const chHeaders = ['Channel', 'Sessions', 'Users', 'Conversions', 'Bounce Rate', 'Avg. Duration (s)'];
  sheet.getRange(chRow + 1, 1, 1, chHeaders.length).setValues([chHeaders]);
  styleHeader(sheet, chRow + 1, chHeaders.length);

  const topCh = ga4Data.channels.slice(0, 10);
  if (topCh.length) {
    const chData = topCh.map(r => [r.channel, r.sessions, r.users, r.conversions, r.bounceRate, r.avgDuration]);
    sheet.getRange(chRow + 2, 1, chData.length, 6).setValues(chData);
    pctFormat(sheet, chRow + 2, 5, chData.length);
    numFormat(sheet, chRow + 2, 6, chData.length, '0');
    sheet.getRange(chRow + 2, 1, chData.length, 6).setBackground(COLORS.gray);
  }

  autoResizeCols(sheet, 14);
  sheet.setFrozenRows(1);
}

// ── Search Console sheet ───────────────────────────────────────────────────

function writeSearchConsoleSheet(ss, scData) {
  // ── Keywords tab ──────────────────────────────────────────────────────────
  const kwSheet = getOrCreateSheet(ss, '🔍 Keywords', 1);
  kwSheet.setTabColor('#34a853');

  const kwHeaders = ['Keyword', 'Clicks', 'Impressions', 'CTR', 'Avg Position'];
  kwSheet.getRange(1, 1, 1, kwHeaders.length).setValues([kwHeaders]);
  styleHeader(kwSheet, 1, kwHeaders.length);

  if (scData.keywords.length) {
    const kwData = scData.keywords.map(r => [r.keyword, r.clicks, r.impressions, r.ctr, r.position]);
    kwSheet.getRange(2, 1, kwData.length, 5).setValues(kwData);
    pctFormat(kwSheet, 2, 4, kwData.length);
    numFormat(kwSheet, 2, 5, kwData.length, '0.0');
  }

  kwSheet.setFrozenRows(1);
  autoResizeCols(kwSheet, 5);

  // ── Pages tab ─────────────────────────────────────────────────────────────
  const pgSheet = getOrCreateSheet(ss, '📄 SC Pages', 2);
  pgSheet.setTabColor('#34a853');

  const pgHeaders = ['Page URL', 'Clicks', 'Impressions', 'CTR', 'Avg Position'];
  pgSheet.getRange(1, 1, 1, pgHeaders.length).setValues([pgHeaders]);
  styleHeader(pgSheet, 1, pgHeaders.length);

  if (scData.pages.length) {
    const pgData = scData.pages.map(r => [r.page, r.clicks, r.impressions, r.ctr, r.position]);
    pgSheet.getRange(2, 1, pgData.length, 5).setValues(pgData);
    pctFormat(pgSheet, 2, 4, pgData.length);
    numFormat(pgSheet, 2, 5, pgData.length, '0.0');
  }

  pgSheet.setFrozenRows(1);
  autoResizeCols(pgSheet, 5);
}

// ── GA4 sheets ─────────────────────────────────────────────────────────────

function writeGA4Sheets(ss, ga4Data) {
  // ── Channels ──────────────────────────────────────────────────────────────
  const chSheet = getOrCreateSheet(ss, '📈 GA4 Channels', 3);
  chSheet.setTabColor('#fbbc04');

  const chHeaders = ['Channel', 'Sessions', 'Users', 'Conversions', 'Bounce Rate', 'Avg. Duration (s)'];
  chSheet.getRange(1, 1, 1, chHeaders.length).setValues([chHeaders]);
  styleHeader(chSheet, 1, chHeaders.length);

  if (ga4Data.channels.length) {
    const chData = ga4Data.channels.map(r => [r.channel, r.sessions, r.users, r.conversions, r.bounceRate, r.avgDuration]);
    chSheet.getRange(2, 1, chData.length, 6).setValues(chData);
    pctFormat(chSheet, 2, 5, chData.length);
    numFormat(chSheet, 2, 6, chData.length, '0');
  }

  chSheet.setFrozenRows(1);
  autoResizeCols(chSheet, 6);

  // ── Landing pages ─────────────────────────────────────────────────────────
  const lpSheet = getOrCreateSheet(ss, '🏠 Landing Pages', 4);
  lpSheet.setTabColor('#fbbc04');

  const lpHeaders = ['Landing Page', 'Sessions', 'Users', 'Conversions', 'Bounce Rate'];
  lpSheet.getRange(1, 1, 1, lpHeaders.length).setValues([lpHeaders]);
  styleHeader(lpSheet, 1, lpHeaders.length);

  if (ga4Data.landingPages.length) {
    const lpData = ga4Data.landingPages.map(r => [r.page, r.sessions, r.users, r.conversions, r.bounceRate]);
    lpSheet.getRange(2, 1, lpData.length, 5).setValues(lpData);
    pctFormat(lpSheet, 2, 5, lpData.length);
  }

  lpSheet.setFrozenRows(1);
  autoResizeCols(lpSheet, 5);

  // ── Trends (daily) ────────────────────────────────────────────────────────
  const trSheet = getOrCreateSheet(ss, '📅 Daily Trends', 5);
  trSheet.setTabColor('#fbbc04');

  const trHeaders = ['Date', 'Sessions', 'Users', 'Conversions'];
  trSheet.getRange(1, 1, 1, trHeaders.length).setValues([trHeaders]);
  styleHeader(trSheet, 1, trHeaders.length);

  if (ga4Data.trends.length) {
    const trData = ga4Data.trends.map(r => [r.date, r.sessions, r.users, r.conversions]);
    trSheet.getRange(2, 1, trData.length, 4).setValues(trData);

    // Add a line chart for sessions over time
    _addTrendChart(trSheet, trData.length);
  }

  trSheet.setFrozenRows(1);
  autoResizeCols(trSheet, 4);

  // ── Countries ─────────────────────────────────────────────────────────────
  const geoSheet = getOrCreateSheet(ss, '🌍 Countries', 6);
  geoSheet.setTabColor('#ea4335');

  const geoHeaders = ['Country', 'Sessions', 'Users'];
  geoSheet.getRange(1, 1, 1, geoHeaders.length).setValues([geoHeaders]);
  styleHeader(geoSheet, 1, geoHeaders.length);

  if (ga4Data.countries.length) {
    const geoData = ga4Data.countries.map(r => [r.country, r.sessions, r.users]);
    geoSheet.getRange(2, 1, geoData.length, 3).setValues(geoData);
  }

  geoSheet.setFrozenRows(1);
  autoResizeCols(geoSheet, 3);
}

function _addTrendChart(sheet, dataRows) {
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, dataRows + 1, 2))  // Date + Sessions
    .setPosition(5, 6, 0, 0)
    .setOption('title', 'Sessions Over Time')
    .setOption('legend', { position: 'bottom' })
    .setOption('hAxis', { title: 'Date' })
    .setOption('vAxis', { title: 'Sessions' })
    .setOption('width', 600)
    .setOption('height', 300);

  sheet.insertChart(chartBuilder.build());
}
