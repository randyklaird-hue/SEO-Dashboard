// ── Analytics.gs ───────────────────────────────────────────────────────────
// Wraps the GA4 Data API v1beta (properties/{id}:runReport).

const GA4_BASE = 'https://analyticsdata.googleapis.com/v1beta';

function fetchGA4Data(propertyId, startDate, endDate) {
  if (!propertyId) throw new Error('GA4: propertyId is required (e.g. "123456789").');

  // Strip "properties/" prefix if user pasted the full resource name
  const pid = propertyId.toString().replace(/^properties\//, '');
  const token = ScriptApp.getOAuthToken();
  const headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
  const endpoint = `${GA4_BASE}/properties/${pid}:runReport`;

  // ── 1. Sessions + users + conversions by channel group ───────────────────
  const channelReport = _ga4Fetch(endpoint, headers, {
    dateRanges: [{ startDate, endDate }],
    dimensions: [{ name: 'sessionDefaultChannelGrouping' }],
    metrics: [
      { name: 'sessions' },
      { name: 'totalUsers' },
      { name: 'conversions' },
      { name: 'bounceRate' },
      { name: 'averageSessionDuration' },
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 50,
  });

  const channels = _parseGA4Report(channelReport, ['channel'], ['sessions', 'users', 'conversions', 'bounceRate', 'avgDuration']);

  // ── 2. Top landing pages ──────────────────────────────────────────────────
  const landingReport = _ga4Fetch(endpoint, headers, {
    dateRanges: [{ startDate, endDate }],
    dimensions: [{ name: 'landingPage' }],
    metrics: [
      { name: 'sessions' },
      { name: 'totalUsers' },
      { name: 'conversions' },
      { name: 'bounceRate' },
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 200,
  });

  const landingPages = _parseGA4Report(landingReport, ['page'], ['sessions', 'users', 'conversions', 'bounceRate']);

  // ── 3. Sessions over time (daily) ─────────────────────────────────────────
  const trendReport = _ga4Fetch(endpoint, headers, {
    dateRanges: [{ startDate, endDate }],
    dimensions: [{ name: 'date' }],
    metrics: [
      { name: 'sessions' },
      { name: 'totalUsers' },
      { name: 'conversions' },
    ],
    orderBys: [{ dimension: { dimensionName: 'date' } }],
    limit: 365,
  });

  const trends = _parseGA4Report(trendReport, ['date'], ['sessions', 'users', 'conversions']);

  // ── 4. Top countries ──────────────────────────────────────────────────────
  const geoReport = _ga4Fetch(endpoint, headers, {
    dateRanges: [{ startDate, endDate }],
    dimensions: [{ name: 'country' }],
    metrics: [{ name: 'sessions' }, { name: 'totalUsers' }],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 20,
  });

  const countries = _parseGA4Report(geoReport, ['country'], ['sessions', 'users']);

  return { channels, landingPages, trends, countries };
}

function _ga4Fetch(endpoint, headers, body) {
  const options = {
    method: 'post',
    headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };
  const resp = UrlFetchApp.fetch(endpoint, options);
  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error(`GA4 API error ${code}: ${resp.getContentText().substring(0, 400)}`);
  }
  return JSON.parse(resp.getContentText());
}

function _parseGA4Report(report, dimLabels, metLabels) {
  if (!report.rows) return [];
  return report.rows.map(row => {
    const obj = {};
    (row.dimensionValues || []).forEach((d, i) => { obj[dimLabels[i] || `dim${i}`] = d.value; });
    (row.metricValues   || []).forEach((m, i) => { obj[metLabels[i]  || `met${i}`] = parseFloat(m.value) || 0; });
    return obj;
  });
}
