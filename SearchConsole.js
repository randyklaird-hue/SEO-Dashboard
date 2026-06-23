// ── SearchConsole.gs ───────────────────────────────────────────────────────
// Wraps the Search Console searchAnalytics/query REST endpoint.

const SC_BASE = 'https://www.googleapis.com/webmasters/v3';

function fetchSearchConsoleData(siteUrl, startDate, endDate) {
  if (!siteUrl) throw new Error('Search Console: siteUrl is required.');

  const token = ScriptApp.getOAuthToken();
  const headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };

  // Encode the site URL for the path (e.g. "sc-domain:example.com" stays as-is)
  const encodedSite = encodeURIComponent(siteUrl);

  const rows = [];

  // ── 1. Query performance by keyword ──────────────────────────────────────
  const queryBody = {
    startDate,
    endDate,
    dimensions: ['query'],
    rowLimit: 1000,
    dataState: 'final',
  };

  const queryResp = _scFetch(`/sites/${encodedSite}/searchAnalytics/query`, headers, queryBody);
  if (queryResp.rows) {
    queryResp.rows.forEach(r => {
      rows.push({
        type: 'query',
        keyword:     r.keys[0],
        clicks:      r.clicks,
        impressions: r.impressions,
        ctr:         r.ctr,
        position:    r.position,
      });
    });
  }

  // ── 2. Query performance by page ─────────────────────────────────────────
  const pageBody = {
    startDate,
    endDate,
    dimensions: ['page'],
    rowLimit: 500,
    dataState: 'final',
  };

  const pageResp = _scFetch(`/sites/${encodedSite}/searchAnalytics/query`, headers, pageBody);
  const pages = [];
  if (pageResp.rows) {
    pageResp.rows.forEach(r => {
      pages.push({
        page:        r.keys[0],
        clicks:      r.clicks,
        impressions: r.impressions,
        ctr:         r.ctr,
        position:    r.position,
      });
    });
  }

  // ── 3. Device breakdown ───────────────────────────────────────────────────
  const deviceBody = {
    startDate,
    endDate,
    dimensions: ['device'],
    rowLimit: 10,
  };
  const deviceResp = _scFetch(`/sites/${encodedSite}/searchAnalytics/query`, headers, deviceBody);
  const devices = deviceResp.rows || [];

  return { keywords: rows, pages, devices };
}

function _scFetch(path, headers, body) {
  const options = {
    method: 'post',
    headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };
  const resp = UrlFetchApp.fetch(SC_BASE + path, options);
  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error(`Search Console API error ${code}: ${resp.getContentText().substring(0, 400)}`);
  }
  return JSON.parse(resp.getContentText());
}
