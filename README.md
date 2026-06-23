# SEO Dashboard — Google Apps Script

Pulls data from **Google Search Console** and **Google Analytics 4** and writes a formatted multi-tab report directly into Google Sheets. No API keys to manage — OAuth runs automatically through your Google account.

---

## Setup (5 minutes)

### 1. Create a Google Spreadsheet

Open [Google Sheets](https://sheets.new) and create a new spreadsheet.

### 2. Open the Apps Script editor

Click **Extensions → Apps Script**.

### 3. Copy the files

Delete the default `Code.gs` content. Then create the following files and paste in the contents from this repo:

| File | Description |
|------|-------------|
| `appsscript.json` | OAuth scopes manifest — enable via **Project Settings → Show appsscript.json** |
| `Code.gs` | Menu, config dialog, report runners |
| `Config.gs` | Reads/writes settings to a hidden sheet |
| `SearchConsole.gs` | Search Console API calls |
| `Analytics.gs` | GA4 Data API calls |
| `Report.gs` | Sheet writing and chart creation |

> **Tip:** You can also use [clasp](https://github.com/google/clasp) to push all files at once:
> ```bash
> npm install -g @google/clasp
> clasp login
> clasp create --type sheets --title "SEO Dashboard"
> clasp push
> ```

### 4. Enable the manifest

In the Apps Script editor, go to **Project Settings** (gear icon) and check **Show "appsscript.json" manifest file in editor**. Paste the `appsscript.json` contents there.

### 5. Save and reload

Press **Ctrl+S** (or Cmd+S), then reload the Google Sheets tab. A new **📊 SEO Report** menu will appear.

### 6. Configure

Click **📊 SEO Report → ⚙️ Configure…** and enter:

- **Search Console site URL** — use `sc-domain:example.com` for a domain property, or `https://example.com/` for a URL-prefix property. Must be verified in your Search Console account.
- **GA4 Property ID** — a number like `123456789`. Find it in Google Analytics → Admin → Property Settings → Property ID.
- **Date range** — 30, 60, 90 days, or 6 months.

### 7. Run the report

Click **📊 SEO Report → ▶ Run Full Report**. On the first run, Google will ask you to authorize the script — click through the permission screens.

---

## Output tabs

| Tab | Contents |
|-----|----------|
| 📊 Dashboard | KPI summary cards + top 10 keywords + channel breakdown |
| 🔍 Keywords | All keywords: clicks, impressions, CTR, avg. position |
| 📄 SC Pages | Top pages from Search Console |
| 📈 GA4 Channels | Sessions by channel group (organic, paid, social, etc.) |
| 🏠 Landing Pages | Top landing pages with sessions, users, conversions |
| 📅 Daily Trends | Daily sessions/users/conversions + line chart |
| 🌍 Countries | Top 20 countries by sessions |

---

## Scheduling

Click **📊 SEO Report → 🕒 Schedule Daily (6 AM)** to auto-refresh every morning. To remove, click **🗑 Remove Schedule**.

---

## Required permissions

The script requests these OAuth scopes (declared in `appsscript.json`):

| Scope | Why |
|-------|-----|
| `spreadsheets` | Read/write the report spreadsheet |
| `webmasters.readonly` | Read Search Console data |
| `analytics.readonly` | Read GA4 data |
| `script.external_request` | Make HTTPS calls to Google APIs via UrlFetchApp |

---

## Troubleshooting

**"Search Console API error 403"** — The Google account running the script must have access to the site in Search Console. Check [search.google.com/search-console](https://search.google.com/search-console).

**"GA4 API error 403"** — The account needs Viewer access to the GA4 property. Check Google Analytics → Admin → Property Access Management.

**"Missing configuration"** — Run **Configure…** from the menu and fill in both fields.

**No data rows returned** — Search Console only has data for verified properties with traffic. Make sure the site URL format matches exactly what's shown in your Search Console property list (domain vs. URL-prefix).
