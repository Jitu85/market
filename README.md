# MARKET ORACLE

## Google Sheets setup

1. Create these tabs in the bound Google Sheet: `Recommendations` (required), `Allocation`, `Alerts`, and `Logs`.
2. Put field names in row 1. Recommended `Recommendations` headers:

   `id, name, symbol, type, schemeCode, price, move, score, signal, rsi, macd, pe, debt, technical, fundamental, risk`

3. Copy `google-apps-script.gs` into **Extensions > Apps Script**.
4. Deploy it as a **Web app**, execute as yourself, and grant access to the site's intended users.
5. Paste the deployment URL into `config.js` as `googleSheetsEndpoint`.

The dashboard loads the Sheet on startup and whenever Refresh is clicked. If the endpoint is missing or unavailable, it keeps the last usable/sample data and shows the sync error in the Scheduler panel.

## Live price and NAV rules

- Stock rows use `type=stock` and an exact NSE `symbol` such as `RELIANCE`.
- Mutual-fund rows use `type=fund` and must include the exact numeric AMFI `schemeCode`.
- Stock prices use Google Finance NSE formulas and may be delayed by up to 20 minutes. Run `setupMarketData` once in Apps Script after adding the headers.
- AMFI NAV is the latest officially published end-of-day NAV and is cached for six hours.
- If an official source fails, the dashboard keeps the Sheet value and reports `fallback-sheet-value` instead of silently presenting it as live.

### Current fund mappings

The two original sample funds are treated as Regular Growth plans unless a different `schemeCode` is entered:

- Parag Parikh Flexi Cap Fund - Regular Plan - Growth: `122640`
- UTI Nifty 50 Index Fund - Regular Plan - Growth Option: `100822`

Set an explicit scheme code to use a Direct or IDCW variant.

