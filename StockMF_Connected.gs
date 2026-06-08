// ================================================================
//  MARKET ORACLE — Connected Apps Script
//  Google Sheets Backend + Web App JSON Endpoint
//  Jitu · Abhijit Kumar Misra
//
//  SETUP STEPS:
//  1. Extensions → Apps Script → paste this file → Save
//  2. Run setupAll() once
//  3. Deploy → New Deployment → Web App
//     · Execute as: Me
//     · Who has access: Anyone
//  4. Copy the Web App URL
//  5. Paste it into index.html where it says WEB_APP_URL
// ================================================================

const CFG = {
  SHEET_STOCKS  : "Stocks",
  SHEET_MF      : "Mutual Funds",
  SHEET_CAPITAL : "Capital Tracker",
  SHEET_LOG     : "Decision Log",
  ALERT_EMAIL   : Session.getActiveUser().getEmail(),
  CHECK_MINUTES : 60,
  BUY_THRESHOLD : 70,
};

// ================================================================
//  WEB APP ENDPOINT — This is what your index.html calls
//  Returns all live data as JSON in one request
// ================================================================
function doGet(e) {
  try {
    const params   = (e && e.parameter) ? e.parameter : {};
    const action   = params.action   || "";
    const callback = params.callback || "";   // JSONP callback name

    // Force-refresh sheet data when requested
    if (action === "refresh") {
      refreshAll();
    }

    const payload = buildPayload_();
    const json    = JSON.stringify(payload);

    // ── JSONP mode (browser fetch from GitHub Pages) ──────────
    // Browser passes ?callback=__jp0 — we wrap JSON in a function
    // call so it bypasses CORS entirely via a <script> tag load.
    if (callback) {
      return ContentService
        .createTextOutput(callback + "(" + json + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    // ── Plain JSON mode (direct browser tab / Postman test) ───
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    const params   = (e && e.parameter) ? e.parameter : {};
    const callback = params.callback || "";
    const errJson  = JSON.stringify({ error: true, message: err.message, timestamp: new Date().toISOString() });
    if (callback) {
      return ContentService
        .createTextOutput(callback + "(" + errJson + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(errJson)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── BUILD JSON PAYLOAD ───────────────────────────────────────────
function buildPayload_() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const stocks  = readSheet_(ss, CFG.SHEET_STOCKS,  13); // 13 columns
  const mfs     = readSheet_(ss, CFG.SHEET_MF,      14); // 14 columns
  const capital = readCapital_(ss);
  const log     = readLog_(ss);

  return {
    timestamp : new Date().toISOString(),
    capital,
    stocks    : stocks.map(rowToStock_),
    mfs       : mfs.map(rowToMF_),
    log       : log.slice(0, 50), // last 50 entries
  };
}

function readSheet_(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues()
    .filter(r => r[0]); // skip empty rows
}

function rowToStock_(r) {
  return {
    symbol  : r[0],  name    : r[1],  sector  : r[2],
    price   : r[3],  high52  : r[4],  low52   : r[5],
    pe      : r[6],  target  : r[7],  score   : r[8],
    signal  : r[9],  alloc   : r[10], status  : r[11],
    updated : r[12] ? new Date(r[12]).toISOString() : null,
  };
}

function rowToMF_(r) {
  return {
    code    : r[0],  name    : r[1],  category: r[2],
    nav     : r[3],  low1y   : r[4],  high1y  : r[5],
    cagr1y  : r[6],  cagr3y  : r[7],  target  : r[8],
    score   : r[9],  signal  : r[10], alloc   : r[11],
    status  : r[12], updated : r[13] ? new Date(r[13]).toISOString() : null,
  };
}

function readCapital_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_CAPITAL);
  if (!sh) return {};
  return {
    total    : sh.getRange("B3").getValue() || 100000,
    stockPct : sh.getRange("B4").getValue() || 60,
    mfPct    : sh.getRange("B5").getValue() || 40,
    deployed : sh.getRange("B13").getValue() || 0,
    remaining: sh.getRange("B14").getValue() || 0,
  };
}

function readLog_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
    .filter(r => r[0])
    .reverse(); // most recent first
  return rows.map(r => ({
    time: r[0] ? new Date(r[0]).toISOString() : '',
    type: r[1], code: r[2], name: r[3],
    price: r[4], score: r[5], signal: r[6],
    action: r[7], email: r[8],
  }));
}

// ================================================================
//  SETUP
// ================================================================
function setupAll() {
  createAllSheets_();
  installTrigger_();
  SpreadsheetApp.getUi().alert(
    "✅ Market Oracle — Connected Setup Complete!\n\n" +
    "Sheets created. Trigger installed.\n\n" +
    "NEXT STEP — Deploy as Web App:\n" +
    "1. Click Deploy → New Deployment\n" +
    "2. Type: Web App\n" +
    "3. Execute as: Me\n" +
    "4. Who has access: Anyone\n" +
    "5. Copy the Web App URL\n" +
    "6. Paste into index.html at WEB_APP_URL\n\n" +
    "Then fill your watchlist and run 'Refresh All Now'."
  );
}

function createAllSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildStocksSheet_(ss);
  buildMFSheet_(ss);
  buildCapitalSheet_(ss);
  buildLogSheet_(ss);
}

function buildStocksSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SHEET_STOCKS) || ss.insertSheet(CFG.SHEET_STOCKS);
  if (sh.getLastRow() > 0) return;
  sh.appendRow(["Symbol","Company Name","Sector","Price ₹","52W High","52W Low","P/E Ratio","Target ₹","Buy Score","Signal","Allocate ₹","Status","Updated"]);
  // Jitu's stocks & ETFs
  sh.appendRow(["NSE:SILVERBEES","Nippon India Silver ETF","Silver ETF","","","","","115","","","0","Watching",""]);
  sh.appendRow(["NSE:HDFCSILVER","HDFC Silver ETF","Silver ETF","","","","","70","","","0","Watching",""]);
  sh.appendRow(["NSE:ICICIB22","ICICI Pru Bharat 22 ETF","Govt PSU ETF","","","","","88","","","0","Watching",""]);
  sh.appendRow(["NSE:SUZLON","Suzlon Energy Ltd","Renewable Energy","","","","","48","","","0","Watching",""]);
  sh.appendRow(["NSE:TRIDENT","Trident Ltd","Textiles","","","","","30","","","0","Watching",""]);
  sh.appendRow(["NSE:RELIANCE","Reliance Industries","Energy","","","","","2600","","","0","Watching",""]);
  sh.appendRow(["NSE:HDFCBANK","HDFC Bank","Banking","","","","","1450","","","0","Watching",""]);
  styleHeader_(sh, "#0a1628");
  sh.setColumnWidth(2, 220); sh.setColumnWidth(10, 130);
}

function buildMFSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SHEET_MF) || ss.insertSheet(CFG.SHEET_MF);
  if (sh.getLastRow() > 0) return;
  sh.appendRow(["AMFI Code","Scheme Name","Category","NAV ₹","1Y Low","1Y High","1Y CAGR%","3Y CAGR%","Target NAV","Buy Score","Signal","Allocate ₹","Status","Updated"]);
  // Jitu's 9 funds
  sh.appendRow(["120403","Invesco India Mid Cap Fund — Direct Growth","Mid Cap","","","","","","200","","","0","Watching",""]);
  sh.appendRow(["147622","Motilal Oswal Multicap Fund — Direct Growth","Multicap","","","","","","11","","","0","Watching",""]);
  sh.appendRow(["122639","Parag Parikh Flexi Cap Fund — Direct Growth","Flexi Cap","","","","","","75","","","0","Watching",""]);
  sh.appendRow(["118825","Nippon India Large Cap Fund — Direct Growth","Large Cap","","","","","","88","","","0","Watching",""]);
  sh.appendRow(["119019","SBI Gold Direct Plan Growth","Gold","","","","","","25","","","0","Watching",""]);
  sh.appendRow(["153137","JioBlackRock Flexi Cap Fund — Direct Growth","Flexi Cap","","","","","","9.2","","","0","Watching",""]);
  sh.appendRow(["130503","SBI Multi Asset Allocation Fund — Direct Growth","Multi Asset","","","","","","48","","","0","Watching",""]);
  sh.appendRow(["149240","Bandhan Small Cap Fund — Direct Growth","Small Cap","","","","","","34","","","0","Watching",""]);
  sh.appendRow(["119773","DSP Natural Resources & New Energy Fund — Direct","Sectoral","","","","","","78","","","0","Watching",""]);
  styleHeader_(sh, "#0d2818");
  sh.setColumnWidth(2, 260);
}

function buildCapitalSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SHEET_CAPITAL) || ss.insertSheet(CFG.SHEET_CAPITAL);
  if (sh.getLastRow() > 0) return;
  sh.appendRow(["CAPITAL TRACKER"]);
  sh.getRange("A1").setFontSize(16).setFontWeight("bold").setFontColor("#c9a84c");
  sh.appendRow([""]);
  sh.appendRow(["Total Available Capital (₹)", 100000]);
  sh.appendRow(["Allocated to Stocks (%)", 60]);
  sh.appendRow(["Allocated to Mutual Funds (%)", 40]);
  sh.appendRow([""]);
  sh.appendRow(["--- Auto Calculated ---"]);
  sh.appendRow(["Stock Budget (₹)", "=B3*(B4/100)"]);
  sh.appendRow(["MF Budget (₹)", "=B3*(B5/100)"]);
  sh.appendRow([""]);
  sh.appendRow(["Deployed in Stocks (₹)", '=IFERROR(SUM(INDIRECT("Stocks!K2:K1000")),0)']);
  sh.appendRow(["Deployed in MFs (₹)", '=IFERROR(SUM(INDIRECT("\'Mutual Funds\'!L2:L1000")),0)']);
  sh.appendRow(["Total Deployed (₹)", "=B11+B12"]);
  sh.appendRow(["Remaining Capital (₹)", "=B3-B13"]);
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 180);
  sh.getRange("A3:A14").setFontWeight("bold");
  sh.getRange("B3:B14").setNumberFormat("₹#,##0.00");
}

function buildLogSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SHEET_LOG) || ss.insertSheet(CFG.SHEET_LOG);
  if (sh.getLastRow() > 0) return;
  sh.appendRow(["Timestamp","Type","Symbol/Code","Name","Price/NAV","Buy Score","Signal","Action","Email Sent"]);
  styleHeader_(sh, "#1a0a2e");
}

function styleHeader_(sh, bg) {
  const r = sh.getRange(1, 1, 1, sh.getMaxColumns());
  r.setBackground(bg).setFontColor("#f0d080").setFontWeight("bold").setFontSize(10);
  sh.setFrozenRows(1); sh.setRowHeight(1, 32);
}

// ================================================================
//  TRIGGER
// ================================================================
function installTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "refreshAll")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("refreshAll").timeBased().everyMinutes(CFG.CHECK_MINUTES).create();
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "refreshAll")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ================================================================
//  MAIN REFRESH ENGINE
// ================================================================
function refreshAll() {
  refreshStocks_();
  refreshMutualFunds_();
  updateCapitalSplit_();
}

// ── STOCKS ──────────────────────────────────────────────────────
function refreshStocks_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_STOCKS);
  if (!sh || sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const symbol = String(row[0]).trim();
    if (!symbol) continue;

    const price  = gfValue_(symbol, "price");
    const high52 = gfValue_(symbol, "52weekhigh");
    const low52  = gfValue_(symbol, "52weeklow");
    const pe     = gfValue_(symbol, "pe");
    if (price === null) continue;

    const target = parseFloat(row[7]) || 0;
    const score  = calcStockScore_(price, high52, low52, pe, target);
    const signal = scoreToSignal_(score);

    sh.getRange(i+1, 4).setValue(price);
    sh.getRange(i+1, 5).setValue(high52 || "");
    sh.getRange(i+1, 6).setValue(low52  || "");
    sh.getRange(i+1, 7).setValue(pe     || "");
    sh.getRange(i+1, 9).setValue(score);
    sh.getRange(i+1,10).setValue(signal);
    sh.getRange(i+1,13).setValue(new Date());
    colorScore_(sh, i+1, 9, score);
    colorSignal_(sh, i+1, 10, signal);

    if (score >= CFG.BUY_THRESHOLD && row[11] !== "Alerted") {
      sendAlert_("Stock", symbol, row[1], price, score, signal);
      logEntry_("Stock", symbol, row[1], price, score, signal, "AUTO ALERT");
      sh.getRange(i+1, 12).setValue("Alerted");
    }
  }
}

// ── MUTUAL FUNDS ─────────────────────────────────────────────────
function refreshMutualFunds_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_MF);
  if (!sh || sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const code = String(row[0]).trim();
    if (!code) continue;

    const mf = fetchMFData_(code);
    if (!mf) continue;

    const { nav, low1y, high1y, cagr1y, cagr3y } = mf;
    const target = parseFloat(row[8]) || 0;
    const score  = calcMFScore_(nav, low1y, high1y, cagr1y, cagr3y, target);
    const signal = scoreToSignal_(score);

    sh.getRange(i+1, 4).setValue(nav);
    sh.getRange(i+1, 5).setValue(low1y  || "");
    sh.getRange(i+1, 6).setValue(high1y || "");
    sh.getRange(i+1, 7).setValue(cagr1y !== null ? cagr1y.toFixed(2) : "");
    sh.getRange(i+1, 8).setValue(cagr3y !== null ? cagr3y.toFixed(2) : "");
    sh.getRange(i+1,10).setValue(score);
    sh.getRange(i+1,11).setValue(signal);
    sh.getRange(i+1,14).setValue(new Date());
    colorScore_(sh, i+1, 10, score);
    colorSignal_(sh, i+1, 11, signal);

    if (score >= CFG.BUY_THRESHOLD && row[12] !== "Alerted") {
      sendAlert_("Mutual Fund", code, row[1], nav, score, signal);
      logEntry_("Mutual Fund", code, row[1], nav, score, signal, "AUTO ALERT");
      sh.getRange(i+1, 13).setValue("Alerted");
    }
  }
}

// ── GOOGLEFINANCE HELPER ─────────────────────────────────────────
function gfValue_(symbol, attr) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let   tmp = ss.getSheetByName("__tmp__") || ss.insertSheet("__tmp__");
    tmp.hideSheet();
    const cell = tmp.getRange("A1");
    cell.setFormula('=GOOGLEFINANCE("' + symbol + '","' + attr + '")');
    SpreadsheetApp.flush();
    Utilities.sleep(1800);
    const val = cell.getValue();
    cell.clearContent();
    return (typeof val === "number" && val > 0) ? val : null;
  } catch(e) { return null; }
}

// ── AMFI NAV FETCH ───────────────────────────────────────────────
function fetchMFData_(code) {
  try {
    const resp = UrlFetchApp.fetch("https://api.mfapi.in/mf/" + code, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const json = JSON.parse(resp.getContentText());
    const data = json.data;
    if (!data || data.length < 2) return null;

    const currentNav = parseFloat(data[0].nav);
    const now        = new Date();
    const yr1        = new Date(); yr1.setFullYear(now.getFullYear() - 1);
    const yr3        = new Date(); yr3.setFullYear(now.getFullYear() - 3);

    let navs1y = [], nav1yAgo = null, nav3yAgo = null;
    for (const d of data) {
      const dt  = parseAMFIDate_(d.date);
      const val = parseFloat(d.nav);
      if (!dt || isNaN(val)) continue;
      if (dt >= yr1) navs1y.push(val);
      if (!nav1yAgo && dt <= yr1) nav1yAgo = val;
      if (!nav3yAgo && dt <= yr3) nav3yAgo = val;
    }

    return {
      nav    : currentNav,
      low1y  : navs1y.length ? Math.min(...navs1y) : null,
      high1y : navs1y.length ? Math.max(...navs1y) : null,
      cagr1y : nav1yAgo ? ((currentNav / nav1yAgo) - 1) * 100 : null,
      cagr3y : nav3yAgo ? (Math.pow(currentNav / nav3yAgo, 1/3) - 1) * 100 : null,
    };
  } catch(e) {
    Logger.log("MF fetch error [" + code + "]: " + e.message);
    return null;
  }
}

function parseAMFIDate_(str) {
  try {
    const p = str.split("-");
    return new Date(p[2] + "-" + p[1] + "-" + p[0]);
  } catch(e) { return null; }
}

// ── SCORING ──────────────────────────────────────────────────────
function calcStockScore_(price, high52, low52, pe, target) {
  let s = 0;
  if (high52 && low52 && high52 > low52)
    s += Math.round((1 - (price - low52) / (high52 - low52)) * 40);
  if (high52 && high52 > 0)
    s += Math.min(35, Math.round(((high52 - price) / high52 * 100) * 0.875));
  s += (pe && pe > 0) ? Math.min(25, Math.max(0, Math.round(25 - ((pe-10)/30)*25))) : 12;
  if (target > 0 && price <= target) s = Math.min(100, s + 10);
  return Math.min(100, Math.max(0, s));
}

function calcMFScore_(nav, low1y, high1y, cagr1y, cagr3y, target) {
  let s = 0;
  if (high1y && low1y && high1y > low1y)
    s += Math.round((1 - (nav - low1y) / (high1y - low1y)) * 35);
  if (cagr1y) s += Math.min(35, Math.max(0, Math.round(cagr1y * 1.75)));
  if (cagr3y) s += Math.min(20, Math.max(0, Math.round(cagr3y * 1.33)));
  if (target > 0 && nav <= target) s = Math.min(100, s + 10);
  return Math.min(100, Math.max(0, s));
}

function scoreToSignal_(score) {
  if (score >= 80) return "⚡ STRONG BUY";
  if (score >= 65) return "✅ BUY";
  if (score >= 45) return "👁 WATCH";
  return "⛔ AVOID";
}

function colorScore_(sh, row, col, score) {
  const cell = sh.getRange(row, col);
  if (score >= 80)      cell.setBackground("#00c853").setFontColor("#000");
  else if (score >= 65) cell.setBackground("#64dd17").setFontColor("#000");
  else if (score >= 45) cell.setBackground("#ffd600").setFontColor("#000");
  else                  cell.setBackground("#d50000").setFontColor("#fff");
}

function colorSignal_(sh, row, col, signal) {
  const cell = sh.getRange(row, col);
  if (signal.includes("STRONG")) cell.setBackground("#003300").setFontColor("#00ff88").setFontWeight("bold");
  else if (signal.includes("BUY")) cell.setBackground("#0a3320").setFontColor("#4ade80").setFontWeight("bold");
  else if (signal.includes("WATCH")) cell.setBackground("#2d2000").setFontColor("#fbbf24").setFontWeight("bold");
  else cell.setBackground("#2d0000").setFontColor("#f87171").setFontWeight("bold");
}

// ── CAPITAL SPLIT ─────────────────────────────────────────────────
function updateCapitalSplit_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cap = ss.getSheetByName(CFG.SHEET_CAPITAL);
  if (!cap) return;
  const total    = parseFloat(cap.getRange("B3").getValue()) || 0;
  const stockPct = parseFloat(cap.getRange("B4").getValue()) || 60;
  const mfPct    = parseFloat(cap.getRange("B5").getValue()) || 40;
  allocateBudget_(CFG.SHEET_STOCKS, total * stockPct / 100, 9,  11, 10);
  allocateBudget_(CFG.SHEET_MF,     total * mfPct   / 100, 10, 12, 11);
}

function allocateBudget_(sheetName, budget, scoreCol, allocCol, signalCol) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getValues();
  let total  = 0;
  const eligible = [];

  for (let i = 1; i < data.length; i++) {
    const score  = parseFloat(data[i][scoreCol - 1]) || 0;
    const signal = String(data[i][signalCol - 1]);
    if (score >= 65 && signal.includes("BUY")) {
      total += score; eligible.push({ row: i + 1, score });
    }
  }
  for (let i = 2; i <= data.length; i++) sh.getRange(i, allocCol).setValue(0);
  if (!total) return;
  eligible.forEach(e => sh.getRange(e.row, allocCol).setValue(Math.round((e.score / total) * budget)));
}

// ── EMAIL ALERT ───────────────────────────────────────────────────
function sendAlert_(type, symbol, name, price, score, signal) {
  const subject = `🔔 Market Oracle Alert: ${signal} — ${name} (Score: ${score}/100)`;
  const body = `Market Oracle — Investment Signal\n\n` +
    `Type     : ${type}\nSymbol   : ${symbol}\nName     : ${name}\n` +
    `Price    : ₹${typeof price === "number" ? price.toFixed(2) : price}\n` +
    `Score    : ${score}/100\nSignal   : ${signal}\nTime     : ${new Date().toLocaleString("en-IN")}\n\n` +
    `Review your Capital Tracker before placing an order.\n\n— Market Oracle · Abhijit Kumar Misra`;
  GmailApp.sendEmail(CFG.ALERT_EMAIL, subject, body);
}

// ── DECISION LOG ──────────────────────────────────────────────────
function logEntry_(type, code, name, price, score, signal, action) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.SHEET_LOG);
  if (!log) return;
  log.appendRow([new Date(), type, code, name,
    typeof price === "number" ? price.toFixed(2) : price,
    score, signal, action, "✅ " + CFG.ALERT_EMAIL]);
}

// ── MANUAL ACTIONS ────────────────────────────────────────────────
function markInvested() {
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("Select a data row first."); return; }
  const data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const ui   = SpreadsheetApp.getUi();
  if (ui.alert("Mark Invested", "Record investment in: " + (data[1]||data[0]) + "?", ui.ButtonSet.YES_NO) === ui.Button.YES) {
    sh.getRange(row, sh.getLastColumn() - 1).setValue("🟢 Invested");
    logEntry_(sh.getName().includes("Stock") ? "Stock" : "MF",
      data[0], data[1]||data[0], data[3], data[8]||data[9], "—", "INVESTED");
    ui.alert("✅ Recorded.");
  }
}

function resetAlertStatus() {
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row < 2) return;
  sh.getRange(row, sh.getLastColumn() - 1).setValue("Watching")
    .setBackground("#ffffff").setFontColor("#000000");
}

// ── MENU ──────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📈 Market Oracle")
    .addItem("▶ Refresh All Now",              "refreshAll")
    .addSeparator()
    .addItem("🏦 Update Capital Split",         "updateCapitalSplit_")
    .addItem("✅ Mark Row as Invested",         "markInvested")
    .addItem("🔄 Reset Row Alert Status",       "resetAlertStatus")
    .addSeparator()
    .addItem("⚙ Setup (first time only)",      "setupAll")
    .addItem("🗑 Remove Trigger",               "removeTrigger")
    .addToUi();
}
