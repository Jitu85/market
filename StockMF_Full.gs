// ================================================================
//  INVESTMENT DECISION DASHBOARD — Google Apps Script Backend
//  Jitu's Stock & Mutual Fund Buy Signal Engine
//  
//  SETUP: Extensions → Apps Script → Paste → Run setupAll() once
// ================================================================

// ── CONFIGURATION ───────────────────────────────────────────────
const CFG = {
  SHEET_STOCKS   : "Stocks",
  SHEET_MF       : "Mutual Funds",
  SHEET_CAPITAL  : "Capital Tracker",
  SHEET_LOG      : "Decision Log",
  ALERT_EMAIL    : Session.getActiveUser().getEmail(),
  CHECK_MINUTES  : 60,          // trigger interval
  BUY_SCORE_THRESHOLD : 70,     // email alert threshold
};

// ── COLUMN MAP: STOCKS ──────────────────────────────────────────
// A: Symbol | B: Name | C: Sector | D: Current Price | E: 52W High
// F: 52W Low | G: P/E Ratio | H: Target Price | I: Buy Score
// J: Signal | K: Capital Allocated ₹ | L: Status | M: Last Updated

// ── COLUMN MAP: MUTUAL FUNDS ────────────────────────────────────
// A: AMFI Code | B: Scheme Name | C: Category | D: Current NAV
// E: 1Y Low NAV | F: 1Y High NAV | G: 1Y CAGR% | H: 3Y CAGR%
// I: Target NAV | J: Buy Score | K: Signal | L: Capital Allocated ₹
// M: Status | N: Last Updated

// ================================================================
//  MASTER SETUP
// ================================================================
function setupAll() {
  createAllSheets_();
  installTrigger_();
  createMenu_();
  SpreadsheetApp.getUi().alert(
    "✅ Investment Dashboard Ready!\n\n" +
    "Sheets created:\n" +
    "  • Stocks\n  • Mutual Funds\n  • Capital Tracker\n  • Decision Log\n\n" +
    "Trigger: Every " + CFG.CHECK_MINUTES + " minutes\n" +
    "Alerts → " + CFG.ALERT_EMAIL + "\n\n" +
    "Next: Fill your watchlist and run 'Refresh All Now' from the menu."
  );
}

// ================================================================
//  SHEET CREATION
// ================================================================
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
  const headers = ["Symbol","Company Name","Sector","Price ₹","52W High","52W Low",
                   "P/E Ratio","Target ₹","Buy Score","Signal","Allocate ₹","Status","Updated"];
  sh.appendRow(headers);
  // Sample data
  sh.appendRow(["NSE:RELIANCE","Reliance Industries","Energy","","","","","2700","","","0","Watching",""]);
  sh.appendRow(["NSE:INFY","Infosys Ltd","IT","","","","","1350","","","0","Watching",""]);
  sh.appendRow(["NSE:HDFCBANK","HDFC Bank","Banking","","","","","1500","","","0","Watching",""]);
  sh.appendRow(["NSE:TCS","Tata Consultancy","IT","","","","","3500","","","0","Watching",""]);
  styleHeader_(sh, "#0a1628");
  sh.setColumnWidth(2, 200); sh.setColumnWidth(10, 120);
}

function buildMFSheet_(ss) {
  let sh = ss.getSheetByName(CFG.SHEET_MF) || ss.insertSheet(CFG.SHEET_MF);
  if (sh.getLastRow() > 0) return;
  const headers = ["AMFI Code","Scheme Name","Category","NAV ₹","1Y Low","1Y High",
                   "1Y CAGR%","3Y CAGR%","Target NAV","Buy Score","Signal","Allocate ₹","Status","Updated"];
  sh.appendRow(headers);
  sh.appendRow(["120503","Mirae Asset Large Cap - Direct","Large Cap","","","","","","100","","","0","Watching",""]);
  sh.appendRow(["100033","HDFC Top 100 - Direct","Large Cap","","","","","","850","","","0","Watching",""]);
  sh.appendRow(["119598","Parag Parikh Flexi Cap - Direct","Flexi Cap","","","","","","55","","","0","Watching",""]);
  styleHeader_(sh, "#0d2818");
  sh.setColumnWidth(2, 230);
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
  sh.appendRow(["MF Budget (₹)",    "=B3*(B5/100)"]);
  sh.appendRow([""]);
  sh.appendRow(["Deployed in Stocks (₹)",   '=IFERROR(SUM(INDIRECT("Stocks!K2:K1000")),0)']);
  sh.appendRow(["Deployed in MFs (₹)",      '=IFERROR(SUM(INDIRECT("\'Mutual Funds\'!L2:L1000")),0)']);
  sh.appendRow(["Total Deployed (₹)",       "=B11+B12"]);
  sh.appendRow(["Remaining Capital (₹)",    "=B3-B13"]);
  sh.setColumnWidth(1, 250); sh.setColumnWidth(2, 180);
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
  const n = sh.getMaxColumns();
  const r = sh.getRange(1, 1, 1, n);
  r.setBackground(bg).setFontColor("#f0d080").setFontWeight("bold").setFontSize(10);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
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

// ── STOCKS ENGINE ────────────────────────────────────────────────
function refreshStocks_() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(CFG.SHEET_STOCKS);
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const symbol = String(row[0]).trim();
    if (!symbol) continue;

    // Fetch via GoogleFinance helper
    const price   = gfValue_(symbol, "price");
    const high52  = gfValue_(symbol, "52weekhigh");
    const low52   = gfValue_(symbol, "52weeklow");
    const pe      = gfValue_(symbol, "pe");
    if (price === null) continue;

    const target  = parseFloat(row[7]) || 0;
    const score   = calcStockScore_(price, high52, low52, pe, target);
    const signal  = scoreToSignal_(score);

    sh.getRange(i+1, 4).setValue(price);
    sh.getRange(i+1, 5).setValue(high52 || "");
    sh.getRange(i+1, 6).setValue(low52  || "");
    sh.getRange(i+1, 7).setValue(pe     || "");
    sh.getRange(i+1, 9).setValue(score);
    sh.getRange(i+1,10).setValue(signal);
    sh.getRange(i+1,13).setValue(new Date());

    colorSignalCell_(sh, i+1, 10, signal);
    colorScoreCell_(sh, i+1, 9, score);

    if (score >= CFG.BUY_SCORE_THRESHOLD && row[11] !== "Alerted") {
      sendBuyAlert_("Stock", symbol, row[1], price, score, signal);
      logDecision_("Stock", symbol, row[1], price, score, signal, "AUTO ALERT");
      sh.getRange(i+1, 12).setValue("Alerted");
    }
  }
}

// ── BUY SCORE ALGORITHM: STOCKS ─────────────────────────────────
// Score 0-100 based on 3 factors:
//   40pts — Price position in 52W range (lower = better)
//   35pts — % correction from 52W high (deeper = better entry)
//   25pts — P/E attractiveness (lower = better)
function calcStockScore_(price, high52, low52, pe, target) {
  let score = 0;

  // Factor 1: Position in 52-week range (0=at high, 40=at low)
  if (high52 && low52 && high52 > low52) {
    const range    = high52 - low52;
    const position = (price - low52) / range; // 0=low, 1=high
    score += Math.round((1 - position) * 40);
  }

  // Factor 2: % drop from 52W high (max 35pts at -40% or more)
  if (high52 && high52 > 0) {
    const dropPct = ((high52 - price) / high52) * 100;
    const pts     = Math.min(35, Math.round(dropPct * 0.875)); // 40% drop = 35pts
    score += pts;
  }

  // Factor 3: P/E ratio (max 25pts; P/E ≤10 = full marks, P/E ≥40 = 0)
  if (pe && pe > 0) {
    const pePts = Math.max(0, Math.round(25 - ((pe - 10) / 30) * 25));
    score += Math.min(25, pePts);
  } else {
    score += 12; // neutral if P/E unavailable
  }

  // Bonus: price is below or at target
  if (target > 0 && price <= target) score = Math.min(100, score + 10);

  return Math.min(100, Math.max(0, score));
}

// ── MF ENGINE ────────────────────────────────────────────────────
function refreshMutualFunds_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_MF);
  if (!sh || sh.getLastRow() < 2) return;

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const code = String(row[0]).trim();
    if (!code) continue;

    const mfData = fetchMFData_(code);
    if (!mfData) continue;

    const { nav, low1y, high1y, cagr1y, cagr3y } = mfData;
    const target = parseFloat(row[8]) || 0;
    const score  = calcMFScore_(nav, low1y, high1y, cagr1y, cagr3y, target);
    const signal = scoreToSignal_(score);

    sh.getRange(i+1, 4).setValue(nav);
    sh.getRange(i+1, 5).setValue(low1y   || "");
    sh.getRange(i+1, 6).setValue(high1y  || "");
    sh.getRange(i+1, 7).setValue(cagr1y  !== null ? cagr1y.toFixed(2) : "");
    sh.getRange(i+1, 8).setValue(cagr3y  !== null ? cagr3y.toFixed(2) : "");
    sh.getRange(i+1,10).setValue(score);
    sh.getRange(i+1,11).setValue(signal);
    sh.getRange(i+1,14).setValue(new Date());

    colorSignalCell_(sh, i+1, 11, signal);
    colorScoreCell_(sh, i+1, 10, score);

    if (score >= CFG.BUY_SCORE_THRESHOLD && row[12] !== "Alerted") {
      sendBuyAlert_("Mutual Fund", code, row[1], nav, score, signal);
      logDecision_("Mutual Fund", code, row[1], nav, score, signal, "AUTO ALERT");
      sh.getRange(i+1, 13).setValue("Alerted");
    }
  }
}

// ── BUY SCORE ALGORITHM: MUTUAL FUNDS ───────────────────────────
// Score 0-100:
//   35pts — NAV position in 1Y range (near 1Y low = better entry)
//   35pts — 1Y CAGR performance
//   20pts — 3Y CAGR consistency
//   10pts — NAV at or below target
function calcMFScore_(nav, low1y, high1y, cagr1y, cagr3y, target) {
  let score = 0;

  // Factor 1: NAV position in 1Y range
  if (high1y && low1y && high1y > low1y) {
    const pos = (nav - low1y) / (high1y - low1y);
    score += Math.round((1 - pos) * 35);
  }

  // Factor 2: 1Y CAGR (max 35pts; 20%+ = full)
  if (cagr1y !== null) {
    const pts = Math.min(35, Math.max(0, Math.round(cagr1y * 1.75)));
    score += pts;
  }

  // Factor 3: 3Y CAGR consistency (max 20pts; 15%+ = full)
  if (cagr3y !== null) {
    const pts = Math.min(20, Math.max(0, Math.round(cagr3y * 1.33)));
    score += pts;
  }

  // Bonus: NAV at/below target
  if (target > 0 && nav <= target) score = Math.min(100, score + 10);

  return Math.min(100, Math.max(0, score));
}

// ── FETCH MF DATA FROM AMFI API ──────────────────────────────────
function fetchMFData_(code) {
  try {
    // Fetch historical NAV (all records)
    const url  = "https://api.mfapi.in/mf/" + code;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const json = JSON.parse(resp.getContentText());
    const data = json.data; // [{date, nav}, ...]
    if (!data || data.length < 2) return null;

    const currentNav = parseFloat(data[0].nav);
    const today      = new Date();
    const oneYrAgo   = new Date(); oneYrAgo.setFullYear(today.getFullYear() - 1);
    const threeYrAgo = new Date(); threeYrAgo.setFullYear(today.getFullYear() - 3);

    let navs1y = [], navs3y = [], nav1yAgo = null, nav3yAgo = null;

    for (const d of data) {
      const dt  = parseDate_(d.date);
      const val = parseFloat(d.nav);
      if (!dt || isNaN(val)) continue;
      if (dt >= oneYrAgo)   navs1y.push(val);
      if (dt >= threeYrAgo) navs3y.push(val);
      // Closest to 1Y ago
      if (!nav1yAgo && dt <= oneYrAgo)   nav1yAgo = val;
      if (!nav3yAgo && dt <= threeYrAgo) nav3yAgo = val;
    }

    const low1y  = navs1y.length ? Math.min(...navs1y) : null;
    const high1y = navs1y.length ? Math.max(...navs1y) : null;
    const cagr1y = nav1yAgo ? ((currentNav / nav1yAgo) - 1) * 100 : null;
    const cagr3y = nav3yAgo ? (Math.pow(currentNav / nav3yAgo, 1/3) - 1) * 100 : null;

    return { nav: currentNav, low1y, high1y, cagr1y, cagr3y };
  } catch(e) {
    Logger.log("MF fetch error [" + code + "]: " + e.message);
    return null;
  }
}

function parseDate_(str) {
  // AMFI date format: DD-MMM-YYYY e.g. 05-Jun-2024
  try {
    const parts = str.split("-");
    return new Date(parts[2] + "-" + parts[1] + "-" + parts[0]);
  } catch(e) { return null; }
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
    Utilities.sleep(1500);
    const val = cell.getValue();
    cell.clearContent();
    return (typeof val === "number" && val > 0) ? val : null;
  } catch(e) {
    Logger.log("GF error [" + symbol + "/" + attr + "]: " + e.message);
    return null;
  }
}

// ── SIGNAL & COLOR HELPERS ───────────────────────────────────────
function scoreToSignal_(score) {
  if (score >= 80) return "⚡ STRONG BUY";
  if (score >= 65) return "✅ BUY";
  if (score >= 45) return "👁 WATCH";
  return "⛔ AVOID";
}

function colorSignalCell_(sh, row, col, signal) {
  const cell = sh.getRange(row, col);
  const map  = {
    "⚡ STRONG BUY" : ["#003300","#00ff88"],
    "✅ BUY"        : ["#0a3320","#4ade80"],
    "👁 WATCH"      : ["#2d2000","#fbbf24"],
    "⛔ AVOID"      : ["#2d0000","#f87171"],
  };
  const [bg, fg] = map[signal] || ["#111","#fff"];
  cell.setBackground(bg).setFontColor(fg).setFontWeight("bold");
}

function colorScoreCell_(sh, row, col, score) {
  const cell = sh.getRange(row, col);
  if      (score >= 80) cell.setBackground("#00c853").setFontColor("#000");
  else if (score >= 65) cell.setBackground("#64dd17").setFontColor("#000");
  else if (score >= 45) cell.setBackground("#ffd600").setFontColor("#000");
  else                  cell.setBackground("#d50000").setFontColor("#fff");
}

// ── CAPITAL SPLIT ────────────────────────────────────────────────
function updateCapitalSplit_() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cap  = ss.getSheetByName(CFG.SHEET_CAPITAL);
  if (!cap) return;
  const total      = parseFloat(cap.getRange("B3").getValue()) || 0;
  const stockPct   = parseFloat(cap.getRange("B4").getValue()) || 60;
  const mfPct      = parseFloat(cap.getRange("B5").getValue()) || 40;
  const stockBudget = total * stockPct / 100;
  const mfBudget    = total * mfPct   / 100;

  // Distribute stock budget proportionally by score among BUY+ signals
  distributeCapital_(ss, CFG.SHEET_STOCKS, stockBudget, 9, 11, 10);   // score col9, alloc col11, status col10 (signal)
  distributeCapital_(ss, CFG.SHEET_MF,     mfBudget,   10, 12, 11);
}

function distributeCapital_(ss, sheetName, budget, scoreCol, allocCol, signalCol) {
  const sh   = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getValues();
  let totalScore = 0;
  const eligible = [];
  for (let i = 1; i < data.length; i++) {
    const score  = parseFloat(data[i][scoreCol-1]) || 0;
    const signal = String(data[i][signalCol-1]);
    if (score >= 65 && (signal.includes("BUY"))) {
      totalScore += score;
      eligible.push({ row: i+1, score });
    }
  }
  // Clear all allocations first
  for (let i = 2; i <= data.length; i++) sh.getRange(i, allocCol).setValue(0);
  // Distribute proportionally
  if (totalScore === 0) return;
  for (const e of eligible) {
    const alloc = Math.round((e.score / totalScore) * budget);
    sh.getRange(e.row, allocCol).setValue(alloc);
  }
}

// ── EMAIL ALERT ──────────────────────────────────────────────────
function sendBuyAlert_(type, symbol, name, price, score, signal) {
  const subject = `🔔 Investment Signal: ${signal} — ${name} (Score: ${score}/100)`;
  const body = `
Investment Decision Alert

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Type        : ${type}
  Symbol/Code : ${symbol}
  Name        : ${name}
  Price/NAV   : ₹${typeof price === "number" ? price.toFixed(2) : price}
  Buy Score   : ${score} / 100
  Signal      : ${signal}
  Time        : ${new Date().toLocaleString("en-IN")}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Your scoring engine has flagged this as a high-confidence entry opportunity.
Review your Capital Tracker sheet before placing an order.

— Your Investment Dashboard
  `;
  GmailApp.sendEmail(CFG.ALERT_EMAIL, subject, body);
}

// ── DECISION LOG ─────────────────────────────────────────────────
function logDecision_(type, symbol, name, price, score, signal, action) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.SHEET_LOG);
  if (!log) return;
  log.appendRow([new Date(), type, symbol, name,
    typeof price === "number" ? price.toFixed(2) : price,
    score, signal, action, "✅ " + CFG.ALERT_EMAIL]);
}

// ── MANUAL: Mark Investment Decision ─────────────────────────────
function markInvested() {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row  = sh.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("Select a data row first."); return; }
  const data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const name = data[1] || data[0];
  const ui   = SpreadsheetApp.getUi();
  const res  = ui.alert("Mark Invested", "Record investment decision for:\n" + name + "?", ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    sh.getRange(row, sh.getLastColumn() - 1).setValue("🟢 Invested");
    logDecision_(sh.getName().includes("Stock") ? "Stock" : "MF",
                 data[0], name, data[3] || data[3], data[8] || data[9], "—", "INVESTED");
    ui.alert("✅ Recorded in Decision Log.");
  }
}

// ── RESET ALERT STATUS ───────────────────────────────────────────
function resetAlertStatus() {
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row < 2) return;
  const lastCol = sh.getLastColumn();
  sh.getRange(row, lastCol - 1).setValue("Watching");
  SpreadsheetApp.getUi().alert("Alert status reset to Watching.");
}

// ================================================================
//  CUSTOM MENU
// ================================================================
function onOpen() { createMenu_(); }

function createMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("📈 Investment Dashboard")
    .addItem("▶ Refresh All Now",              "refreshAll")
    .addSeparator()
    .addItem("🏦 Update Capital Split",         "updateCapitalSplit_")
    .addItem("✅ Mark Row as Invested",         "markInvested")
    .addItem("🔄 Reset Row Alert Status",       "resetAlertStatus")
    .addSeparator()
    .addItem("⚙ Full Setup (first time only)", "setupAll")
    .addItem("🗑 Remove Auto-Trigger",          "removeTrigger")
    .addToUi();
}
