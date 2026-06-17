// ================================================================
//  MARKET ORACLE — Google Apps Script
//  Fetches live stock prices + MF NAVs, scores them,
//  then pushes data.json to GitHub Pages.
//  No JSONP. No CORS. No deployment URLs.
//
//  SETUP (one time only):
//  1. Paste this into Apps Script → Save → Run setupAll()
//  2. Authorise permissions
//  3. Done — trigger runs every 30 minutes automatically
// ================================================================

const CFG = {
  GITHUB_TOKEN : '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',
  GITHUB_OWNER : 'jitu85',
  GITHUB_REPO  : 'market',
  GITHUB_FILE  : 'data.json',
  GITHUB_BRANCH: 'main',
  ALERT_EMAIL  : Session.getActiveUser().getEmail(),
  SCORE_STRONG : 80,
  SCORE_BUY    : 65,
  SCORE_WATCH  : 45,
};

// ── YOUR WATCHLIST ───────────────────────────────────────────────
const STOCKS = [
  { symbol:'NSE:SILVERBEES', name:'Nippon India Silver ETF',    sector:'Silver ETF',       target:100 },
  { symbol:'NSE:HDFCSILVER', name:'HDFC Silver ETF',            sector:'Silver ETF',       target:70  },
  { symbol:'NSE:ICICIB22',   name:'ICICI Pru Bharat 22 ETF',    sector:'Govt PSU ETF',     target:88  },
  { symbol:'NSE:SUZLON',     name:'Suzlon Energy Ltd',          sector:'Renewable Energy', target:48  },
  { symbol:'NSE:TRIDENT',    name:'Trident Ltd',                sector:'Textiles',         target:30  },
  { symbol:'NSE:RELIANCE',   name:'Reliance Industries',        sector:'Energy',           target:2600},
  { symbol:'NSE:HDFCBANK',   name:'HDFC Bank',                  sector:'Banking',          target:1450},
];

const MFS = [
  { code:'120403', name:'Invesco India Mid Cap Fund — Direct Growth',       category:'Mid Cap',    target:200  },
  { code:'147622', name:'Motilal Oswal Multicap Fund — Direct Growth',      category:'Multicap',   target:38   },
  { code:'122639', name:'Parag Parikh Flexi Cap Fund — Direct Growth',      category:'Flexi Cap',  target:75   },
  { code:'118825', name:'Nippon India Large Cap Fund — Direct Growth',      category:'Large Cap',  target:88   },
  { code:'119019', name:'SBI Gold Direct Plan Growth',                      category:'Gold',       target:370  },
  { code:'153137', name:'JioBlackRock Flexi Cap Fund — Direct Growth',      category:'Flexi Cap',  target:9    },
  { code:'130503', name:'SBI Multi Asset Allocation Fund — Direct Growth',  category:'Multi Asset',target:48   },
  { code:'149240', name:'Bandhan Small Cap Fund — Direct Growth',           category:'Small Cap',  target:12   },
  { code:'119773', name:'DSP Natural Resources & New Energy Fund — Direct', category:'Sectoral',   target:78   },
];

// ================================================================
//  SETUP — run once
// ================================================================
function setupAll() {
  installTrigger_();
  const ui = SpreadsheetApp.getUi
    ? SpreadsheetApp.getUi()
    : null;
  const msg =
    '✅ Market Oracle Setup Complete!\n\n' +
    'Trigger installed — refreshes every 30 minutes.\n' +
    'Running first refresh now...\n\n' +
    'GitHub: github.com/' + CFG.GITHUB_OWNER + '/' + CFG.GITHUB_REPO;
  if (ui) ui.alert(msg);
  else    Logger.log(msg);
  refreshAll(); // run immediately
}

// ================================================================
//  TRIGGER
// ================================================================
function installTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshAll')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshAll')
    .timeBased().everyMinutes(30).create();
  Logger.log('Trigger installed: refreshAll every 30 minutes');
}

// ================================================================
//  MAIN REFRESH — fetches all data, scores, pushes to GitHub
// ================================================================
function refreshAll() {
  Logger.log('=== Market Oracle Refresh Started ===');
  const stocks = refreshStocks_();
  const mfs    = refreshMFs_();
  const payload = buildPayload_(stocks, mfs);
  pushToGitHub_(payload);
  checkAlerts_(stocks, mfs);
  Logger.log('=== Refresh Complete ===');
}

// ================================================================
//  STOCKS — fetch via GoogleFinance
// ================================================================
function refreshStocks_() {
  const results = [];
  // Use a temporary sheet for GoogleFinance formulas
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   tmp = ss.getSheetByName('__tmp__');
  if (!tmp) { tmp = ss.insertSheet('__tmp__'); tmp.hideSheet(); }

  for (const s of STOCKS) {
    try {
      const price  = gfGet_(tmp, s.symbol, 'price');
      const high52 = gfGet_(tmp, s.symbol, '52weekhigh');
      const low52  = gfGet_(tmp, s.symbol, '52weeklow');
      const pe     = gfGet_(tmp, s.symbol, 'pe');
      const score  = scoreStock_(price, high52, low52, pe, s.target);
      const signal = toSignal_(score);

      results.push({
        symbol : s.symbol,
        name   : s.name,
        sector : s.sector,
        price  : price  || null,
        high52 : high52 || null,
        low52  : low52  || null,
        pe     : pe     || null,
        target : s.target,
        score,
        signal,
        updated: new Date().toISOString(),
      });
      Logger.log(s.symbol + ' → ₹' + price + ' | Score: ' + score + ' | ' + signal);
    } catch(e) {
      Logger.log('Error fetching ' + s.symbol + ': ' + e.message);
      results.push({
        symbol:s.symbol, name:s.name, sector:s.sector,
        price:null, high52:null, low52:null, pe:null,
        target:s.target, score:0, signal:'⛔ AVOID',
        updated:new Date().toISOString(), error:e.message
      });
    }
  }
  return results;
}

// ── GoogleFinance single value ────────────────────────────────────
function gfGet_(tmp, symbol, attr) {
  const cell = tmp.getRange('A1');
  cell.setFormula('=GOOGLEFINANCE("' + symbol + '","' + attr + '")');
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  const val = cell.getValue();
  cell.clearContent();
  return (typeof val === 'number' && val > 0) ? val : null;
}

// ================================================================
//  MUTUAL FUNDS — fetch from AMFI via mfapi.in
// ================================================================
function refreshMFs_() {
  const results = [];
  for (const m of MFS) {
    try {
      const data   = fetchMFAPI_(m.code);
      const nav    = data.nav;
      const low1y  = data.low1y;
      const high1y = data.high1y;
      const cagr1y = data.cagr1y;
      const cagr3y = data.cagr3y;
      const score  = scoreMF_(nav, low1y, high1y, cagr1y, cagr3y, m.target);
      const signal = toSignal_(score);

      results.push({
        code    : m.code,
        name    : m.name,
        category: m.category,
        nav, low1y, high1y, cagr1y, cagr3y,
        target  : m.target,
        score, signal,
        updated : new Date().toISOString(),
      });
      Logger.log(m.code + ' NAV:₹' + nav + ' | Score: ' + score + ' | ' + signal);
    } catch(e) {
      Logger.log('Error fetching MF ' + m.code + ': ' + e.message);
      results.push({
        code:m.code, name:m.name, category:m.category,
        nav:null, low1y:null, high1y:null, cagr1y:null, cagr3y:null,
        target:m.target, score:0, signal:'⛔ AVOID',
        updated:new Date().toISOString(), error:e.message
      });
    }
  }
  return results;
}

function fetchMFAPI_(code) {
  const url  = 'https://api.mfapi.in/mf/' + code;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions:true });
  if (resp.getResponseCode() !== 200)
    throw new Error('mfapi returned ' + resp.getResponseCode());

  const json = JSON.parse(resp.getContentText());
  const data = json.data;
  if (!data || data.length < 2) throw new Error('Insufficient NAV data');

  const currentNav = parseFloat(data[0].nav);
  const now        = new Date();
  const yr1        = new Date(now); yr1.setFullYear(now.getFullYear() - 1);
  const yr3        = new Date(now); yr3.setFullYear(now.getFullYear() - 3);

  let navs1y=[], nav1yAgo=null, nav3yAgo=null;

  for (const d of data) {
    const dt  = parseMFDate_(d.date);
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
    cagr1y : nav1yAgo ? ((currentNav/nav1yAgo)-1)*100 : null,
    cagr3y : nav3yAgo ? (Math.pow(currentNav/nav3yAgo,1/3)-1)*100 : null,
  };
}

function parseMFDate_(str) {
  try {
    const p = str.split('-');
    return new Date(p[2]+'-'+p[1]+'-'+p[0]);
  } catch(e) { return null; }
}

// ================================================================
//  SCORING
// ================================================================
function scoreStock_(price, high52, low52, pe, target) {
  if (!price) return 0;
  let score = 0;
  // 40pts — position in 52W range (lower = better entry)
  if (high52 && low52 && high52 > low52)
    score += Math.round((1-(price-low52)/(high52-low52))*40);
  // 35pts — % drop from 52W high
  if (high52)
    score += Math.min(35, Math.round(((high52-price)/high52*100)*0.875));
  // 25pts — P/E attractiveness
  if (pe && pe > 0)
    score += Math.min(25, Math.max(0, Math.round(25-((pe-10)/30)*25)));
  else
    score += 12; // neutral when P/E unavailable (ETFs)
  // Bonus: price at or below target
  if (target > 0 && price <= target) score = Math.min(100, score+10);
  return Math.min(100, Math.max(0, score));
}

function scoreMF_(nav, low1y, high1y, cagr1y, cagr3y, target) {
  if (!nav) return 0;
  let score = 0;
  // 35pts — NAV position in 1Y range
  if (high1y && low1y && high1y > low1y)
    score += Math.round((1-(nav-low1y)/(high1y-low1y))*35);
  // 35pts — 1Y CAGR
  if (cagr1y) score += Math.min(35, Math.max(0, Math.round(cagr1y*1.75)));
  // 20pts — 3Y CAGR
  if (cagr3y) score += Math.min(20, Math.max(0, Math.round(cagr3y*1.33)));
  // Bonus: NAV at or below target
  if (target > 0 && nav <= target) score = Math.min(100, score+10);
  return Math.min(100, Math.max(0, score));
}

function toSignal_(score) {
  if (score >= CFG.SCORE_STRONG) return '⚡ STRONG BUY';
  if (score >= CFG.SCORE_BUY)    return '✅ BUY';
  if (score >= CFG.SCORE_WATCH)  return '👁 WATCH';
  return '⛔ AVOID';
}

// ================================================================
//  BUILD PAYLOAD
// ================================================================
function buildPayload_(stocks, mfs) {
  return {
    timestamp : new Date().toISOString(),
    stocks,
    mfs,
    capital   : { total:100000, stockPct:60, mfPct:40 },
    log       : [],
  };
}

// ================================================================
//  PUSH TO GITHUB — writes data.json to jitu85/market repo
// ================================================================
function pushToGitHub_(payload) {
  const content = JSON.stringify(payload, null, 2);
  const encoded = Utilities.base64Encode(
    Utilities.newBlob(content, 'application/json').getBytes()
  );

  const apiUrl = 'https://api.github.com/repos/'
    + CFG.GITHUB_OWNER + '/'
    + CFG.GITHUB_REPO  + '/contents/'
    + CFG.GITHUB_FILE;

  const headers = {
    'Authorization' : 'token ' + CFG.GITHUB_TOKEN,
    'Content-Type'  : 'application/json',
    'Accept'        : 'application/vnd.github.v3+json',
    'User-Agent'    : 'MarketOracle-GAS',
  };

  // Get current file SHA (required for updates)
  let sha = null;
  try {
    const getResp = UrlFetchApp.fetch(apiUrl, {
      method          : 'get',
      headers,
      muteHttpExceptions: true,
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
    }
  } catch(e) {
    Logger.log('Could not get SHA (first push?): ' + e.message);
  }

  // Push new content
  const body = {
    message : 'Market Oracle data update — ' + new Date().toISOString(),
    content : encoded,
    branch  : CFG.GITHUB_BRANCH,
  };
  if (sha) body.sha = sha;

  const putResp = UrlFetchApp.fetch(apiUrl, {
    method            : 'put',
    headers,
    payload           : JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const code = putResp.getResponseCode();
  if (code === 200 || code === 201) {
    Logger.log('✅ data.json pushed to GitHub successfully (HTTP ' + code + ')');
  } else {
    Logger.log('❌ GitHub push failed: HTTP ' + code);
    Logger.log(putResp.getContentText().slice(0, 500));
  }
}

// ================================================================
//  EMAIL ALERTS
// ================================================================
function checkAlerts_(stocks, mfs) {
  const all = [...stocks, ...mfs];
  const strongBuys = all.filter(x => x.score >= CFG.SCORE_STRONG);
  const buys       = all.filter(x => x.score >= CFG.SCORE_BUY && x.score < CFG.SCORE_STRONG);

  if (strongBuys.length === 0 && buys.length === 0) return;

  const lines = [
    'Market Oracle — Investment Signal Alert',
    'Generated: ' + new Date().toLocaleString('en-IN'),
    '',
  ];

  if (strongBuys.length) {
    lines.push('⚡ STRONG BUY SIGNALS:');
    strongBuys.forEach(x => lines.push(
      '  ' + (x.symbol||x.code) + ' — ' + x.name +
      ' | Score: ' + x.score + '/100'
    ));
    lines.push('');
  }
  if (buys.length) {
    lines.push('✅ BUY SIGNALS:');
    buys.forEach(x => lines.push(
      '  ' + (x.symbol||x.code) + ' — ' + x.name +
      ' | Score: ' + x.score + '/100'
    ));
  }
  lines.push('', '— Market Oracle · Abhijit Kumar Misra (Jitu)');

  GmailApp.sendEmail(
    CFG.ALERT_EMAIL,
    '🔔 Market Oracle: ' + strongBuys.length + ' Strong Buy, ' + buys.length + ' Buy signals',
    lines.join('\n')
  );
  Logger.log('Alert email sent to ' + CFG.ALERT_EMAIL);
}
