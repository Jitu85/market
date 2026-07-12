/**
 * MARKET ORACLE Google Apps Script Web App.
 *
 * Recommendations headers:
 * id, name, symbol, type, schemeCode, price, move, score, signal, rsi,
 * macd, pe, debt, technical, fundamental, risk
 *
 * Stocks use NSE symbols. Funds must use the exact AMFI schemeCode so a Direct
 * plan is never confused with a Regular plan or Growth with IDCW.
 */
function doGet() {
  try {
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var recommendations = readSheet_(book, "Recommendations", true);
    var marketResult = enrichRecommendations_(recommendations);

    var result = {
      recommendations: marketResult.rows,
      allocation: readSheet_(book, "Allocation"),
      alerts: readSheet_(book, "Alerts"),
      logs: readSheet_(book, "Logs"),
      updatedAt: marketResult.updatedAt,
      marketData: {
        stockSource: "Google Finance (NSE)",
        mutualFundSource: "AMFI India",
        stockQuotePolicy: "Google Finance NSE quote; may be delayed up to 20 minutes.",
        navPolicy: "Latest published end-of-day NAV; mutual funds do not trade intraday.",
        errors: marketResult.errors,
      },
    };

    return json_(result);
  } catch (error) {
    return json_({ error: error.message });
  }
}

function enrichRecommendations_(rows) {
  var fundCodes = rows.filter(function (row) {
    return String(value_(row, "type") || "stock").toLowerCase() === "fund";
  }).map(schemeCodeFor_).filter(String);
  var amfi = fundCodes.length ? fetchAmfiNav_(fundCodes) : {};
  var errors = [];
  var timestamps = [];

  var enriched = rows.map(function (row) {
    var type = String(value_(row, "type") || "stock").toLowerCase();
    var copy = Object.assign({}, row);

    try {
      if (type === "fund") {
        var schemeCode = schemeCodeFor_(row);
        if (!schemeCode) throw new Error("schemeCode is required for an exact AMFI NAV match");
        if (!amfi) amfi = fetchAmfiNav_();
        var fund = amfi[schemeCode];
        if (!fund) throw new Error("AMFI schemeCode " + schemeCode + " was not found");
        copy.price = fund.nav;
        copy.navDate = fund.date;
        copy.marketSource = "AMFI";
        copy.marketStatus = "official-eod-nav";
        copy.marketUpdatedAt = fund.date;
        timestamps.push(parseAmfiDate_(fund.date));
      } else {
        var symbol = String(value_(row, "symbol") || "").trim().toUpperCase();
        if (!symbol) throw new Error("NSE symbol is required");
        var sheetPrice = Number(String(value_(row, "price") || "").replace(/[₹,\s]/g, ""));
        if (!isFinite(sheetPrice)) throw new Error("Google Finance price is not available");
        copy.price = sheetPrice;
        copy.move = Number(String(value_(row, "move") || "0").replace(/[%,\s]/g, "")) || 0;
        copy.marketSource = "Google Finance";
        copy.marketStatus = "delayed-market-quote";
        copy.marketUpdatedAt = new Date().toISOString();
        timestamps.push(new Date());
      }
    } catch (error) {
      copy.marketStatus = "fallback-sheet-value";
      copy.marketError = error.message;
      errors.push((copy.symbol || copy.name || "Unknown") + ": " + error.message);
    }

    return copy;
  });

  var validTimes = timestamps.filter(function (date) {
    return date && !isNaN(date.getTime());
  });
  var latest = validTimes.length
    ? new Date(Math.max.apply(null, validTimes.map(function (date) { return date.getTime(); })))
    : new Date();

  return { rows: enriched, errors: errors, updatedAt: latest.toISOString() };
}

function schemeCodeFor_(row) {
  var explicit = String(value_(row, "schemeCode") || "").trim();
  if (explicit) return explicit;

  var normalized = String(value_(row, "name") || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  var known = {
    "parag parikh flexi cap": "122640",
    "parag parikh flexi cap regular growth": "122640",
    "parag parikh flexi cap fund regular plan growth": "122640",
    "uti nifty 50 index": "100822",
    "uti nifty 50 index regular growth": "100822",
    "uti nifty 50 index fund regular plan growth option": "100822",
  };
  return known[normalized] || "";
}

/**
 * Run once from the Apps Script editor after adding the sheet headers.
 * It installs supported Google Finance formulas for every stock row.
 */
function setupMarketData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recommendations");
  if (!sheet) throw new Error('Required sheet "Recommendations" was not found');

  var range = sheet.getDataRange();
  var values = range.getDisplayValues();
  if (values.length < 1) throw new Error("Recommendations has no header row");
  var headers = values[0].map(function (header) { return String(header).trim().toLowerCase(); });
  var symbolColumn = headers.indexOf("symbol") + 1;
  var typeColumn = headers.indexOf("type") + 1;
  var priceColumn = headers.indexOf("price") + 1;
  var moveColumn = headers.indexOf("move") + 1;
  if (!symbolColumn || !typeColumn || !priceColumn || !moveColumn) {
    throw new Error("Recommendations must contain symbol, type, price, and move headers");
  }


  addRequestedFunds_(sheet, headers);
  values = sheet.getDataRange().getDisplayValues();
  for (var row = 2; row <= values.length; row++) {
    var type = String(sheet.getRange(row, typeColumn).getDisplayValue() || "stock").toLowerCase();
    if (type !== "stock") continue;
    var symbolCell = sheet.getRange(row, symbolColumn).getA1Notation();
    sheet.getRange(row, priceColumn).setFormula('=GOOGLEFINANCE("NSE:"&' + symbolCell + ',"price")');
    sheet.getRange(row, moveColumn).setFormula('=GOOGLEFINANCE("NSE:"&' + symbolCell + ',"changepct")');
  }
  SpreadsheetApp.flush();
}
function addRequestedFunds_(sheet, headers) {
  var symbolIndex = headers.indexOf("symbol");
  var existing = {};
  sheet.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
    existing[String(row[symbolIndex] || "").trim().toUpperCase()] = true;
  });

  var funds = [
    {
      id: "sbi-gold-direct",
      name: "SBI Gold Fund - Direct Plan - Growth",
      symbol: "SBIGOLD",
      type: "fund",
      schemecode: "119788",
      move: 0,
      score: 0,
      signal: "Watch",
      macd: "NAV",
    },
    {
      id: "jio-flexicap-direct",
      name: "JioBlackRock Flexi Cap Fund - Direct Plan - Growth Option",
      symbol: "JIOFLEXI",
      type: "fund",
      schemecode: "153859",
      move: 0,
      score: 0,
      signal: "Watch",
      macd: "NAV",
    },
  ];

  funds.forEach(function (fund) {
    if (existing[fund.symbol]) return;
    sheet.appendRow(headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(fund, header) ? fund[header] : "";
    }));
  });
}

function fetchAmfiNav_(schemeCodes) {
  var cache = CacheService.getScriptCache();
  var result = {};
  var missing = [];

  schemeCodes.forEach(function (code) {
    var saved = cache.get("amfi:" + code);
    if (saved) result[code] = JSON.parse(saved);
    else missing.push(code);
  });
  if (!missing.length) return result;

  var response = UrlFetchApp.fetch("https://portal.amfiindia.com/spages/NAVAll.txt", {
    muteHttpExceptions: true,
    followRedirects: true,
  });
  if (response.getResponseCode() !== 200) {
    throw new Error("AMFI NAV request returned HTTP " + response.getResponseCode());
  }

  var wanted = {};
  missing.forEach(function (code) { wanted[code] = true; });
  response.getContentText().split(/\r?\n/).forEach(function (line) {
    var fields = line.split(";");
    if (fields.length < 6 || !wanted[fields[0]]) return;
    var nav = Number(fields[4]);
    if (!isFinite(nav)) return;
    var record = { name: fields[3], nav: nav, date: fields[5] };
    result[fields[0]] = record;
    cache.put("amfi:" + fields[0], JSON.stringify(record), 21600);
  });
  return result;
}

function cookiesFrom_(headers) {
  var raw = headers["Set-Cookie"] || headers["set-cookie"] || [];
  if (!Array.isArray(raw)) raw = [raw];
  return raw.map(function (cookie) {
    return String(cookie).split(";")[0];
  }).filter(String).join("; ");
}

function parseAmfiDate_(value) {
  var match = String(value).match(/^(\d{2})-(\w{3})-(\d{4})$/);
  if (!match) return new Date(value);
  var months = { Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5, Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11 };
  return new Date(Date.UTC(Number(match[3]), months[match[2]], Number(match[1])));
}

function value_(row, name) {
  if (Object.prototype.hasOwnProperty.call(row, name)) return row[name];
  var wanted = name.toLowerCase();
  var key = Object.keys(row).find(function (candidate) {
    return candidate.toLowerCase() === wanted;
  });
  return key ? row[key] : "";
}

function readSheet_(book, name, required) {
  var sheet = book.getSheetByName(name);
  if (!sheet) {
    if (required) throw new Error('Required sheet "' + name + '" was not found');
    return [];
  }

  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values.shift().map(function (header) { return String(header).trim(); });

  return values.filter(function (row) {
    return row.some(function (cell) { return String(cell).trim() !== ""; });
  }).map(function (row) {
    return headers.reduce(function (record, header, index) {
      if (header) record[header] = row[index];
      return record;
    }, {});
  });
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}





