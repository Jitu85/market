let recommendations = [
  {
    id: "reliance",
    name: "Reliance Industries",
    symbol: "RELIANCE",
    type: "stock",
    price: "₹2,940.25",
    move: 1.2,
    score: 88,
    signal: "Buy",
    rsi: 58,
    macd: "Bullish",
    pe: "27.4",
    debt: "Low",
    technical: 86,
    fundamental: 83,
    risk: 81,
  },
  {
    id: "hdfc",
    name: "HDFC Bank",
    symbol: "HDFCBANK",
    type: "stock",
    price: "₹1,728.80",
    move: 0.74,
    score: 84,
    signal: "Buy",
    rsi: 54,
    macd: "Bullish",
    pe: "19.8",
    debt: "Low",
    technical: 78,
    fundamental: 89,
    risk: 85,
  },
  {
    id: "ppfas",
    name: "Parag Parikh Flexi Cap – Regular Growth",
    symbol: "PPFAS",
    type: "fund",
    price: "₹82.14",
    move: 0.38,
    score: 82,
    signal: "Buy",
    rsi: 61,
    macd: "Stable",
    pe: "Blend",
    debt: "Very Low",
    technical: 76,
    fundamental: 86,
    risk: 88,
  },
  {
    id: "sbi-gold-direct",
    name: "SBI Gold Fund – Direct Growth",
    symbol: "SBIGOLD",
    schemeCode: "119788",
    type: "fund",
    price: "₹43.6402",
    move: 0,
    score: 0,
    signal: "Watch",
    rsi: 0,
    macd: "NAV",
    pe: "—",
    debt: "—",
    technical: 0,
    fundamental: 0,
    risk: 0,
    marketSource: "AMFI",
    marketStatus: "official-eod-nav",
    marketUpdatedAt: "10-Jul-2026",
  },
  {
    id: "jio-flexicap-direct",
    name: "JioBlackRock Flexi Cap – Direct Growth",
    symbol: "JIOFLEXI",
    schemeCode: "153859",
    type: "fund",
    price: "₹10.0081",
    move: 0,
    score: 0,
    signal: "Watch",
    rsi: 0,
    macd: "NAV",
    pe: "—",
    debt: "—",
    technical: 0,
    fundamental: 0,
    risk: 0,
    marketSource: "AMFI",
    marketStatus: "official-eod-nav",
    marketUpdatedAt: "10-Jul-2026",
  },  {
    id: "tcs",
    name: "Tata Consultancy Services",
    symbol: "TCS",
    type: "stock",
    price: "₹4,012.35",
    move: -0.28,
    score: 76,
    signal: "Hold",
    rsi: 48,
    macd: "Neutral",
    pe: "31.1",
    debt: "Low",
    technical: 70,
    fundamental: 82,
    risk: 78,
  },
  {
    id: "uti",
    name: "UTI Nifty 50 Index – Regular Growth",
    symbol: "UTINIFTY",
    type: "fund",
    price: "₹168.43",
    move: 0.52,
    score: 79,
    signal: "Watch",
    rsi: 57,
    macd: "Stable",
    pe: "Index",
    debt: "None",
    technical: 73,
    fundamental: 80,
    risk: 90,
  },
  {
    id: "itc",
    name: "ITC",
    symbol: "ITC",
    type: "stock",
    price: "₹446.65",
    move: -0.12,
    score: 73,
    signal: "Hold",
    rsi: 46,
    macd: "Neutral",
    pe: "24.2",
    debt: "Low",
    technical: 68,
    fundamental: 80,
    risk: 82,
  },
];

let allocation = [
  { label: "Large Cap", value: 42, color: "#11b981" },
  { label: "Flexi Cap", value: 24, color: "#22c7dc" },
  { label: "Debt Funds", value: 18, color: "#d99116" },
  { label: "Cash", value: 16, color: "#687982" },
];

let alerts = [
  {
    id: "alert-1",
    title: "RELIANCE crossed Oracle Score 85",
    detail: "Technical and fundamentals are aligned for entry.",
    color: "#11b981",
  },
  {
    id: "alert-2",
    title: "TCS momentum cooling",
    detail: "RSI moved below 50 while valuation remains elevated.",
    color: "#d99116",
  },
  {
    id: "alert-3",
    title: "UTI Nifty 50 near rebalance band",
    detail: "Portfolio weight can be adjusted on next SIP cycle.",
    color: "#22c7dc",
  },
];

let logs = [
  { time: "09:45", title: "FetchStocks.gs completed", detail: "6 instruments updated from market API." },
  { time: "09:46", title: "Indicators.gs completed", detail: "RSI, MACD, moving averages recalculated." },
  { time: "09:47", title: "Scoring.gs completed", detail: "Oracle Score generated and stored in Firebase." },
  { time: "09:48", title: "Alerts.gs completed", detail: "3 active rules triggered for watchlist." },
];

const rows = document.querySelector("#recommendationRows");
const detailName = document.querySelector("#detailName");
const detailSignal = document.querySelector("#detailSignal");
const detailSymbol = document.querySelector("#detailSymbol");
const detailPrice = document.querySelector("#detailPrice");
const detailMove = document.querySelector("#detailMove");
const rsiValue = document.querySelector("#rsiValue");
const macdValue = document.querySelector("#macdValue");
const peValue = document.querySelector("#peValue");
const debtValue = document.querySelector("#debtValue");
const technicalMeter = document.querySelector("#technicalMeter");
const fundamentalMeter = document.querySelector("#fundamentalMeter");
const riskMeter = document.querySelector("#riskMeter");
const allocationList = document.querySelector("#allocationList");
const alertList = document.querySelector("#alertList");
const logList = document.querySelector("#logList");
const searchInput = document.querySelector("#searchInput");
const filterButtons = document.querySelectorAll("[data-filter]");
const navFilterLinks = document.querySelectorAll("[data-nav-filter]");
const navItems = document.querySelectorAll(".nav-item");
const refreshButton = document.querySelector("#refreshButton");
const themeToggle = document.querySelector("#themeToggle");
const lastRefresh = document.querySelector("#lastRefresh");
const watchButton = document.querySelector("#watchButton");
const clearAlertsButton = document.querySelector("#clearAlertsButton");
const averageScore = document.querySelector("#averageScore");
const syncState = document.querySelector("#syncState");
const marketStatus = document.querySelector("#marketStatus");
const statusDot = document.querySelector(".status-dot");
const refreshCountdown = document.querySelector("#refreshCountdown");
const sheetsEndpoint = String(window.MARKET_ORACLE_CONFIG?.googleSheetsEndpoint || "").trim();

let selectedId = recommendations[0].id;
let activeFilter = "all";
let hiddenAlerts = new Set();
let watched = new Set();

function scoreClass(signal) {
  return signal.toLowerCase().replace(/\s+/g, "-");
}

function moveClass(move) {
  if (move > 0) return "positive";
  if (move < 0) return "negative";
  return "neutral";
}

function formatMove(move) {
  const prefix = move > 0 ? "+" : "";
  return `${prefix}${move.toFixed(2)}%`;
}

function numberFrom(value, fallback = 0) {
  const parsed = Number(String(value ?? "").replace(/[₹,%\s,]/g, ""));
  return Number.isFinite(parsed) ? parsed : fallback;
}

function recommendationFrom(row, index) {
  const symbol = String(row.symbol || row.Symbol || "").trim().toUpperCase();
  const name = String(row.name || row.Name || symbol).trim();
  if (!symbol || !name) return null;
  const rawPrice = row.price ?? row.Price ?? 0;
  return {
    id: String(row.id || row.ID || symbol || index).toLowerCase().replace(/[^a-z0-9_-]/g, "-"),
    name, symbol,
    type: String(row.type || row.Type || "stock").toLowerCase() === "fund" ? "fund" : "stock",
    price: String(rawPrice).includes("₹") ? String(rawPrice) : numberFrom(rawPrice).toLocaleString("en-IN", { style: "currency", currency: "INR" }),
    move: numberFrom(row.move ?? row.Move),
    score: numberFrom(row.score ?? row.Score),
    signal: String(row.signal || row.Signal || "Watch"),
    rsi: numberFrom(row.rsi ?? row.RSI),
    macd: String(row.macd || row.MACD || "Neutral"),
    pe: String(row.pe ?? row.PE ?? "—"),
    debt: String(row.debt || row.Debt || "—"),
    technical: numberFrom(row.technical ?? row.Technical),
    fundamental: numberFrom(row.fundamental ?? row.Fundamental),
    risk: numberFrom(row.risk ?? row.Risk),
    marketSource: String(row.marketSource || ""),
    marketStatus: String(row.marketStatus || "sample"),
    marketUpdatedAt: String(row.marketUpdatedAt || ""),
  };
}

async function syncGoogleSheets() {
  if (!sheetsEndpoint) {
    syncState.textContent = "Sample data · add endpoint in config.js";
    return;
  }
  syncState.textContent = "Syncing Google Sheets…";
  const joiner = sheetsEndpoint.includes("?") ? "&" : "?";
  const response = await fetch(`${sheetsEndpoint}${joiner}t=${Date.now()}`, { cache: "no-store", headers: { Accept: "application/json" } });
  if (!response.ok) throw new Error(`Google Sheets returned HTTP ${response.status}`);
  const payload = await response.json();
  if (payload.error) throw new Error(payload.error);
  const rows = Array.isArray(payload) ? payload : payload.recommendations;
  if (!Array.isArray(rows)) throw new Error("Response is missing a recommendations array");
  const imported = rows.map(recommendationFrom).filter(Boolean);
  if (!imported.length) throw new Error("Google Sheet contains no valid recommendation rows");
  recommendations = imported;
  averageScore.textContent = (recommendations.reduce((sum, item) => sum + item.score, 0) / recommendations.length).toFixed(1);
  if (Array.isArray(payload.allocation)) allocation = payload.allocation;
  if (Array.isArray(payload.alerts)) alerts = payload.alerts;
  if (Array.isArray(payload.logs)) logs = payload.logs;
  selectedId = recommendations.some((item) => item.id === selectedId) ? selectedId : recommendations[0].id;
  hiddenAlerts = new Set();
  renderRecommendations();
  renderDetail(recommendations.find((item) => item.id === selectedId));
  renderAllocation(); renderAlerts(); renderLogs();
  const updated = payload.updatedAt ? new Date(payload.updatedAt) : new Date();
  const time = Number.isNaN(updated.getTime()) ? "now" : updated.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", hour12: false, timeZone: "Asia/Kolkata" });
  lastRefresh.textContent = `Updated ${time} IST`;
  const errorCount = Array.isArray(payload.marketData?.errors) ? payload.marketData.errors.length : 0;
  syncState.textContent = errorCount ? `Synced · ${errorCount} fallback` : "NSE + AMFI synced";
  if (errorCount) syncState.title = payload.marketData.errors.join("\n");
  else syncState.removeAttribute("title");
}

function updateMarketStatus() {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Kolkata",
      weekday: "short",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).formatToParts(new Date()).filter((part) => part.type !== "literal").map((part) => [part.type, part.value]),
  );
  const minute = Number(parts.hour) * 60 + Number(parts.minute);
  const weekday = parts.weekday;
  const isWeekday = !["Sat", "Sun"].includes(weekday);
  const isOpen = isWeekday && minute >= 555 && minute < 930;
  marketStatus.textContent = isOpen ? "Market Open" : "Market Closed";
  statusDot.classList.toggle("is-closed", !isOpen);
}

let nextAutomaticSync = Date.now() + 5 * 60 * 1000;
function updateRefreshCountdown() {
  const remaining = Math.max(0, nextAutomaticSync - Date.now());
  const minutes = Math.floor(remaining / 60000);
  const seconds = Math.floor((remaining % 60000) / 1000);
  refreshCountdown.textContent = `${minutes}:${String(seconds).padStart(2, "0")}`;
  if (remaining === 0) {
    nextAutomaticSync = Date.now() + 5 * 60 * 1000;
    syncGoogleSheets().catch((error) => {
      console.error("Automatic market sync failed", error);
      syncState.textContent = "Auto-sync failed · showing last data";
      syncState.title = error.message;
    });
  }
}
function buildCell(tag, className, text) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  element.textContent = text;
  return element;
}

function renderRecommendations() {
  const query = searchInput.value.trim().toLowerCase();
  const visible = recommendations.filter((item) => {
    const matchesFilter = activeFilter === "all" || item.type === activeFilter;
    const matchesQuery = `${item.name} ${item.symbol}`.toLowerCase().includes(query);
    return matchesFilter && matchesQuery;
  });

  rows.replaceChildren();

  visible.forEach((item) => {
    const tr = document.createElement("tr");
    tr.tabIndex = 0;
    tr.className = item.id === selectedId ? "is-selected" : "";
    tr.dataset.id = item.id;

    const instrument = document.createElement("td");
    const instrumentWrap = document.createElement("div");
    instrumentWrap.className = "instrument";
    instrumentWrap.append(buildCell("strong", "", item.name), buildCell("span", "", item.symbol));
    instrument.append(instrumentWrap);

    const type = document.createElement("td");
    type.append(buildCell("span", "type-chip", item.type === "stock" ? "Stock" : "Mutual Fund"));

    const price = document.createElement("td");
    price.className = "market-price";
    price.append(buildCell("strong", "", item.price));
    if (item.marketSource) {
      const source = buildCell("small", "quote-source", item.marketSource);
      if (item.marketStatus === "fallback-sheet-value") source.classList.add("is-fallback");
      source.title = item.marketUpdatedAt || item.marketStatus;
      price.append(source);
    }
    const move = buildCell("td", moveClass(item.move), formatMove(item.move));
    const score = buildCell("td", "", String(item.score));

    const signal = document.createElement("td");
    signal.append(buildCell("span", `signal-chip ${scoreClass(item.signal)}`, item.signal));

    const action = document.createElement("td");
    const button = buildCell("button", "row-button", "Inspect");
    button.type = "button";
    action.append(button);

    tr.append(instrument, type, price, move, score, signal, action);
    rows.append(tr);
  });
}

function renderDetail(item) {
  detailName.textContent = item.name;
  detailSignal.textContent = item.signal;
  detailSignal.className = `signal-badge ${scoreClass(item.signal)}`;
  detailSymbol.textContent = item.symbol;
  detailPrice.textContent = item.price;
  detailMove.textContent = formatMove(item.move);
  detailMove.className = moveClass(item.move);
  rsiValue.textContent = item.rsi;
  macdValue.textContent = item.macd;
  peValue.textContent = item.pe;
  debtValue.textContent = item.debt;
  technicalMeter.value = item.technical;
  fundamentalMeter.value = item.fundamental;
  riskMeter.value = item.risk;
  watchButton.lastChild.textContent = watched.has(item.id) ? "Watching" : "Add to Watchlist";
}

function renderAllocation() {
  allocationList.replaceChildren();
  allocation.forEach((item) => {
    const row = document.createElement("div");
    row.className = "allocation-row";

    const label = buildCell("span", "", item.label);
    const track = document.createElement("div");
    track.className = "allocation-track";
    const fill = document.createElement("span");
    fill.style.setProperty("--value", `${item.value}%`);
    fill.style.setProperty("--color", item.color);
    track.append(fill);
    const value = buildCell("strong", "", `${item.value}%`);

    row.append(label, track, value);
    allocationList.append(row);
  });
}

function renderAlerts() {
  alertList.replaceChildren();
  alerts
    .filter((alert) => !hiddenAlerts.has(alert.id))
    .forEach((alert) => {
      const item = document.createElement("div");
      item.className = "alert-item";

      const level = document.createElement("span");
      level.className = "alert-level";
      level.style.setProperty("--level-color", alert.color);

      const copy = document.createElement("div");
      copy.append(buildCell("strong", "", alert.title), buildCell("span", "", alert.detail));

      const button = buildCell("button", "", "Ã—");
      button.type = "button";
      button.setAttribute("aria-label", `Dismiss ${alert.title}`);
      button.addEventListener("click", () => {
        hiddenAlerts.add(alert.id);
        renderAlerts();
      });

      item.append(level, copy, button);
      alertList.append(item);
    });

  if (!alertList.children.length) {
    const empty = buildCell("p", "empty-state", "No unread alerts.");
    alertList.append(empty);
  }
}

function renderLogs() {
  logList.replaceChildren();
  logs.forEach((log) => {
    const item = document.createElement("li");
    item.append(buildCell("span", "log-time", log.time));
    const copy = document.createElement("div");
    copy.append(buildCell("strong", "", log.title), buildCell("span", "", log.detail));
    item.append(copy);
    logList.append(item);
  });
}

function setSelected(id) {
  const item = recommendations.find((entry) => entry.id === id) || recommendations[0];
  selectedId = item.id;
  renderRecommendations();
  renderDetail(item);
}

rows.addEventListener("click", (event) => {
  const row = event.target.closest("tr");
  if (row) setSelected(row.dataset.id);
});

rows.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  const row = event.target.closest("tr");
  if (!row) return;
  event.preventDefault();
  setSelected(row.dataset.id);
});

function applyRecommendationFilter(filter) {
  activeFilter = filter;
  filterButtons.forEach((button) => button.classList.toggle("is-active", button.dataset.filter === filter));
  navItems.forEach((link) => link.classList.remove("is-active"));
  navFilterLinks.forEach((link) => link.classList.toggle("is-active", link.dataset.navFilter === filter));
  renderRecommendations();
}

filterButtons.forEach((button) => {
  button.addEventListener("click", () => applyRecommendationFilter(button.dataset.filter));
});

navFilterLinks.forEach((link) => {
  link.addEventListener("click", () => applyRecommendationFilter(link.dataset.navFilter));
});

searchInput.addEventListener("input", renderRecommendations);

refreshButton.addEventListener("click", async () => {
  refreshButton.classList.add("is-loading");
  refreshButton.disabled = true;
  try {
    await syncGoogleSheets();
  } catch (error) {
    console.error("Google Sheets sync failed", error);
    syncState.textContent = "Sheets sync failed · showing cached data";
    syncState.title = error.message;
  } finally {
    refreshButton.classList.remove("is-loading");
    refreshButton.disabled = false;
  }
});

themeToggle.addEventListener("click", () => {
  const root = document.documentElement;
  root.dataset.theme = root.dataset.theme === "dark" ? "light" : "dark";
});

watchButton.addEventListener("click", () => {
  if (watched.has(selectedId)) {
    watched.delete(selectedId);
  } else {
    watched.add(selectedId);
  }
  renderDetail(recommendations.find((item) => item.id === selectedId));
});

clearAlertsButton.addEventListener("click", () => {
  hiddenAlerts = new Set(alerts.map((alert) => alert.id));
  renderAlerts();
});

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js?v=8").catch(() => {});
  });
}

renderRecommendations();
renderDetail(recommendations[0]);
renderAllocation();
renderAlerts();
renderLogs();

updateMarketStatus();
updateRefreshCountdown();
window.setInterval(updateMarketStatus, 60 * 1000);
window.setInterval(updateRefreshCountdown, 1000);

syncGoogleSheets().catch((error) => {
  console.error("Google Sheets sync failed", error);
  syncState.textContent = "Sheets sync failed · showing sample data";
  syncState.title = error.message;
});










