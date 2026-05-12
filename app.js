const priceDates = [
  "2024-01-15",
  "2024-03-15",
  "2024-06-17",
  "2024-09-16",
  "2024-12-16",
  "2025-03-17",
  "2025-06-16",
  "2025-09-15",
  "2025-12-15",
  "2026-03-16",
  "2026-05-01",
];

const stocks = {
  "2330": {
    name: "台積電",
    market: "TWSE",
    sector: "半導體",
    volume: 58000,
    marketCap: 37800000,
    prices: [584, 753, 909, 947, 1070, 955, 1015, 1265, 1420, 1385, 1460],
  },
  "2454": {
    name: "聯發科",
    market: "TWSE",
    sector: "半導體",
    volume: 14200,
    marketCap: 2680000,
    prices: [942, 1045, 1455, 1170, 1260, 1425, 1365, 1480, 1610, 1525, 1685],
  },
  "0050": {
    name: "元大台灣50",
    market: "TWSE",
    sector: "ETF",
    volume: 91000,
    marketCap: 415000,
    prices: [135, 150, 171, 181, 190, 177, 184, 205, 218, 211, 224],
  },
  "006208": {
    name: "富邦台50",
    market: "TWSE",
    sector: "ETF",
    volume: 38000,
    marketCap: 182000,
    prices: [78, 87, 99, 104, 110, 102, 106, 119, 126, 122, 130],
  },
  "2303": {
    name: "聯電",
    market: "TWSE",
    sector: "半導體",
    volume: 66000,
    marketCap: 650000,
    prices: [50, 52, 55, 53, 46, 42, 45, 48, 51, 49, 52],
  },
  "2317": {
    name: "鴻海",
    market: "TWSE",
    sector: "電子代工",
    volume: 73000,
    marketCap: 3150000,
    prices: [104, 126, 182, 188, 184, 160, 168, 193, 205, 198, 214],
  },
  "2881": {
    name: "富邦金",
    market: "TWSE",
    sector: "金融",
    volume: 24000,
    marketCap: 1180000,
    prices: [64, 68, 78, 86, 91, 88, 92, 96, 101, 98, 103],
  },
  "1301": {
    name: "台塑",
    market: "TWSE",
    sector: "塑化",
    volume: 18500,
    marketCap: 430000,
    prices: [78, 75, 70, 66, 61, 58, 56, 59, 62, 60, 63],
  },
};

const periodSteps = {
  today: 1,
  "1d": 1,
  "1w": 1,
  "1m": 1,
  "1y": 4,
  all: priceDates.length - 1,
};

const makeId = () => globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`;

let trades = [
  { id: makeId(), symbol: "2330", date: "2024-01-15", amount: 50000 },
  { id: makeId(), symbol: "0050", date: "2024-06-17", amount: 80000 },
  { id: makeId(), symbol: "2454", date: "2025-03-17", amount: 60000 },
];

const state = {
  benchmark: "0050",
  selectedTradeId: trades[0].id,
  watchlist: ["2330", "2454", "0050", "006208", "2317"],
  portfolioSymbols: ["2330", "0050", "2454", "2881"],
  portfolioAmount: 300000,
  portfolioWeights: {},
  compareSymbols: ["2330", "2454", "0050", "2317"],
  comparePeriod: "1y",
  compareStartDate: "2025-03-17",
  compareEndDate: "2026-05-08",
  rankingPeriod: "today",
  sector: "全部",
  dcaSymbols: ["0050", "006208"],
  dcaFrequency: "monthly",
  dcaAmount: 10000,
  dcaStartDate: "2024-01-15",
  dcaEndDate: "2026-05-01",
  riskSymbol: "2330",
  marketDataSource: "sample",
  marketDataDate: "",
};

const TWSE_DAILY_URL = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL";
const TWSE_COMPANY_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L";
const TPEX_DAILY_URL = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=data";
const TWSE_HISTORY_URL = "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX";
const TPEX_HISTORY_URL = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php";
const CACHE_KEY = "decision-ledger-twse-cache-v1";
const HISTORY_CLOSE_CACHE_KEY = "decision-ledger-history-close-cache-v1";
const ADJUSTED_CLOSE_CACHE_KEY = "decision-ledger-adjusted-close-cache-v1";
const ADJUSTED_HISTORY_CACHE_KEY = "decision-ledger-adjusted-history-cache-v1";
let marketHistoryPeriods = {};
let marketDcaSnapshots = {};
let marketDcaSeries = {};
const USER_STATE_KEY = "decision-ledger-user-state-v1";
const usEtfSymbols = new Set(["SPY", "QQQ", "VOO", "VTI", "IVV", "SCHD", "VGT", "XLK", "SMH", "SOXX", "DIA", "IWM", "TLT", "BND", "AGG", "IBIT"]);
let userStateLoaded = false;
let singleTradeRenderToken = 0;
let compareRenderToken = 0;
let dcaRenderToken = 0;
let scenarioRenderToken = 0;

const sectorNames = {
  "01": "水泥",
  "02": "食品",
  "03": "塑膠",
  "04": "紡織",
  "05": "電機",
  "06": "電器電纜",
  "08": "玻璃",
  "09": "造紙",
  10: "鋼鐵",
  11: "橡膠",
  12: "汽車",
  14: "建材營造",
  15: "航運",
  16: "觀光",
  17: "金融",
  18: "貿易百貨",
  20: "其他",
  21: "化工",
  22: "生技",
  23: "油電燃氣",
  24: "半導體",
  25: "電腦週邊",
  26: "光電",
  27: "通信網路",
  28: "電子零組件",
  29: "電子通路",
  30: "資訊服務",
  31: "其他電子",
  32: "文化創意",
  33: "農業科技",
  35: "綠能環保",
  36: "數位雲端",
  37: "運動休閒",
  38: "居家生活",
};

const currency = new Intl.NumberFormat("zh-TW", {
  style: "currency",
  currency: "TWD",
  maximumFractionDigits: 0,
});

const compactCurrency = new Intl.NumberFormat("zh-TW", {
  notation: "compact",
  maximumFractionDigits: 1,
});

const percent = new Intl.NumberFormat("zh-TW", {
  style: "percent",
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});

const $ = (selector) => document.querySelector(selector);

function toNumber(value) {
  if (value === undefined || value === null) return 0;
  const parsed = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function rocDateToIso(value) {
  const text = String(value || "");
  if (text.length !== 7) return "";
  const year = Number(text.slice(0, 3)) + 1911;
  return `${year}-${text.slice(3, 5)}-${text.slice(5, 7)}`;
}

function addDaysIso(isoDate, days) {
  const date = new Date(`${isoDate}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function isoToTwseDate(isoDate) {
  return isoDate.replaceAll("-", "");
}

function isoToTpexRocDate(isoDate) {
  const [year, month, day] = isoDate.split("-");
  return `${Number(year) - 1911}/${month}/${day}`;
}

function formatMarketCap(value) {
  if (!Number.isFinite(value) || value <= 0) return "--";
  if (value >= 1_000_000_000_000) {
    return `${(value / 1_000_000_000_000).toFixed(value >= 10_000_000_000_000 ? 1 : 2)} 兆`;
  }
  return `${(value / 100_000_000).toFixed(value >= 100_000_000_000 ? 0 : 1)} 億`;
}

function formatVolume(value) {
  if (!Number.isFinite(value) || value <= 0) return "--";
  const lots = value >= 1_000_000 ? value / 1000 : value;
  return `${compactCurrency.format(lots)} 張`;
}

function isCompanyStock(symbol) {
  return /^[1-9]\d{3}$/.test(symbol);
}

function isTaiwanSymbol(symbol) {
  return /^([0-9]{4,6}[A-Z]?|[0-9]{4}[A-Z])$/.test(symbol);
}

function isTaiwanFundSymbol(symbol) {
  return /^(00|02)\d{2,4}[A-Z]?$/.test(symbol);
}

function isTaiwanRankingSymbol(symbol) {
  return (isCompanyStock(symbol) || isTaiwanFundSymbol(symbol)) && stocks[symbol]?.market !== "US";
}

function setMarketStatus(message, kind = "info") {
  const status = $("#marketStatus");
  if (!status) return;
  status.textContent = message;
  status.dataset.kind = kind;
}

function uniqueSymbols(symbols) {
  return [...new Set(symbols)].filter((symbol) => stocks[symbol]);
}

function visibleSymbols() {
  state.watchlist = uniqueSymbols(state.watchlist);
  return state.watchlist;
}

function optionHtml(symbols, selected = "") {
  if (symbols.length === 0) {
    return `<option value="">沒有可加入的股票</option>`;
  }
  return symbols
    .map((symbol) => `<option value="${symbol}" ${symbol === selected ? "selected" : ""}>${symbol} ${stocks[symbol].name}</option>`)
    .join("");
}

function resolveStockInput(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const firstToken = text.split(/\s+/)[0].toUpperCase();
  const exactCode = firstToken.match(/^[A-Z0-9._-]+/)?.[0]?.replace(".", "_").replace("-", "_");
  if (exactCode && stocks[exactCode]) return exactCode;
  if (exactCode && exactCode.length === 5 && stocks[`${exactCode}L`]) return `${exactCode}L`;
  if (exactCode && exactCode.length === 5 && stocks[`${exactCode}A`]) return `${exactCode}A`;
  const lowered = text.toLowerCase();
  const prefixMatches = Object.keys(stocks).filter((symbol) => symbol.toLowerCase().startsWith(lowered));
  if (prefixMatches.length === 1) return prefixMatches[0];
  if (prefixMatches.length > 1) {
    const leveragedOrPlain = prefixMatches.find((symbol) => /[A-Z]$/.test(symbol) || symbol.length === lowered.length);
    if (leveragedOrPlain) return leveragedOrPlain;
  }
  const match = Object.entries(stocks).find(
    ([symbol, stock]) =>
      symbol.toLowerCase() === lowered ||
      symbol.toLowerCase().startsWith(lowered) ||
      stock.name.toLowerCase().includes(lowered),
  );
  return match?.[0] || "";
}

function loadUserState() {
  try {
    const saved = JSON.parse(localStorage.getItem(USER_STATE_KEY) || "null");
    if (!saved) return;
    if (Array.isArray(saved.trades)) trades = saved.trades.filter((trade) => trade.id && stocks[trade.symbol]);
    ["watchlist", "portfolioSymbols", "compareSymbols", "dcaSymbols"].forEach((key) => {
      if (Array.isArray(saved[key])) state[key] = uniqueSymbols(saved[key]);
    });
    ["benchmark", "riskSymbol", "comparePeriod", "compareStartDate", "compareEndDate", "rankingPeriod", "sector", "dcaFrequency", "dcaStartDate", "dcaEndDate"].forEach((key) => {
      if (saved[key]) state[key] = saved[key];
    });
    if (Number.isFinite(Number(saved.dcaAmount))) state.dcaAmount = Number(saved.dcaAmount);
    if (Number.isFinite(Number(saved.portfolioAmount))) state.portfolioAmount = Number(saved.portfolioAmount);
    if (saved.portfolioWeights && typeof saved.portfolioWeights === "object") state.portfolioWeights = saved.portfolioWeights;
    if (saved.selectedTradeId) state.selectedTradeId = saved.selectedTradeId;
    if (state.comparePeriod === "all") {
      state.comparePeriod = "custom";
      state.compareStartDate = priceDates[0];
    }
  } catch {
    // Ignore corrupt local state and keep the default starter data.
  }
}

function saveUserState() {
  try {
    localStorage.setItem(
      USER_STATE_KEY,
      JSON.stringify({
        trades,
        watchlist: state.watchlist,
        portfolioSymbols: state.portfolioSymbols,
        portfolioAmount: state.portfolioAmount,
        portfolioWeights: state.portfolioWeights,
        compareSymbols: state.compareSymbols,
        dcaSymbols: state.dcaSymbols,
        benchmark: state.benchmark,
        selectedTradeId: state.selectedTradeId,
        comparePeriod: state.comparePeriod,
        compareStartDate: state.compareStartDate,
        compareEndDate: state.compareEndDate,
        rankingPeriod: state.rankingPeriod,
        sector: state.sector,
        dcaFrequency: state.dcaFrequency,
        dcaAmount: state.dcaAmount,
        dcaStartDate: state.dcaStartDate,
        dcaEndDate: state.dcaEndDate,
        riskSymbol: state.riskSymbol,
      }),
    );
  } catch {
    // localStorage may be disabled in some file:// contexts.
  }
}

function hydrateUserStateOnce() {
  if (userStateLoaded) return;
  loadUserState();
  userStateLoaded = true;
}

function syncSelections() {
  const list = visibleSymbols();
  if (!list.includes(state.benchmark)) state.benchmark = list[0] || "2330";
  if (!list.includes(state.riskSymbol)) state.riskSymbol = list[0] || "2330";
  state.compareSymbols = uniqueSymbols(state.compareSymbols.filter((symbol) => list.includes(symbol)));
  state.portfolioSymbols = uniqueSymbols(state.portfolioSymbols.filter((symbol) => list.includes(symbol)));
  state.dcaSymbols = uniqueSymbols(state.dcaSymbols.filter((symbol) => list.includes(symbol)));
}

function readCache() {
  try {
    const cache = JSON.parse(localStorage.getItem(CACHE_KEY) || "null");
    if (!cache?.date || !cache?.daily?.length || !cache?.companies?.length) return null;
    return cache;
  } catch {
    return null;
  }
}

function normalizeMarketPayload(payload) {
  if (!payload?.daily?.length || !payload?.companies?.length) return null;
  return {
    ...payload,
    date: payload.date?.includes("-") ? payload.date : rocDateToIso(payload.date || payload.daily[0]?.Date),
  };
}

function writeCache(payload) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify(payload));
  } catch {
    // localStorage may be disabled in some file:// contexts.
  }
}

function readHistoryCloseCache() {
  try {
    return JSON.parse(localStorage.getItem(HISTORY_CLOSE_CACHE_KEY) || "{}") || {};
  } catch {
    return {};
  }
}

function writeHistoryCloseCache(cache) {
  try {
    const entries = Object.entries(cache).sort(([a], [b]) => b.localeCompare(a)).slice(0, 40);
    localStorage.setItem(HISTORY_CLOSE_CACHE_KEY, JSON.stringify(Object.fromEntries(entries)));
  } catch {
    // localStorage may be disabled in some file:// contexts.
  }
}

function readAdjustedCloseCache() {
  try {
    return JSON.parse(localStorage.getItem(ADJUSTED_CLOSE_CACHE_KEY) || "{}") || {};
  } catch {
    return {};
  }
}

function writeAdjustedCloseCache(cache) {
  try {
    const entries = Object.entries(cache).slice(-200);
    localStorage.setItem(ADJUSTED_CLOSE_CACHE_KEY, JSON.stringify(Object.fromEntries(entries)));
  } catch {
    // localStorage may be disabled in some file:// contexts.
  }
}

function readAdjustedHistoryCache() {
  try {
    return JSON.parse(localStorage.getItem(ADJUSTED_HISTORY_CACHE_KEY) || "{}") || {};
  } catch {
    return {};
  }
}

function writeAdjustedHistoryCache(cache) {
  try {
    const entries = Object.entries(cache).slice(-80);
    localStorage.setItem(ADJUSTED_HISTORY_CACHE_KEY, JSON.stringify(Object.fromEntries(entries)));
  } catch {
    // localStorage may be disabled in some file:// contexts.
  }
}

function anchoredPriceSeries(symbol, close, fallbackPrices) {
  const existing = stocks[symbol]?.prices;
  if (!existing?.length) return fallbackPrices.slice(-priceDates.length);
  const anchor = existing.at(-1);
  if (!Number.isFinite(anchor) || anchor <= 0) return fallbackPrices.slice(-priceDates.length);
  const scale = close / anchor;
  const scaled = existing.map((price, index) => (index === existing.length - 1 ? close : price * scale));
  return scaled.slice(-priceDates.length);
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (cell || row.length) {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
      }
      if (char === "\r" && next === "\n") index += 1;
    } else {
      cell += char;
    }
  }
  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }
  const headers = rows.shift() || [];
  return rows.map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] || ""])));
}

function parseTwseHistory(json) {
  const table = json?.tables?.find((item) => item.fields?.includes("證券代號") && item.fields?.includes("收盤價"));
  if (!table?.data?.length) return {};
  const codeIndex = table.fields.indexOf("證券代號");
  const closeIndex = table.fields.indexOf("收盤價");
  return Object.fromEntries(
    table.data
      .map((row) => [String(row[codeIndex] || "").trim(), toNumber(row[closeIndex])])
      .filter(([symbol, close]) => symbol && close > 0),
  );
}

function parseTpexHistory(json) {
  const table = json?.tables?.[0];
  if (!table?.data?.length) return {};
  const codeIndex = table.fields.findIndex((field) => field.trim() === "代號");
  const closeIndex = table.fields.findIndex((field) => field.trim() === "收盤");
  if (codeIndex === -1 || closeIndex === -1) return {};
  return Object.fromEntries(
    table.data
      .map((row) => [String(row[codeIndex] || "").trim(), toNumber(row[closeIndex])])
      .filter(([symbol, close]) => symbol && close > 0),
  );
}

async function fetchJson(url) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 12_000);
  try {
    const response = await fetch(url, { cache: "no-store", signal: controller.signal });
    if (!response.ok) throw new Error(`Request failed: ${response.status}`);
    return response.json();
  } finally {
    clearTimeout(timeout);
  }
}

async function fetchTwseHistory(date) {
  const params = new URLSearchParams({ date: isoToTwseDate(date), type: "ALLBUT0999", response: "json" });
  return parseTwseHistory(await fetchJson(`${TWSE_HISTORY_URL}?${params}`));
}

async function fetchTpexHistory(date) {
  const params = new URLSearchParams({ l: "zh-tw", d: isoToTpexRocDate(date), se: "EW", o: "json" });
  return parseTpexHistory(await fetchJson(`${TPEX_HISTORY_URL}?${params}`));
}

async function historicalCloseSnapshotOnOrBefore(targetDate) {
  const bundledSnapshot = Object.values({ ...marketHistoryPeriods, ...marketDcaSnapshots })
    .filter((snapshot) => snapshot?.date && snapshot.date <= targetDate && snapshot.date >= addDaysIso(targetDate, -14) && snapshot.closes)
    .sort((a, b) => b.date.localeCompare(a.date))[0];
  if (bundledSnapshot) return bundledSnapshot;

  const cache = readHistoryCloseCache();
  for (let offset = 0; offset <= 14; offset += 1) {
    const date = addDaysIso(targetDate, -offset);
    if (cache[date]?.closes) return cache[date];

    const [twseResult, tpexResult] = await Promise.allSettled([fetchTwseHistory(date), fetchTpexHistory(date)]);
    const twse = twseResult.status === "fulfilled" ? twseResult.value : {};
    const tpex = tpexResult.status === "fulfilled" ? tpexResult.value : {};
    const closes = { ...twse, ...tpex };
    if (Object.keys(closes).length > 100) {
      const snapshot = { date, closes, fetchedAt: new Date().toISOString() };
      cache[date] = snapshot;
      writeHistoryCloseCache(cache);
      return snapshot;
    }
  }
  return null;
}

function taiwanDateFromTimestamp(timestamp) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date(timestamp * 1000));
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function yahooSymbolFor(symbol) {
  if (stocks[symbol]?.market === "US") return symbol.replace("_", "-");
  if (!isTaiwanSymbol(symbol)) return "";
  return `${symbol}.${stocks[symbol]?.market === "TPEX" ? "TWO" : "TW"}`;
}

async function adjustedCloseOnOrBefore(symbol, targetDate) {
  const yahooSymbol = yahooSymbolFor(symbol);
  if (!yahooSymbol) return null;

  const cache = readAdjustedCloseCache();
  const cacheKey = `${symbol}:${targetDate}`;
  if (cache[cacheKey]) return cache[cacheKey];

  const start = Math.floor(new Date(`${addDaysIso(targetDate, -7)}T00:00:00Z`).getTime() / 1000);
  const end = Math.floor(new Date(`${addDaysIso(targetDate, 2)}T00:00:00Z`).getTime() / 1000);
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooSymbol)}?period1=${start}&period2=${end}&interval=1d&events=history%7Cdiv%7Csplit`;
  const result = await fetchJson(url);
  const chart = result?.chart?.result?.[0];
  const timestamps = chart?.timestamp || [];
  const adjusted = chart?.indicators?.adjclose?.[0]?.adjclose || [];
  const closes = chart?.indicators?.quote?.[0]?.close || [];
  const points = timestamps
    .map((timestamp, index) => ({
      date: taiwanDateFromTimestamp(timestamp),
      value: toNumber(adjusted[index]) || toNumber(closes[index]),
    }))
    .filter((point) => point.date <= targetDate && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  const point = points.at(-1) || null;
  if (point) {
    cache[cacheKey] = { ...point, source: "adjusted" };
    writeAdjustedCloseCache(cache);
    return cache[cacheKey];
  }
  return null;
}

async function adjustedHistoryFor(symbol, startDate, endDate) {
  const yahooSymbol = yahooSymbolFor(symbol);
  if (!yahooSymbol) return [];

  const cache = readAdjustedHistoryCache();
  const cacheKey = `${symbol}:${startDate}:${endDate}`;
  if (Array.isArray(cache[cacheKey])) return cache[cacheKey];

  const start = Math.floor(new Date(`${addDaysIso(startDate, -10)}T00:00:00Z`).getTime() / 1000);
  const end = Math.floor(new Date(`${addDaysIso(endDate, 3)}T00:00:00Z`).getTime() / 1000);
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooSymbol)}?period1=${start}&period2=${end}&interval=1d&events=history%7Cdiv%7Csplit`;
  const result = await fetchJson(url);
  const chart = result?.chart?.result?.[0];
  const timestamps = chart?.timestamp || [];
  const adjusted = chart?.indicators?.adjclose?.[0]?.adjclose || [];
  const closes = chart?.indicators?.quote?.[0]?.close || [];
  const points = timestamps
    .map((timestamp, index) => ({
      date: taiwanDateFromTimestamp(timestamp),
      value: toNumber(adjusted[index]) || toNumber(closes[index]),
    }))
    .filter((point) => point.date >= addDaysIso(startDate, -10) && point.date <= addDaysIso(endDate, 3) && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));

  cache[cacheKey] = points;
  writeAdjustedHistoryCache(cache);
  return points;
}

function pricePointOnOrAfter(points, date) {
  return points.find((point) => point.date >= date) || points.at(-1) || null;
}

function pricePointOnOrBefore(points, date) {
  for (let index = points.length - 1; index >= 0; index -= 1) {
    if (points[index].date <= date) return points[index];
  }
  return points[0] || null;
}

async function mapWithConcurrency(items, limit, mapper) {
  const results = new Array(items.length);
  let nextIndex = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await mapper(items[currentIndex], currentIndex);
    }
  });
  await Promise.all(workers);
  return results;
}

function latestTradingDate() {
  const date = new Date();
  date.setHours(date.getHours() - 18);
  while (date.getDay() === 0 || date.getDay() === 6) {
    date.setDate(date.getDate() - 1);
  }
  return date.toISOString().slice(0, 10);
}

function isFreshMarketPayload(payload) {
  if (!payload?.date) return false;
  const payloadDate = payload.date.includes("-") ? payload.date : rocDateToIso(payload.date);
  return payloadDate >= latestTradingDate();
}

async function fetchMarketData() {
  const [dailyResponse, companyResponse, tpexResponse] = await Promise.all([
    fetch(TWSE_DAILY_URL, { cache: "no-store" }),
    fetch(TWSE_COMPANY_URL, { cache: "no-store" }),
    fetch(TPEX_DAILY_URL, { cache: "no-store" }),
  ]);
  if (!dailyResponse.ok || !companyResponse.ok || !tpexResponse.ok) {
    throw new Error("TWSE API request failed");
  }
  const daily = await dailyResponse.json();
  const companies = await companyResponse.json();
  const tpexDaily = parseCsv(await tpexResponse.text());
  return { daily, companies, tpexDaily, fetchedAt: new Date().toISOString(), date: rocDateToIso(daily[0]?.Date) };
}

function applyTwseMarketData(payload) {
  const companyMap = new Map(payload.companies.map((item) => [item["公司代號"], item]));
  marketHistoryPeriods = payload.history?.periods || {};
  marketDcaSnapshots = payload.history?.dcaMonthly || {};
  marketDcaSeries = payload.history?.dcaSeries || {};
  let added = 0;

  payload.daily.forEach((row) => {
    const symbol = row.Code;
    const close = toNumber(row.ClosingPrice);
    if (!symbol || !close || !isTaiwanSymbol(symbol)) return;

    const company = companyMap.get(symbol);
    const issuedShares = toNumber(company?.["已發行普通股數或TDR原股發行股數"]);
    const sector = isTaiwanFundSymbol(symbol)
      ? "ETF / ETN"
      : sectorNames[company?.["產業別"]] || stocks[symbol]?.sector || "上市公司";
    const previous = close - toNumber(row.Change);
    const fallbackPrices = [
      previous * 0.76,
      previous * 0.82,
      previous * 0.88,
      previous * 0.92,
      previous * 0.97,
      previous,
      close,
      close * 0.99,
      close * 1.01,
      previous,
      close,
    ];
    const generatedPrices = anchoredPriceSeries(symbol, close, fallbackPrices);

    stocks[symbol] = {
      name: company?.["公司簡稱"] || row.Name,
      market: "TWSE",
      sector,
      volume: toNumber(row.TradeVolume),
      marketCap: issuedShares ? issuedShares * close : stocks[symbol]?.marketCap || 0,
      issuedShares,
      live: true,
      dailyChange: toNumber(row.Change),
      dailyReturn: previous > 0 ? close / previous - 1 : 0,
      periodReturns: historicalReturnsFor(symbol, close),
      tradeValue: toNumber(row.TradeValue),
      prices: generatedPrices,
    };
    added += 1;
  });

  state.marketDataSource = "twse";
  state.marketDataDate = payload.history?.latestDate || (String(payload.date || "").includes("-") ? payload.date : rocDateToIso(payload.date)) || payload.fetchedAt.slice(0, 10);
  if (state.marketDataDate) {
    const previousLastDate = priceDates.at(-1);
    priceDates[priceDates.length - 1] = state.marketDataDate;
    if (!userStateLoaded && state.compareEndDate >= previousLastDate) state.compareEndDate = state.marketDataDate;
    if (!userStateLoaded && state.dcaEndDate >= previousLastDate) state.dcaEndDate = state.marketDataDate;
  }
  return added;
}

function applyTpexMarketData(payload) {
  let added = 0;
  (payload.tpexDaily || []).forEach((row) => {
    const symbol = row["代號"];
    const close = toNumber(row["收盤"]);
    if (!symbol || !close) return;

    const previous = close - toNumber(row["漲跌"]);
    const issuedShares = toNumber(row["發行股數"]);
    const isEtfLike = symbol.startsWith("00") || symbol.startsWith("02");
    stocks[symbol] = {
      name: row["名稱"],
      market: "TPEX",
      sector: isEtfLike ? "ETF / ETN" : stocks[symbol]?.sector || "上櫃",
      volume: toNumber(row["成交股數"]),
      marketCap: issuedShares * close,
      issuedShares,
      live: true,
      dailyChange: toNumber(row["漲跌"]),
      dailyReturn: previous > 0 ? close / previous - 1 : 0,
      periodReturns: historicalReturnsFor(symbol, close),
      tradeValue: toNumber(row["成交金額"]),
      prices: anchoredPriceSeries(symbol, close, [
        previous * 0.78,
        previous * 0.83,
        previous * 0.88,
        previous * 0.93,
        previous * 0.97,
        previous,
        close,
        close * 0.99,
        close * 1.01,
        previous,
        close,
      ]),
    };
    added += 1;
  });
  return added;
}

function applyUsMarketData(payload) {
  let added = 0;
  (payload.usDaily || []).forEach((row) => {
    const symbol = row.symbol;
    const close = toNumber(row.close);
    const open = toNumber(row.open) || close;
    if (!symbol || !close) return;
    const previous = open;
    stocks[symbol] = {
      name: row.name || symbol,
      market: "US",
      sector: usEtfSymbols.has(symbol) ? "US ETF" : "US Stock",
      volume: toNumber(row.volume),
      marketCap: 0,
      live: true,
      dailyChange: close - previous,
      dailyReturn: previous > 0 ? close / previous - 1 : 0,
      tradeValue: toNumber(row.volume) * close,
      prices: anchoredPriceSeries(symbol, close, [
        close * 0.72,
        close * 0.78,
        close * 0.84,
        close * 0.9,
        close * 0.96,
        previous,
        close * 0.98,
        close * 1.02,
        close * 0.99,
        previous,
        close,
      ]),
    };
    added += 1;
  });
  return added;
}

async function loadDailyMarketData() {
  const today = new Date().toISOString().slice(0, 10);
  const bundled = normalizeMarketPayload(window.TWSE_MARKET_DATA);
  if (bundled?.fetchedAt && isFreshMarketPayload(bundled)) {
    const added = applyTwseMarketData(bundled);
    const tpexAdded = applyTpexMarketData(bundled);
    const usAdded = applyUsMarketData(bundled);
    hydrateUserStateOnce();
    renderControls();
    render();
    setMarketStatus(`已載入本地每日資料 ${state.marketDataDate}，TWSE ${added} 檔、TPEx ${tpexAdded} 檔、美股 ${usAdded} 檔；排程會每天收盤後更新。`, "ok");
    return;
  }

  const cache = readCache();
  if (cache?.cachedAt === today) {
    const added = applyTwseMarketData(cache);
    const tpexAdded = applyTpexMarketData(cache);
    const usAdded = applyUsMarketData(cache);
    hydrateUserStateOnce();
    renderControls();
    render();
    setMarketStatus(`已載入每日快取 ${state.marketDataDate}，TWSE ${added} 檔、TPEx ${tpexAdded} 檔、美股 ${usAdded} 檔。`, "ok");
    return;
  }

  try {
    setMarketStatus("正在從證交所與櫃買中心取得每日收盤行情...");
    const payload = await fetchMarketData();
    const cachePayload = { ...payload, usDaily: bundled?.usDaily || cache?.usDaily || [], cachedAt: today };
    writeCache(cachePayload);
    const added = applyTwseMarketData(cachePayload);
    const tpexAdded = applyTpexMarketData(cachePayload);
    const usAdded = applyUsMarketData(cachePayload);
    hydrateUserStateOnce();
    renderControls();
    render();
    setMarketStatus(`已更新每日資料 ${state.marketDataDate}，TWSE ${added} 檔、TPEx ${tpexAdded} 檔；美股 ${usAdded} 檔沿用本地快照，會由每日更新腳本刷新。`, "ok");
  } catch (error) {
    console.warn(error);
    if (bundled?.fetchedAt) {
      const added = applyTwseMarketData(bundled);
      const tpexAdded = applyTpexMarketData(bundled);
      const usAdded = applyUsMarketData(bundled);
      hydrateUserStateOnce();
      renderControls();
      render();
      setMarketStatus(`即時資料暫時無法載入，先使用本地快照 ${state.marketDataDate}：TWSE ${added} 檔、TPEx ${tpexAdded} 檔、美股 ${usAdded} 檔。`, "warn");
      return;
    }
    hydrateUserStateOnce();
    setMarketStatus("每日資料暫時無法載入，已改用內建樣本；重新整理時會再自動嘗試。", "warn");
  }
}

function seriesFor(symbol) {
  return priceDates.map((date, index) => [date, stocks[symbol].prices[index]]);
}

function priceOnOrAfter(symbol, date) {
  const index = priceDates.findIndex((day) => day >= date);
  const safeIndex = index === -1 ? priceDates.length - 1 : index;
  return stocks[symbol].prices[safeIndex];
}

function previousClose(symbol) {
  const stock = stocks[symbol];
  const dailyReturn = stock?.dailyReturn;
  const latest = latestPrice(symbol);
  if (!stock?.live || !Number.isFinite(dailyReturn) || dailyReturn <= -1) return 0;
  return latest / (1 + dailyReturn);
}

function priceOnOrBefore(symbol, date) {
  if (stocks[symbol]?.live && state.marketDataDate && date < state.marketDataDate) {
    const lastKnownHistoricalDate = priceDates.at(-2);
    if (!lastKnownHistoricalDate || date >= lastKnownHistoricalDate) {
      const previous = previousClose(symbol);
      if (previous > 0) return previous;
    }
  }
  return priceAtIndex(symbol, dateIndexOnOrBefore(date));
}

function entryPriceForTrade(trade, symbol = trade.symbol) {
  if (symbol === trade.symbol && Number.isFinite(Number(trade.price))) return Number(trade.price);
  return priceOnOrBefore(symbol, trade.date);
}

function priceAtIndex(symbol, index) {
  return stocks[symbol].prices[Math.max(0, Math.min(index, priceDates.length - 1))];
}

function dateIndexOnOrAfter(date) {
  const index = priceDates.findIndex((day) => day >= date);
  return index === -1 ? priceDates.length - 1 : index;
}

function dateIndexOnOrBefore(date) {
  for (let index = priceDates.length - 1; index >= 0; index -= 1) {
    if (priceDates[index] <= date) return index;
  }
  return 0;
}

function nearestDateIndex(date) {
  const target = new Date(`${date}T00:00:00`).getTime();
  return priceDates.reduce((bestIndex, day, index) => {
    const bestDistance = Math.abs(new Date(`${priceDates[bestIndex]}T00:00:00`).getTime() - target);
    const distance = Math.abs(new Date(`${day}T00:00:00`).getTime() - target);
    return distance < bestDistance ? index : bestIndex;
  }, 0);
}

function latestPrice(symbol) {
  return stocks[symbol].prices.at(-1);
}

function compareWindow(period = "1y") {
  if (period === "custom") {
    const start = dateIndexOnOrAfter(state.compareStartDate);
    const end = Math.max(start, dateIndexOnOrBefore(state.compareEndDate));
    return { start, end, label: "自訂" };
  }
  const end = priceDates.length - 1;
  const start = Math.max(0, end - (periodSteps[period] ?? 1));
  return { start, end, label: $("#comparePeriod")?.selectedOptions?.[0]?.textContent || period };
}

function periodReturnDetail(symbol, period = "1y") {
  if ((period === "today" || period === "1d") && stocks[symbol]?.live) {
    const endPrice = latestPrice(symbol);
    const returnRate = stocks[symbol].dailyReturn || 0;
    const startPrice = returnRate === -1 ? endPrice : endPrice / (1 + returnRate);
    return {
      symbol,
      startDate: "前一交易日",
      endDate: state.marketDataDate || priceDates.at(-1),
      startPrice,
      endPrice,
      returnRate,
    };
  }
  const { start, end } = compareWindow(period);
  const startPrice = priceAtIndex(symbol, start);
  const endPrice = priceAtIndex(symbol, end);
  return {
    symbol,
    startDate: priceDates[start],
    endDate: priceDates[end],
    startPrice,
    endPrice,
    returnRate: startPrice ? endPrice / startPrice - 1 : 0,
  };
}

function periodReturn(symbol, period = "1y") {
  return periodReturnDetail(symbol, period).returnRate;
}

function historicalReturnsFor(symbol, close) {
  return Object.fromEntries(
    ["1w", "1m", "1y"].map((period) => {
      const baseClose = toNumber(marketHistoryPeriods[period]?.closes?.[symbol]);
      return [period, baseClose > 0 ? close / baseClose - 1 : null];
    }),
  );
}

function isRankingPeriodAvailable(period) {
  if (period === "today" || period === "1d") return true;
  return Object.values(stocks).some((stock) => Number.isFinite(stock.periodReturns?.[period]));
}

function normalizeRankingPeriod() {
  if (!isRankingPeriodAvailable(state.rankingPeriod)) state.rankingPeriod = "today";
}

function rankingReturn(symbol, period) {
  if (period === "today" || period === "1d") return stocks[symbol]?.dailyReturn || 0;
  const value = stocks[symbol]?.periodReturns?.[period];
  return Number.isFinite(value) ? value : null;
}

function rankingPeriodLabel(period) {
  if (period === "today") return "今日";
  if (period === "1d") return "1日";
  if (period === "1w") return "1週";
  if (period === "1m") return "1月";
  if (period === "1y") return "1年";
  return period;
}

function updateRankingPeriodOptions() {
  [...$("#rankingPeriod").options].forEach((option) => {
    const available = isRankingPeriodAvailable(option.value);
    option.disabled = !available;
    option.textContent = available ? rankingPeriodLabel(option.value) : `${rankingPeriodLabel(option.value)}（需歷史資料）`;
  });
}

function allPeriodReturns(symbol) {
  return ["1d", "1w", "1m", "1y"].map((period) => periodReturn(symbol, period));
}

function comparePeriodDates(period = state.comparePeriod) {
  if (period === "custom") {
    return { startTarget: state.compareStartDate, endTarget: state.compareEndDate, label: "自訂時間" };
  }
  const endTarget = state.marketDataDate || priceDates.at(-1);
  const lookbackDays = { "1d": 1, "1w": 7, "1m": 30, "1y": 365 }[period] || 30;
  return {
    startTarget: addDaysIso(endTarget, -lookbackDays),
    endTarget,
    label: $("#comparePeriod")?.selectedOptions?.[0]?.textContent || period,
  };
}

function closeFromHistoricalSnapshot(symbol, snapshot) {
  return toNumber(snapshot?.closes?.[symbol]);
}

async function compareSnapshotFor(date) {
  if (state.marketDataDate && date >= state.marketDataDate) {
    return { date: state.marketDataDate, latest: true };
  }
  return historicalCloseSnapshotOnOrBefore(date);
}

function compareClose(symbol, snapshot) {
  if (snapshot?.latest) return latestPrice(symbol);
  return closeFromHistoricalSnapshot(symbol, snapshot);
}

function sharesFor(trade, symbol = trade.symbol) {
  if (symbol === trade.symbol && Number.isFinite(Number(trade.shares))) return Number(trade.shares);
  const entryPrice = entryPriceForTrade(trade, symbol);
  return trade.amount / entryPrice;
}

function currentValueFor(trade, symbol = trade.symbol) {
  return sharesFor(trade, symbol) * latestPrice(symbol);
}

function returnFor(trade, symbol = trade.symbol) {
  const invested = investedAmountFor(trade);
  return invested ? currentValueFor(trade, symbol) / invested - 1 : 0;
}

function hypotheticalTradeResult(trade, symbol, entryPoint) {
  const invested = investedAmountFor(trade);
  const entryPrice = toNumber(entryPoint?.value);
  if (!(invested > 0) || !(entryPrice > 0)) {
    return { symbol, entryPrice: 0, value: 0, returnRate: null, available: false };
  }
  const shares = invested / entryPrice;
  const value = shares * latestPrice(symbol);
  return {
    symbol,
    entryPrice,
    entryDate: entryPoint.date,
    entrySource: entryPoint.source,
    value,
    returnRate: invested ? value / invested - 1 : null,
    available: true,
  };
}

async function entryPointForScenario(symbol, targetDate) {
  const bundledPoint = pricePointOnOrBefore(marketDcaSeries[symbol]?.points || [], targetDate);
  if (bundledPoint?.value > 0 && bundledPoint.date <= targetDate) return { date: bundledPoint.date, value: bundledPoint.value, source: "adjusted" };
  try {
    const adjusted = await adjustedCloseOnOrBefore(symbol, targetDate);
    if (adjusted?.value > 0) return adjusted;
  } catch (error) {
    console.warn(`Scenario adjusted close unavailable for ${symbol}: ${error.message}`);
  }
  return null;
}

async function scenarioTotals(symbol) {
  let invested = 0;
  let value = 0;
  let countedTrades = 0;

  for (const trade of trades) {
    const amount = investedAmountFor(trade);
    if (!(amount > 0)) continue;

    if (symbol === trade.symbol && Number.isFinite(Number(trade.shares))) {
      invested += amount;
      value += Number(trade.shares) * latestPrice(symbol);
      countedTrades += 1;
      continue;
    }

    const entryPoint = await entryPointForScenario(symbol, trade.date);
    const entryPrice = toNumber(entryPoint?.value);
    if (!(entryPrice > 0)) continue;
    invested += amount;
    value += (amount / entryPrice) * latestPrice(symbol);
    countedTrades += 1;
  }

  return { invested, value, countedTrades, missingTrades: trades.length - countedTrades, returnRate: invested ? value / invested - 1 : null };
}

function actualTradeResult(trade) {
  const invested = investedAmountFor(trade);
  const entryPrice = entryPriceForTrade(trade);
  const value = currentValueFor(trade);
  return {
    symbol: trade.symbol,
    entryPrice,
    entryDate: trade.date,
    entrySource: "actual",
    value,
    returnRate: invested ? value / invested - 1 : null,
    available: true,
  };
}

async function entryPointForHypothetical(symbol, targetDate, snapshot) {
  const bundledPoint = pricePointOnOrBefore(marketDcaSeries[symbol]?.points || [], targetDate);
  if (bundledPoint?.value > 0 && bundledPoint.date <= targetDate) return { date: bundledPoint.date, value: bundledPoint.value, source: "adjusted" };
  try {
    const adjusted = await adjustedCloseOnOrBefore(symbol, targetDate);
    if (adjusted?.value > 0) return adjusted;
  } catch (error) {
    console.warn(`Adjusted close unavailable for ${symbol}: ${error.message}`);
  }
  const raw = toNumber(snapshot?.closes?.[symbol]);
  return raw > 0 ? { date: snapshot.date, value: raw, source: "raw" } : null;
}

function classForReturn(value) {
  if (!Number.isFinite(value)) return "";
  return value >= 0 ? "return-positive" : "return-negative";
}

function buildEquitySeries(symbolMode) {
  return priceDates.map((date, index) => {
    const value = trades.reduce((sum, trade) => {
      if (trade.date > date) return sum;
      const symbol = symbolMode === "real" ? trade.symbol : symbolMode;
      const currentPrice = priceAtIndex(symbol, index);
      return sum + sharesFor(trade, symbol) * currentPrice;
    }, 0);
    return { date, value };
  });
}

function marketSeriesPrice(symbol, date) {
  const bundledPoint = pricePointOnOrBefore(marketDcaSeries[symbol]?.points || [], date);
  if (bundledPoint?.value > 0 && bundledPoint.date <= date) return bundledPoint.value;
  return priceOnOrBefore(symbol, date);
}

function entryPriceForScenarioSync(symbol, trade) {
  if (symbol === trade.symbol && Number.isFinite(Number(trade.price))) return Number(trade.price);
  const bundledPoint = pricePointOnOrBefore(marketDcaSeries[symbol]?.points || [], trade.date);
  if (bundledPoint?.value > 0 && bundledPoint.date <= trade.date) return bundledPoint.value;
  return entryPriceForTrade(trade, symbol);
}

function buildScenarioEquitySeries(symbolMode = "real") {
  const dcaDates = Object.values(marketDcaSeries)
    .flatMap((history) => (history.points || []).map((point) => point.date))
    .filter((date) => trades.some((trade) => trade.date <= date));
  const dates = [...new Set([...priceDates, ...Object.values(marketHistoryPeriods).map((snapshot) => snapshot.date).filter(Boolean), ...dcaDates, state.marketDataDate].filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));

  return dates
    .map((date) => {
      const value = trades.reduce((sum, trade) => {
        if (trade.date > date) return sum;
        const symbol = symbolMode === "real" ? trade.symbol : symbolMode;
        const entryPrice = entryPriceForScenarioSync(symbol, trade);
        const currentPrice = marketSeriesPrice(symbol, date);
        if (!(entryPrice > 0) || !(currentPrice > 0)) return sum;
        const shares = symbolMode === "real" && symbol === trade.symbol && Number.isFinite(Number(trade.shares))
          ? Number(trade.shares)
          : investedAmountFor(trade) / entryPrice;
        return sum + shares * currentPrice;
      }, 0);
      return { date, value };
    })
    .filter((point) => point.value > 0);
}

function returnsFromValues(series) {
  return series.slice(1).map((point, index) => {
    const previous = series[index].value;
    return previous > 0 ? point.value / previous - 1 : 0;
  });
}

function maxDrawdown(series) {
  let peak = 0;
  let worst = 0;
  for (const point of series) {
    peak = Math.max(peak, point.value);
    if (peak > 0) worst = Math.min(worst, point.value / peak - 1);
  }
  return worst;
}

function volatility(series) {
  const returns = returnsFromValues(series).filter(Number.isFinite);
  if (returns.length === 0) return 0;
  const avg = returns.reduce((sum, value) => sum + value, 0) / returns.length;
  const variance = returns.reduce((sum, value) => sum + (value - avg) ** 2, 0) / returns.length;
  return Math.sqrt(variance);
}

function winRate(series) {
  const returns = returnsFromValues(series);
  if (returns.length === 0) return 0;
  return returns.filter((value) => value > 0).length / returns.length;
}

function sharpeRatio(series) {
  const returns = returnsFromValues(series);
  if (returns.length === 0) return 0;
  const avg = returns.reduce((sum, value) => sum + value, 0) / returns.length;
  const vol = volatility(series);
  return vol === 0 ? 0 : avg / vol;
}

function snapshotPricePoints(symbol) {
  const points = Object.values(marketHistoryPeriods)
    .filter((snapshot) => snapshot?.date && toNumber(snapshot.closes?.[symbol]) > 0)
    .map((snapshot) => ({ date: snapshot.date, value: toNumber(snapshot.closes[symbol]), source: "history" }));
  const currentDate = state.marketDataDate || priceDates.at(-1);
  const current = latestPrice(symbol);
  if (currentDate && current > 0) points.push({ date: currentDate, value: current, source: "current" });

  return [...new Map(points.sort((a, b) => a.date.localeCompare(b.date)).map((point) => [point.date, point])).values()];
}

function closeFromSnapshot(symbol, point) {
  if (point.source === "current") return latestPrice(symbol);
  return toNumber(point.closes?.[symbol]);
}

function portfolioSnapshotSeries(symbolMode = "real") {
  const snapshots = Object.values(marketHistoryPeriods)
    .filter((snapshot) => snapshot?.date)
    .map((snapshot) => ({ date: snapshot.date, source: "history", closes: snapshot.closes || {} }));
  const currentDate = state.marketDataDate || priceDates.at(-1);
  if (currentDate) snapshots.push({ date: currentDate, source: "current", closes: {} });

  return [...new Map(snapshots.sort((a, b) => a.date.localeCompare(b.date)).map((point) => [point.date, point])).values()]
    .map((point) => {
      const value = trades.reduce((sum, trade) => {
        if (trade.date > point.date) return sum;
        const symbol = symbolMode === "real" ? trade.symbol : symbolMode;
        const price = closeFromSnapshot(symbol, point);
        if (!(price > 0)) return sum;
        return sum + sharesFor(trade, symbol) * price;
      }, 0);
      return { date: point.date, value };
    })
    .filter((point) => point.value > 0);
}

function hasEnoughSeries(series) {
  return series.length >= 3;
}

function metricValue(value, formatter = percent) {
  return Number.isFinite(value) ? formatter.format(value) : "--";
}

function investedAmountFor(trade) {
  if (Number.isFinite(Number(trade.amount))) return Number(trade.amount);
  return sharesFor(trade) * entryPriceForTrade(trade);
}

function summarizePortfolio() {
  const totalInvested = trades.reduce((sum, trade) => sum + investedAmountFor(trade), 0);
  const currentValue = trades.reduce((sum, trade) => sum + currentValueFor(trade), 0);
  const series = buildEquitySeries("real");
  return {
    totalInvested,
    currentValue,
    totalReturn: totalInvested ? currentValue / totalInvested - 1 : 0,
    drawdown: maxDrawdown(series),
  };
}

function portfolioWeights() {
  const inputs = [...document.querySelectorAll("[data-weight]")];
  return inputs.map((input) => ({ symbol: input.dataset.weight, weight: Number(input.value) / 100 || 0 }));
}

function syncPortfolioWeightsFromInputs() {
  state.portfolioWeights = Object.fromEntries(
    [...document.querySelectorAll("[data-weight]")].map((input) => [input.dataset.weight, Number(input.value) || 0]),
  );
}

function buildWeightedPortfolioSeries(amount) {
  const weights = portfolioWeights();
  return priceDates.map((date, index) => {
    const value = weights.reduce((sum, item) => {
      const startPrice = priceAtIndex(item.symbol, 0);
      const currentPrice = priceAtIndex(item.symbol, index);
      return sum + amount * item.weight * (currentPrice / startPrice);
    }, 0);
    return { date, value };
  });
}

function buildDcaSeries(symbol = state.dcaSymbols[0], startDate = state.dcaStartDate, endDate = state.dcaEndDate) {
  let invested = 0;
  let shares = 0;
  const installments = buildDcaInstallments(startDate, endDate, state.dcaFrequency);
  let installmentIndex = 0;
  return dcaValuationPoints(startDate, endDate).map(({ date, index }) => {
    while (installmentIndex < installments.length && installments[installmentIndex] <= date) {
      invested += state.dcaAmount;
      shares += state.dcaAmount / priceAtIndex(symbol, nearestDateIndex(installments[installmentIndex]));
      installmentIndex += 1;
    }
    return {
      date,
      invested,
      value: shares * priceAtIndex(symbol, index),
    };
  });
}

function buildDcaSeriesFromPricePoints(symbol, points, startDate, endDate) {
  const usablePoints = points
    .filter((point) => point.date >= addDaysIso(startDate, -14) && point.date <= endDate && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  if (usablePoints.length === 0) {
    return { symbol, series: [], cashflows: [], installments: [], error: "這個日期區間沒有可用日線" };
  }

  const schedule = buildDcaInstallments(startDate, endDate, state.dcaFrequency);
  const installments = schedule
    .map((scheduledDate) => {
      const point = pricePointOnOrBefore(usablePoints, scheduledDate);
      return point?.value > 0 ? { scheduledDate, date: point.date, price: point.value } : null;
    })
    .filter((item) => item && item.date <= endDate)
    .sort((a, b) => a.date.localeCompare(b.date));
  const finalPoint = pricePointOnOrBefore(usablePoints, endDate);
  if (installments.length === 0 || !finalPoint?.value) {
    return { symbol, series: [], cashflows: [], installments: [], error: "投入日或結束日沒有可用價格" };
  }

  let invested = 0;
  let shares = 0;
  let installmentIndex = 0;
  const cashflows = [];
  const valuationPoints = usablePoints.filter((point) => point.date >= installments[0].date && point.date <= finalPoint.date);
  const series = valuationPoints
    .map((point) => {
      while (installmentIndex < installments.length && installments[installmentIndex].date <= point.date) {
        const installment = installments[installmentIndex];
        invested += state.dcaAmount;
        shares += state.dcaAmount / installment.price;
        cashflows.push({ date: installment.date, amount: -state.dcaAmount });
        installmentIndex += 1;
      }
      return {
        date: point.date,
        invested,
        value: shares * point.value,
      };
    })
    .filter((point) => point.invested > 0);

  const last = series.at(-1);
  if (last) cashflows.push({ date: last.date, amount: last.value });
  return { symbol, series, cashflows, installments, error: series.length ? "" : "日期區間沒有產生估值點" };
}

function buildDcaSeriesFromSnapshots(symbol, snapshotRequests, endDate) {
  if (snapshotRequests.length === 0) return { symbol, series: [], cashflows: [], installments: [], error: "沒有可用官方收盤資料" };

  const installments = snapshotRequests
    .filter((request) => request.kind === "installment")
    .map((request) => {
      const price = closeFromHistoricalSnapshot(symbol, request.snapshot);
      return price > 0 ? { scheduledDate: request.targetDate, date: request.snapshot.date, price } : null;
    })
    .filter((item) => item && item.date <= endDate);
  if (installments.length === 0) return { symbol, series: [], cashflows: [], installments: [], error: "投入日沒有這檔標的價格" };

  const valuationSnapshots = [
    ...new Map(snapshotRequests.map((request) => [request.snapshot.date, request.snapshot])).values(),
  ].sort((a, b) => a.date.localeCompare(b.date));
  let invested = 0;
  let shares = 0;
  let installmentIndex = 0;
  const cashflows = [];
  const series = valuationSnapshots.map((snapshot) => {
    while (installmentIndex < installments.length && installments[installmentIndex].date <= snapshot.date) {
      const installment = installments[installmentIndex];
      invested += state.dcaAmount;
      shares += state.dcaAmount / installment.price;
      cashflows.push({ date: installment.date, amount: -state.dcaAmount });
      installmentIndex += 1;
    }
    const price = closeFromHistoricalSnapshot(symbol, snapshot);
    return {
      date: snapshot.date,
      invested,
      value: price > 0 ? shares * price : 0,
    };
  }).filter((point) => point.invested > 0);

  const last = series.at(-1);
  if (last) cashflows.push({ date: last.date, amount: last.value });
  return { symbol, series, cashflows, installments, error: series.length ? "" : "日期區間沒有可用價格" };
}

async function buildRealDcaSeries(symbol = state.dcaSymbols[0], startDate = state.dcaStartDate, endDate = state.dcaEndDate) {
  const bundledPoints = marketDcaSeries[symbol]?.points || [];
  if (bundledPoints.length) {
    const result = buildDcaSeriesFromPricePoints(symbol, bundledPoints, startDate, endDate);
    if (result.series.length) return result;
  }

  try {
    const points = await adjustedHistoryFor(symbol, startDate, endDate);
    const result = buildDcaSeriesFromPricePoints(symbol, points, startDate, endDate);
    if (result.series.length) return result;
    return { ...result, error: result.error || "日線資料不足" };
  } catch (error) {
    return { symbol, series: [], cashflows: [], installments: [], error: "歷史日線抓取失敗" };
  }
}

async function dcaSnapshotRequests(schedule, endDate) {
  const requests = [
    ...schedule.map((date) => ({ kind: "installment", targetDate: date })),
    { kind: "valuation", targetDate: endDate },
  ].filter((request) => request.targetDate <= endDate);
  const uniqueDates = [...new Set(requests.map((request) => request.targetDate))];
  const snapshots = new Map();
  await mapWithConcurrency(uniqueDates, 4, async (date) => {
    snapshots.set(date, await historicalCloseSnapshotOnOrBefore(date));
  });
  return requests
    .map((request) => ({ ...request, snapshot: snapshots.get(request.targetDate) }))
    .filter((request) => request.snapshot?.closes);
}

function dcaValuationPoints(startDate, endDate) {
  const points = new Map();
  points.set(startDate, nearestDateIndex(startDate));
  priceDates.forEach((date, index) => {
    if (date >= startDate && date <= endDate) points.set(date, index);
  });
  points.set(endDate, nearestDateIndex(endDate));
  return [...points.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, index]) => ({ date, index }));
}

function buildDcaInstallments(startDate, endDate, frequency) {
  const dates = [];
  const [startYear, startMonth, startDay] = startDate.split("-").map(Number);
  const [endYear, endMonth, endDay] = endDate.split("-").map(Number);
  const current = new Date(Date.UTC(startYear, startMonth - 1, startDay));
  const end = new Date(Date.UTC(endYear, endMonth - 1, endDay));
  while (current <= end) {
    dates.push(current.toISOString().slice(0, 10));
    if (frequency === "weekly") {
      current.setUTCDate(current.getUTCDate() + 7);
    } else {
      current.setUTCMonth(current.getUTCMonth() + 1);
    }
  }
  return dates;
}

function annualizedReturn(totalReturn, years) {
  if (years <= 0) return 0;
  return (1 + totalReturn) ** (1 / years) - 1;
}

function xirr(cashflows) {
  const valid = cashflows.filter((flow) => Number.isFinite(flow.amount) && flow.date);
  if (!valid.some((flow) => flow.amount < 0) || !valid.some((flow) => flow.amount > 0)) return NaN;
  const startTime = new Date(`${valid[0].date}T00:00:00`).getTime();
  const yearsFromStart = (date) => (new Date(`${date}T00:00:00`).getTime() - startTime) / 31_557_600_000;
  let low = -0.9999;
  let high = 10;
  const npv = (rate) => valid.reduce((sum, flow) => sum + flow.amount / (1 + rate) ** yearsFromStart(flow.date), 0);
  let lowValue = npv(low);
  let highValue = npv(high);
  while (lowValue * highValue > 0 && high < 1000) {
    high *= 2;
    highValue = npv(high);
  }
  if (lowValue * highValue > 0) return NaN;
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const mid = (low + high) / 2;
    const midValue = npv(mid);
    if (Math.abs(midValue) < 0.01) return mid;
    if (lowValue * midValue <= 0) {
      high = mid;
      highValue = midValue;
    } else {
      low = mid;
      lowValue = midValue;
    }
  }
  return (low + high) / 2;
}

function yearsBetween(startDate, endDate) {
  const start = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);
  if (!(end > start)) return 0;
  return (end - start) / 31_557_600_000;
}

function riskPosition(symbol) {
  const points = snapshotPricePoints(symbol);
  const prices = points.map((point) => point.value);
  const latest = latestPrice(symbol);
  const highPoint = points.reduce((best, point) => (point.value > best.value ? point : best), points.at(-1));
  const lowPoint = points.reduce((best, point) => (point.value < best.value ? point : best), points.at(-1));
  const high = highPoint?.value || latest;
  const low = lowPoint?.value || latest;
  const sorted = [...prices].sort((a, b) => a - b);
  const belowCount = sorted.filter((price) => price <= latest).length;
  const percentile = sorted.length > 1 ? ((belowCount - 1) / (sorted.length - 1)) * 100 : NaN;
  const snapshotAverage = prices.reduce((sum, price) => sum + price, 0) / Math.max(1, prices.length);
  const highDistance = latest / high - 1;
  const lowDistance = latest / low - 1;
  const averageDistance = snapshotAverage > 0 ? latest / snapshotAverage - 1 : NaN;
  const highLowScore = Number.isFinite(percentile) ? Math.round(Math.max(0, Math.min(100, percentile))) : "--";
  return {
    latest,
    high,
    low,
    highDate: highPoint?.date || "",
    lowDate: lowPoint?.date || "",
    percentile,
    highDistance,
    lowDistance,
    averageDistance,
    highLowScore,
    sampleCount: points.length,
  };
}

function continuityScore(symbol) {
  const returns = ["1d", "1w", "1m", "1y"].map((period) => rankingReturn(symbol, period)).filter(Number.isFinite);
  if (returns.length === 0) return 0;
  return Math.round(returns.reduce((sum, value) => sum + Math.max(0, Math.min(1, value + 0.15)), 0) * (100 / returns.length));
}

function drawLineChart(canvas, seriesList, formatter = currency, options = {}) {
  const ctx = canvas.getContext("2d");
  const values = seriesList.flatMap((item) => item.series.map((point) => point.value));
  const dateValues = seriesList
    .flatMap((item) => item.series.map((point) => point.date))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  if (values.length === 0) {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#fbfcfb";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#667068";
    ctx.font = "13px system-ui";
    ctx.fillText("這個日期區間沒有可用資料。", 24, 44);
    return;
  }
  const max = Math.max(...values, 1) * 1.08;
  const startDate = options.startDate || dateValues[0] || "";
  const endDate = options.endDate || dateValues.at(-1) || "";
  const startTime = startDate ? new Date(`${startDate}T00:00:00Z`).getTime() : 0;
  const endTime = endDate ? new Date(`${endDate}T00:00:00Z`).getTime() : startTime;
  const rangeTime = Math.max(1, endTime - startTime);
  const padding = { top: 24, right: 26, bottom: 38, left: 70 };
  const width = canvas.width - padding.left - padding.right;
  const height = canvas.height - padding.top - padding.bottom;

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#fbfcfb";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.strokeStyle = "#dce2dc";
  ctx.lineWidth = 1;
  ctx.font = "13px system-ui";
  ctx.fillStyle = "#667068";

  for (let i = 0; i <= 4; i += 1) {
    const y = padding.top + height * (i / 4);
    ctx.beginPath();
    ctx.moveTo(padding.left, y);
    ctx.lineTo(canvas.width - padding.right, y);
    ctx.stroke();
    ctx.fillText(formatter.format(max * (1 - i / 4)), 12, y + 4);
  }

  seriesList.forEach(({ series, color }) => {
    ctx.strokeStyle = color;
    ctx.lineWidth = 3;
    ctx.beginPath();
    series.forEach((point, index) => {
      const pointTime = new Date(`${point.date}T00:00:00Z`).getTime();
      const clampedTime = Math.max(startTime, Math.min(endTime, pointTime));
      const x = padding.left + width * ((clampedTime - startTime) / rangeTime);
      const y = padding.top + height * (1 - point.value / max);
      if (index === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
  });

  ctx.fillStyle = "#667068";
  ctx.fillText(startDate, padding.left, canvas.height - 14);
  ctx.fillText(endDate, canvas.width - padding.right - 88, canvas.height - 14);
}

function drawBarChart(canvas, rows) {
  const ctx = canvas.getContext("2d");
  const padding = { top: 30, right: 118, bottom: 42, left: 260 };
  canvas.height = Math.max(260, rows.length * 46 + padding.top + padding.bottom);
  const width = canvas.width - padding.left - padding.right;
  const rowHeight = 42;
  const positiveMax = Math.max(...rows.map((row) => row.returnRate), 0.02);
  const negativeMax = Math.abs(Math.min(...rows.map((row) => row.returnRate), 0));
  const totalRange = positiveMax + negativeMax || 0.1;
  const zeroX = padding.left + width * (negativeMax / totalRange);

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#fbfcfb";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.font = "14px system-ui";

  if (rows.length === 0) {
    ctx.fillStyle = "#667068";
    ctx.fillText("請先加入要比較的股票。", 24, 44);
    return;
  }

  ctx.strokeStyle = "#dce2dc";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(zeroX, padding.top - 8);
  ctx.lineTo(zeroX, padding.top + rows.length * rowHeight);
  ctx.stroke();

  rows.forEach((row, index) => {
    const y = padding.top + index * rowHeight;
    const barWidth = Math.max(6, Math.abs(row.returnRate) / totalRange * width);
    const x = row.returnRate >= 0 ? zeroX : zeroX - barWidth;
    const label = `${row.symbol} ${stocks[row.symbol].name}`;
    ctx.fillStyle = row.returnRate >= 0 ? "#0f766e" : "#b42318";
    ctx.fillRect(x, y, barWidth, 24);
    ctx.fillStyle = "#18201b";
    ctx.textAlign = "right";
    ctx.fillText(label.length > 14 ? `${label.slice(0, 14)}...` : label, padding.left - 14, y + 17);
    const value = percent.format(row.returnRate);
    if (row.returnRate >= 0) {
      ctx.textAlign = "left";
      const outsideX = x + barWidth + 10;
      ctx.fillText(value, Math.min(outsideX, canvas.width - padding.right + 8), y + 17);
    } else {
      ctx.textAlign = "left";
      ctx.fillStyle = "#b42318";
      ctx.fillText(value, zeroX + 10, y + 17);
    }
  });
  ctx.textAlign = "left";
}

function renderWatchlist() {
  $("#watchlistChips").innerHTML = visibleSymbols()
    .map(
      (symbol) => `
        <span class="selected-chip">
          ${symbol} ${stocks[symbol].name}
          <button type="button" aria-label="刪除 ${symbol}" data-remove-watchlist="${symbol}">×</button>
        </span>
      `,
    )
    .join("");
}

function renderCompareList() {
  $("#compareChips").innerHTML = state.compareSymbols
    .map(
      (symbol) => `
        <span class="selected-chip">
          ${symbol} ${stocks[symbol].name}
          <button type="button" aria-label="刪除 ${symbol}" data-remove-compare="${symbol}">×</button>
        </span>
      `,
    )
    .join("");
}

function renderDcaList() {
  $("#dcaChips").innerHTML = state.dcaSymbols
    .map(
      (symbol) => `
        <span class="selected-chip">
          ${symbol} ${stocks[symbol].name}
          <button type="button" aria-label="刪除 ${symbol}" data-remove-dca="${symbol}">×</button>
        </span>
      `,
    )
    .join("");
}

function renderControls() {
  syncSelections();
  normalizeRankingPeriod();
  const sortedEntries = Object.entries(stocks).sort((a, b) => (b[1].marketCap || 0) - (a[1].marketCap || 0));
  const list = visibleSymbols();
  const options = optionHtml(list);
  $("#stockSearchList").innerHTML = sortedEntries
    .filter(([symbol, stock]) => isTaiwanSymbol(symbol) || stock.market === "US")
    .map(([symbol, stock]) => `<option value="${symbol} ${stock.name}"></option>`)
    .join("");
  $("#benchmarkSelect").innerHTML = options;
  $("#symbolInput").innerHTML = options;
  $("#riskSymbol").innerHTML = options;
  $("#portfolioAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.portfolioSymbols.includes(symbol)));
  $("#compareAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.compareSymbols.includes(symbol)));
  $("#dcaAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.dcaSymbols.includes(symbol)));
  $("#portfolioAddForm button").disabled = !$("#portfolioAddSelect").value;
  $("#compareAddForm button").disabled = !$("#compareAddSelect").value;
  $("#dcaAddForm button").disabled = !$("#dcaAddSelect").value;
  $("#benchmarkSelect").value = state.benchmark;
  $("#riskSymbol").value = state.riskSymbol;
  $("#symbolInput").value = list[0] || "2330";
  $("#dateInput").value = "2024-01-15";
  $("#amountInput").value = 1000;
  $("#priceInput").value = latestPrice(list[0] || "2330") || "";
  $("#portfolioAmount").value = state.portfolioAmount;
  $("#comparePeriod").value = state.comparePeriod;
  $("#compareStartDate").value = state.compareStartDate;
  $("#compareEndDate").value = state.compareEndDate;
  updateRankingPeriodOptions();
  $("#rankingPeriod").value = state.rankingPeriod;
  $("#dcaStartDate").value = state.dcaStartDate;
  $("#dcaEndDate").value = state.dcaEndDate;

  const sectors = ["全部", ...new Set(Object.values(stocks).map((stock) => stock.sector).filter(Boolean))];
  $("#sectorFilter").innerHTML = sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("");
  $("#sectorFilter").value = sectors.includes(state.sector) ? state.sector : "全部";

  renderWatchlist();
  renderCompareList();
  renderDcaList();

  const defaultWeights = { "2330": 35, "0050": 35, "2454": 20, "2881": 10, "2317": 10 };
  $("#portfolioWeights").innerHTML = state.portfolioSymbols
    .map(
      (symbol) => `
        <label class="weight-item">
          <span>${symbol} ${stocks[symbol].name}</span>
          <input data-weight="${symbol}" type="number" min="0" max="100" step="5" value="${state.portfolioWeights[symbol] ?? defaultWeights[symbol] ?? 0}" />
          <button class="delete-btn" type="button" data-remove-portfolio="${symbol}">刪除</button>
        </label>
      `,
    )
    .join("");
}

function renderMetrics() {
  const summary = summarizePortfolio();
  $("#totalInvested").textContent = currency.format(summary.totalInvested);
  $("#currentValue").textContent = currency.format(summary.currentValue);
  $("#totalReturn").textContent = percent.format(summary.totalReturn);
  $("#totalReturn").className = classForReturn(summary.totalReturn);
  $("#maxDrawdown").textContent = percent.format(summary.drawdown);
  $("#maxDrawdown").className = "return-negative";
}

async function renderScenarioRanking() {
  const token = ++scenarioRenderToken;
  if (trades.length === 0) {
    $("#scenarioRanking").innerHTML =
      `<div class="rank-row"><strong>尚未有交易</strong><small>新增一筆真實交易後，就能比較同日期同金額買其他標的的結果。</small></div>`;
    return;
  }

  $("#scenarioRanking").innerHTML = `<div class="rank-row"><strong>正在重算替代結果</strong><small>使用每筆交易日期的還原日線買入價，而不是範例價格序列。</small></div>`;
  const rows = await mapWithConcurrency(visibleSymbols(), 4, async (symbol) => ({ symbol, ...(await scenarioTotals(symbol)) }));
  if (token !== scenarioRenderToken) return;

  $("#scenarioRanking").innerHTML = rows
    .filter((item) => item.countedTrades > 0)
    .sort((a, b) => {
      if (a.missingTrades && !b.missingTrades) return 1;
      if (!a.missingTrades && b.missingTrades) return -1;
      return (Number.isFinite(b.returnRate) ? b.returnRate : -Infinity) - (Number.isFinite(a.returnRate) ? a.returnRate : -Infinity);
    })
    .map((item) => {
      const stock = stocks[item.symbol];
      const complete = item.missingTrades === 0 && Number.isFinite(item.returnRate);
      return `
        <div class="rank-row">
          <div>
            <strong>${item.symbol} ${stock.name}</strong>
            <small>${complete
              ? `${item.countedTrades} 筆交易同日期同金額買入，目前 ${currency.format(item.value)}`
              : `只有 ${item.countedTrades}/${trades.length} 筆交易日期有日線，先不列入完整比較`}</small>
          </div>
          <strong class="${classForReturn(complete ? item.returnRate : NaN)}">${complete ? percent.format(item.returnRate) : "--"}</strong>
        </div>
      `;
    })
    .join("") || `<div class="rank-row"><strong>沒有可用替代結果</strong><small>這些交易日期沒有抓到可用日線。</small><strong>--</strong></div>`;
}

function renderTrades() {
  $("#tradeSelect").innerHTML = trades
    .map((trade) => `<option value="${trade.id}">${trade.date} ${trade.symbol} ${stocks[trade.symbol].name}</option>`)
    .join("");
  if (!trades.some((trade) => trade.id === state.selectedTradeId)) state.selectedTradeId = trades[0]?.id;
  $("#tradeSelect").value = state.selectedTradeId || "";

  if (trades.length === 0) {
    $("#tradeTable").innerHTML = `<tr><td colspan="8">尚未有交易，請先新增一筆買入紀錄。</td></tr>`;
    return;
  }

  $("#tradeTable").innerHTML = trades
    .map((trade) => {
      const entryPrice = entryPriceForTrade(trade);
      const shares = sharesFor(trade);
      const invested = investedAmountFor(trade);
      const value = currentValueFor(trade);
      const returnRate = invested ? value / invested - 1 : 0;
      return `
        <tr>
          <td>${trade.date}</td>
          <td>${trade.symbol} ${stocks[trade.symbol].name}</td>
          <td>${shares.toLocaleString("zh-TW", { maximumFractionDigits: 3 })}</td>
          <td>${entryPrice.toFixed(2)}</td>
          <td>${currency.format(invested)}</td>
          <td>${currency.format(value)}</td>
          <td class="${classForReturn(returnRate)}">${percent.format(returnRate)}</td>
          <td><button class="delete-btn" type="button" data-delete="${trade.id}">刪除</button></td>
        </tr>
      `;
    })
    .join("");
}

async function renderSingleTradeAnalysis() {
  const token = ++singleTradeRenderToken;
  const trade = trades.find((item) => item.id === state.selectedTradeId);
  if (!trade) {
    $("#singleTradeAnalysis").innerHTML = "<p>請先新增一筆交易。</p>";
    return;
  }

  $("#singleTradeAnalysis").innerHTML = `<p>正在抓取 ${trade.date} 的官方收盤價，用同一天每檔股票價格重算...</p>`;

  let snapshot;
  try {
    snapshot = await historicalCloseSnapshotOnOrBefore(trade.date);
  } catch (error) {
    console.warn(error);
  }
  if (token !== singleTradeRenderToken) return;

  if (!snapshot) {
    $("#singleTradeAnalysis").innerHTML = `<p>無法取得 ${trade.date} 或前 14 天內的官方收盤價，先不產生假設買入結果，避免顯示錯誤報酬。</p>`;
    return;
  }

  const invested = investedAmountFor(trade);
  const symbols = visibleSymbols();
  const entryPoints = Object.fromEntries(
    await Promise.all(
      symbols.map(async (symbol) => [symbol, symbol === trade.symbol ? null : await entryPointForHypothetical(symbol, trade.date, snapshot)]),
    ),
  );
  if (token !== singleTradeRenderToken) return;

  $("#singleTradeAnalysis").innerHTML = symbols
    .map((symbol) => (symbol === trade.symbol ? actualTradeResult(trade) : hypotheticalTradeResult(trade, symbol, entryPoints[symbol])))
    .filter((item) => item.available)
    .sort((a, b) => b.returnRate - a.returnRate)
    .map((item) => {
      const isReal = item.symbol === trade.symbol ? "真實買入" : "假設買入";
      const basis =
        item.symbol === trade.symbol
          ? `你的買入價 ${item.entryPrice.toFixed(2)}`
          : `${item.entryDate} ${item.entrySource === "adjusted" ? "還原收盤" : "原始收盤"} ${item.entryPrice.toFixed(2)}`;
      return `
        <div class="analysis-card">
          <span>${isReal}</span>
          <strong>${item.symbol} ${stocks[item.symbol].name}</strong>
          <small>${basis}，${currency.format(invested)} 到 ${currency.format(item.value)}</small>
          <b class="${classForReturn(item.returnRate)}">${percent.format(item.returnRate)}</b>
        </div>
      `;
    })
    .join("");
}

function renderPortfolioEngine() {
  const amount = Number($("#portfolioAmount").value) || 0;
  const weightSum = portfolioWeights().reduce((sum, item) => sum + item.weight, 0);
  const series = buildWeightedPortfolioSeries(amount);
  const finalValue = series.at(-1).value;
  const totalReturn = amount ? finalValue / amount - 1 : 0;
  const normalized = Math.abs(weightSum - 1) < 0.001 ? "權重合計 100%" : `權重合計 ${Math.round(weightSum * 100)}%`;

  $("#portfolioResults").innerHTML = `
    <div class="analysis-card"><span>${normalized}</span><strong>${currency.format(finalValue)}</strong><small>目前組合市值</small></div>
    <div class="analysis-card"><span>報酬率</span><strong class="${classForReturn(totalReturn)}">${percent.format(totalReturn)}</strong><small>組合報酬 = 加權個股報酬</small></div>
    <div class="analysis-card"><span>最大跌幅</span><strong class="return-negative">${percent.format(maxDrawdown(series))}</strong><small>歷史資產高點到低點</small></div>
    <div class="analysis-card"><span>波動度</span><strong>${percent.format(volatility(series))}</strong><small>時間序列標準差</small></div>
    <div class="analysis-card"><span>勝率</span><strong>${percent.format(winRate(series))}</strong><small>有多少期間為正報酬</small></div>
    <div class="analysis-card"><span>Sharpe Ratio</span><strong>${sharpeRatio(series).toFixed(2)}</strong><small>報酬 / 風險</small></div>
  `;
}

async function renderCompareEngine() {
  const token = ++compareRenderToken;
  const customFields = $("#compareCustomFields");
  customFields.hidden = state.comparePeriod !== "custom";
  if (state.comparePeriod === "custom" && state.compareStartDate > state.compareEndDate) {
    $("#compareBasis").textContent = "日期區間錯誤：開始日期必須早於結束日期。";
    drawBarChart($("#compareBarChart"), []);
    $("#compareBreakdown").innerHTML = "";
    return;
  }

  const { startTarget, endTarget, label: periodLabel } = comparePeriodDates();
  $("#compareBasis").textContent = `正在取得 ${startTarget} 到 ${endTarget} 的官方收盤價...`;
  drawBarChart($("#compareBarChart"), []);
  $("#compareBreakdown").innerHTML = "";

  const [startSnapshot, endSnapshot] = await Promise.all([compareSnapshotFor(startTarget), compareSnapshotFor(endTarget)]);
  if (token !== compareRenderToken) return;

  if (!startSnapshot || !endSnapshot) {
    $("#compareBasis").textContent = `無法取得 ${startTarget} 到 ${endTarget} 的官方收盤價，先不產生比較結果，避免顯示錯誤資料。`;
    return;
  }

  const rows = state.compareSymbols
    .map((symbol) => {
      const startPrice = compareClose(symbol, startSnapshot);
      const endPrice = compareClose(symbol, endSnapshot);
      return {
        symbol,
        startDate: startSnapshot.date,
        endDate: endSnapshot.date,
        startPrice,
        endPrice,
        returnRate: startPrice > 0 && endPrice > 0 ? endPrice / startPrice - 1 : null,
      };
    })
    .filter((row) => Number.isFinite(row.returnRate))
    .sort((a, b) => b.returnRate - a.returnRate);
  drawBarChart($("#compareBarChart"), rows);

  const first = rows[0];
  const startDate = first?.startDate || "--";
  const endDate = first?.endDate || "--";
  $("#compareBasis").textContent =
    `比較基準：${periodLabel}，用「報酬率 = (結束價 - 起始價) / 起始價」排序。` +
    `目前區間 ${startDate} 到 ${endDate}；1月就是最新資料日往前 30 天，若該日休市則往前取最近交易日。`;

  $("#compareBreakdown").innerHTML = rows.length
    ? rows
      .map(
        (row) => `
          <div>
            <strong>${row.symbol} ${stocks[row.symbol].name}</strong>
            <span>${row.startDate} ${row.startPrice.toFixed(2)} → ${row.endDate} ${row.endPrice.toFixed(2)}</span>
            <b class="${classForReturn(row.returnRate)}">${percent.format(row.returnRate)}</b>
          </div>
        `,
      )
      .join("")
    : `<div><strong>尚未選擇股票</strong><span>先加入要比較的標的。</span><b>--</b></div>`;
}

async function renderDcaEngine() {
  const token = ++dcaRenderToken;
  if (state.dcaStartDate > state.dcaEndDate) {
    $("#dcaBasis").textContent = "日期區間錯誤：開始日期必須早於結束日期。";
    $("#dcaInvested").textContent = "--";
    $("#dcaValue").textContent = "--";
    $("#dcaReturn").textContent = "--";
    $("#dcaMdd").textContent = "--";
    drawLineChart($("#dcaChart"), []);
    $("#dcaLegend").innerHTML = "";
    $("#dcaResults").innerHTML = `<div class="rank-row"><strong>日期區間錯誤</strong><small>開始日期必須早於結束日期。</small></div>`;
    return;
  }

  const schedule = buildDcaInstallments(state.dcaStartDate, state.dcaEndDate, state.dcaFrequency);
  const scheduledCount = schedule.length;
  $("#dcaBasis").textContent =
    `正在取得 ${state.dcaStartDate} 到 ${state.dcaEndDate} 的還原日線，依每期投入金額計算股數與 XIRR 年化...`;
  $("#dcaInvested").textContent = "--";
  $("#dcaValue").textContent = "--";
  $("#dcaReturn").textContent = "--";
  $("#dcaMdd").textContent = "--";
  drawLineChart($("#dcaChart"), []);
  $("#dcaLegend").innerHTML = "";
  $("#dcaResults").innerHTML = "";

  let seriesBySymbol = await mapWithConcurrency(state.dcaSymbols, 4, (symbol) =>
    buildRealDcaSeries(symbol, state.dcaStartDate, state.dcaEndDate),
  );
  if (token !== dcaRenderToken) return;

  const missingSymbols = seriesBySymbol.filter((item) => !item.series.length && isTaiwanSymbol(item.symbol)).map((item) => item.symbol);
  if (missingSymbols.length) {
    $("#dcaBasis").textContent =
      `部分標的日線不足，正在一次補抓 ${state.dcaStartDate} 到 ${state.dcaEndDate} 的投入日官方快照...`;
    const snapshotRequests = await dcaSnapshotRequests(schedule, state.dcaEndDate);
    if (token !== dcaRenderToken) return;
    seriesBySymbol = seriesBySymbol.map((item) =>
      item.series.length ? item : buildDcaSeriesFromSnapshots(item.symbol, snapshotRequests, state.dcaEndDate),
    );
  }

  $("#dcaBasis").textContent =
    `模擬區間：${state.dcaStartDate} 到 ${state.dcaEndDate}。` +
    `規則：${state.dcaFrequency === "weekly" ? "每週" : "每月"}投入 ${currency.format(state.dcaAmount)}，共 ${scheduledCount} 期。` +
    "排定日若非交易日，使用往前最近交易日還原收盤價買入；總報酬 = 總資產 / 累積投入 - 1，年化 = 每期現金流 XIRR。";
  const rows = seriesBySymbol.map(({ symbol, series }) => {
    const last = series.at(-1) || { invested: 0, value: 0 };
    const rate = last.invested ? last.value / last.invested - 1 : 0;
    const result = seriesBySymbol.find((item) => item.symbol === symbol);
    return { symbol, series, last, rate, annualized: xirr(result?.cashflows || []), drawdown: maxDrawdown(series), installments: result?.installments || [], error: result?.error || "" };
  });
  const validRows = rows.filter((row) => row.series.length && row.last.invested > 0);
  if (validRows.length === 0) {
    $("#dcaBasis").textContent = "目前選取標的在這個日期區間都沒有抓到可用官方收盤資料，沒有產生計算結果。";
    $("#dcaInvested").textContent = "--";
    $("#dcaValue").textContent = "--";
    $("#dcaReturn").textContent = "--";
    $("#dcaMdd").textContent = "--";
    drawLineChart($("#dcaChart"), []);
    $("#dcaLegend").innerHTML = "";
    $("#dcaResults").innerHTML = rows
      .map((row) => `<div class="rank-row"><strong>${row.symbol} ${stocks[row.symbol].name}</strong><small>${row.error || "沒有可用資料"}</small><strong>--</strong></div>`)
      .join("");
    return;
  }
  const primary = validRows[0]?.series || [];
  const best = [...validRows].sort((a, b) => (Number.isFinite(b.annualized) ? b.annualized : b.rate) - (Number.isFinite(a.annualized) ? a.annualized : a.rate))[0] || { symbol: "", last: { invested: 0, value: 0 }, rate: 0, annualized: NaN, drawdown: 0 };
  const worstDrawdown = validRows.reduce((worst, row) => Math.min(worst, row.drawdown), 0);
  $("#dcaInvested").textContent = currency.format(best.last.invested);
  $("#dcaValue").textContent = best.symbol ? `${best.symbol} ${currency.format(best.last.value)}` : "--";
  $("#dcaReturn").textContent = `${percent.format(best.rate)} / ${metricValue(best.annualized)}`;
  $("#dcaReturn").className = classForReturn(best.rate);
  $("#dcaMdd").className = "return-negative";
  $("#dcaMdd").textContent = percent.format(worstDrawdown);
  const colors = ["#0f766e", "#a15c07", "#2563eb", "#7c3aed", "#b42318", "#475569", "#15803d"];
  drawLineChart($("#dcaChart"), [
    ...(primary.length ? [{ series: primary.map((point) => ({ date: point.date, value: point.invested })), color: "#667068" }] : []),
    ...validRows.map((item, index) => ({ series: item.series, color: colors[index % colors.length] })),
  ], currency, { startDate: state.dcaStartDate, endDate: state.dcaEndDate });
  $("#dcaLegend").innerHTML = [
    ...(primary.length ? [{ label: "累積投入", color: "#667068" }] : []),
    ...validRows.map((item, index) => ({
      label: `${item.symbol} ${stocks[item.symbol].name}`,
      color: colors[index % colors.length],
    })),
  ]
    .map((item) => `<span><i style="background:${item.color}"></i>${item.label}</span>`)
    .join("");
  if (rows.length === 0) {
    $("#dcaResults").innerHTML = `<div class="rank-row"><strong>尚未選擇標的</strong><small>先從你的公司股票清單加入一檔股票或 ETF。</small></div>`;
    return;
  }
  $("#dcaResults").innerHTML = rows
    .sort((a, b) => {
      if (!a.last.invested && b.last.invested) return 1;
      if (a.last.invested && !b.last.invested) return -1;
      return (Number.isFinite(b.annualized) ? b.annualized : b.rate) - (Number.isFinite(a.annualized) ? a.annualized : a.rate);
    })
    .map(({ symbol, last, rate, annualized, installments, error }) => {
      if (!last.invested) {
        return `
          <div class="rank-row">
            <div>
              <strong>${symbol} ${stocks[symbol].name}</strong>
              <small>${error || "沒有可用歷史日線"}</small>
            </div>
            <strong>--</strong>
          </div>
        `;
      }
      return `
        <div class="rank-row dca-result-row">
          <div>
            <strong>${symbol} ${stocks[symbol].name}</strong>
            <small>${installments.length} 期 · 累積投入 ${currency.format(last.invested)} · 總資產 ${currency.format(last.value)}</small>
          </div>
          <div class="dca-result-metrics">
            <span class="${classForReturn(rate)}">${percent.format(rate)}</span>
            <small>總報酬</small>
            <span class="${classForReturn(annualized)}">${metricValue(annualized)}</span>
            <small>年化</small>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderMarketRanking() {
  normalizeRankingPeriod();
  updateRankingPeriodOptions();
  $("#rankingPeriod").value = state.rankingPeriod;
  $("#rankingReturnHeader").textContent = `${rankingPeriodLabel(state.rankingPeriod)}漲幅`;
  const sourceSymbols = Object.keys(stocks).filter(isTaiwanRankingSymbol);
  const sectors = ["全部", ...new Set(sourceSymbols.map((symbol) => stocks[symbol].sector).filter(Boolean))];
  $("#sectorFilter").innerHTML = sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("");
  if (!sectors.includes(state.sector)) state.sector = "全部";
  $("#sectorFilter").value = state.sector;
  const history = marketHistoryPeriods[state.rankingPeriod];
  $("#rankingNote").textContent =
    state.rankingPeriod === "today" || state.rankingPeriod === "1d"
      ? "今日 / 1日排行使用證交所與櫃買中心每日真實行情漲跌幅。"
      : `${rankingPeriodLabel(state.rankingPeriod)}排行使用 ${history?.date || "歷史基準日"} 到 ${state.marketDataDate} 的官方原始收盤價計算；遇假日會往前取最近交易日，尚未還原除權息。`;

  const rows = sourceSymbols
    .filter((symbol) => state.sector === "全部" || stocks[symbol].sector === state.sector)
    .map((symbol) => {
      const returnRate = rankingReturn(symbol, state.rankingPeriod);
      return {
        symbol,
        returnRate,
        score: continuityScore(symbol),
        marketCap: stocks[symbol].marketCap || 0,
      };
    })
    .filter((row) => Number.isFinite(row.returnRate))
    .sort((a, b) => b.returnRate - a.returnRate)
    .slice(0, 10);

  if (rows.length === 0) {
    $("#marketRanking").innerHTML = `<tr><td colspan="6">此條件沒有資料，請改成「全部」類股或先把台股加入我的公司股票清單。</td></tr>`;
    return;
  }

  $("#marketRanking").innerHTML = rows
    .map((row, index) => {
      const stock = stocks[row.symbol];
      return `
        <tr>
          <td>${index + 1}</td>
          <td>${row.symbol} ${stock.name}<br><small>${stock.sector}</small></td>
          <td class="${classForReturn(row.returnRate)}">${percent.format(row.returnRate)}</td>
          <td>${formatVolume(stock.volume)}</td>
          <td>${formatMarketCap(row.marketCap)}</td>
          <td><span class="score-pill">${row.score}</span></td>
        </tr>
      `;
    })
    .join("");
}

function renderRiskPosition() {
  const symbol = state.riskSymbol;
  const stock = stocks[symbol];
  const risk = riskPosition(symbol);
  $("#riskCards").innerHTML = `
    <div class="analysis-card score-card"><span>高低位分數</span><strong>${risk.highLowScore}</strong><small>依 ${risk.sampleCount} 個官方收盤快照估算</small></div>
    <div class="analysis-card"><span>快照價格百分位</span><strong>${Number.isFinite(risk.percentile) ? risk.percentile.toFixed(0) : "--"}</strong><small>${symbol} ${stock.name} 最新價 ${risk.latest}</small></div>
    <div class="analysis-card"><span>可用快照高點</span><strong class="${classForReturn(risk.highDistance)}">${metricValue(risk.highDistance)}</strong><small>${risk.highDate || "--"} 高點 ${risk.high}</small></div>
    <div class="analysis-card"><span>可用快照低點</span><strong class="${classForReturn(risk.lowDistance)}">${metricValue(risk.lowDistance)}</strong><small>${risk.lowDate || "--"} 低點 ${risk.low}</small></div>
    <div class="analysis-card"><span>與快照均價距離</span><strong class="${classForReturn(risk.averageDistance)}">${metricValue(risk.averageDistance)}</strong><small>200 日均線需每日 price_history</small></div>
  `;
}

function renderAdvancedMetrics() {
  const realSeries = portfolioSnapshotSeries("real");
  const benchmarkSeries = portfolioSnapshotSeries(state.benchmark);
  const enoughData = hasEnoughSeries(realSeries);
  const benchmarkEnough = hasEnoughSeries(benchmarkSeries);
  const drawdown = enoughData ? maxDrawdown(realSeries) : NaN;
  const wins = enoughData ? winRate(realSeries) : NaN;
  const vol = enoughData ? volatility(realSeries) : NaN;
  const sharpe = enoughData ? sharpeRatio(realSeries) : NaN;
  const corr = enoughData && benchmarkEnough ? correlation(realSeries, benchmarkSeries) : NaN;
  const note = enoughData ? `${realSeries.length} 個官方歷史快照估算` : "需至少 3 個歷史快照";
  $("#advancedMetrics").innerHTML = `
    <div class="analysis-card"><span>最大跌幅 MDD</span><strong class="return-negative">${metricValue(drawdown)}</strong><small>${note}</small></div>
    <div class="analysis-card"><span>勝率</span><strong>${metricValue(wins)}</strong><small>${note}</small></div>
    <div class="analysis-card"><span>波動率</span><strong>${metricValue(vol)}</strong><small>快照報酬標準差，非逐日年化</small></div>
    <div class="analysis-card"><span>Sharpe Ratio</span><strong>${Number.isFinite(sharpe) ? sharpe.toFixed(2) : "--"}</strong><small>需每日資料才適合正式使用</small></div>
    <div class="analysis-card"><span>相關性分析</span><strong>${Number.isFinite(corr) ? corr.toFixed(2) : "--"}</strong><small>與 ${state.benchmark} 的快照同步程度</small></div>
    <div class="analysis-card"><span>再平衡模擬</span><strong>年度</strong><small>下一階段會加入每年權重重置</small></div>
  `;
}

function correlation(a, b) {
  const ar = returnsFromValues(a);
  const br = returnsFromValues(b);
  const n = Math.min(ar.length, br.length);
  if (!n) return 0;
  const ax = ar.slice(0, n);
  const bx = br.slice(0, n);
  const ma = ax.reduce((sum, value) => sum + value, 0) / n;
  const mb = bx.reduce((sum, value) => sum + value, 0) / n;
  const numerator = ax.reduce((sum, value, index) => sum + (value - ma) * (bx[index] - mb), 0);
  const da = Math.sqrt(ax.reduce((sum, value) => sum + (value - ma) ** 2, 0));
  const db = Math.sqrt(bx.reduce((sum, value) => sum + (value - mb) ** 2, 0));
  return da && db ? numerator / (da * db) : 0;
}

function renderCharts() {
  drawLineChart($("#equityChart"), [
    { series: buildScenarioEquitySeries("real"), color: "#0f766e" },
    { series: buildScenarioEquitySeries(state.benchmark), color: "#a15c07" },
  ]);
}

function render() {
  renderMetrics();
  renderScenarioRanking();
  renderTrades();
  renderSingleTradeAnalysis();
  renderPortfolioEngine();
  renderCompareEngine();
  renderDcaEngine();
  renderMarketRanking();
  renderRiskPosition();
  renderAdvancedMetrics();
  renderCharts();
}

function bindEvents() {
  $("#benchmarkSelect").addEventListener("change", (event) => {
    state.benchmark = event.target.value;
    saveUserState();
    render();
  });

  $("#tradeSelect").addEventListener("change", (event) => {
    state.selectedTradeId = event.target.value;
    saveUserState();
    renderSingleTradeAnalysis();
  });

  $("#comparePeriod").addEventListener("change", (event) => {
    state.comparePeriod = event.target.value;
    saveUserState();
    renderCompareEngine();
  });
  $("#compareStartDate").addEventListener("change", (event) => {
    state.compareStartDate = event.target.value || priceDates[0];
    saveUserState();
    renderCompareEngine();
  });
  $("#compareEndDate").addEventListener("change", (event) => {
    state.compareEndDate = event.target.value || priceDates.at(-1);
    saveUserState();
    renderCompareEngine();
  });

  $("#watchlistForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const symbol = resolveStockInput($("#stockSearchInput").value);
    if (!symbol) return;
    state.watchlist = uniqueSymbols([...state.watchlist, symbol]);
    $("#stockSearchInput").value = "";
    saveUserState();
    renderControls();
    render();
  });

  $("#watchlistChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-watchlist]");
    if (!button) return;
    if (state.watchlist.length <= 1) return;
    const symbol = button.dataset.removeWatchlist;
    state.watchlist = state.watchlist.filter((item) => item !== symbol);
    syncSelections();
    saveUserState();
    renderControls();
    render();
  });

  $("#portfolioAddForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const symbol = $("#portfolioAddSelect").value;
    if (!symbol) return;
    state.portfolioSymbols = uniqueSymbols([...state.portfolioSymbols, symbol]);
    saveUserState();
    renderControls();
    renderPortfolioEngine();
  });

  $("#compareAddForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const symbol = $("#compareAddSelect").value;
    if (!symbol) return;
    state.compareSymbols = uniqueSymbols([...state.compareSymbols, symbol]);
    saveUserState();
    renderControls();
    renderCompareEngine();
  });

  $("#compareChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-compare]");
    if (!button) return;
    state.compareSymbols = state.compareSymbols.filter((symbol) => symbol !== button.dataset.removeCompare);
    saveUserState();
    renderControls();
    renderCompareEngine();
  });

  $("#portfolioAmount").addEventListener("input", () => {
    state.portfolioAmount = Number($("#portfolioAmount").value) || 0;
    saveUserState();
    renderPortfolioEngine();
  });
  $("#portfolioWeights").addEventListener("input", () => {
    syncPortfolioWeightsFromInputs();
    saveUserState();
    renderPortfolioEngine();
  });
  $("#portfolioWeights").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-portfolio]");
    if (!button) return;
    state.portfolioSymbols = state.portfolioSymbols.filter((symbol) => symbol !== button.dataset.removePortfolio);
    delete state.portfolioWeights[button.dataset.removePortfolio];
    saveUserState();
    renderControls();
    renderPortfolioEngine();
  });

  $("#dcaFrequency").addEventListener("change", (event) => {
    state.dcaFrequency = event.target.value;
    saveUserState();
    renderDcaEngine();
  });
  $("#dcaAmount").addEventListener("input", (event) => {
    state.dcaAmount = Number(event.target.value) || 0;
    saveUserState();
    renderDcaEngine();
  });
  $("#dcaStartDate").addEventListener("change", (event) => {
    state.dcaStartDate = event.target.value || priceDates[0];
    saveUserState();
    renderDcaEngine();
  });
  $("#dcaEndDate").addEventListener("change", (event) => {
    state.dcaEndDate = event.target.value || priceDates.at(-1);
    saveUserState();
    renderDcaEngine();
  });
  $("#dcaAddForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const symbol = $("#dcaAddSelect").value;
    if (!symbol) return;
    state.dcaSymbols = uniqueSymbols([...state.dcaSymbols, symbol]);
    saveUserState();
    renderControls();
    renderDcaEngine();
  });
  $("#dcaChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-dca]");
    if (!button) return;
    state.dcaSymbols = state.dcaSymbols.filter((symbol) => symbol !== button.dataset.removeDca);
    saveUserState();
    renderControls();
    renderDcaEngine();
  });

  $("#rankingPeriod").addEventListener("change", (event) => {
    state.rankingPeriod = isRankingPeriodAvailable(event.target.value) ? event.target.value : "today";
    event.target.value = state.rankingPeriod;
    saveUserState();
    renderMarketRanking();
  });
  $("#sectorFilter").addEventListener("change", (event) => {
    state.sector = event.target.value;
    saveUserState();
    renderMarketRanking();
  });
  $("#riskSymbol").addEventListener("change", (event) => {
    state.riskSymbol = event.target.value;
    saveUserState();
    renderRiskPosition();
  });

  $("#tradeForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const symbol = $("#symbolInput").value;
    const date = $("#dateInput").value;
    const shares = Number($("#amountInput").value);
    const price = Number($("#priceInput").value);
    if (!symbol || !date || !shares || !price) return;

    const trade = { id: makeId(), symbol, date, shares, price, amount: shares * price };
    trades = [trade, ...trades];
    state.selectedTradeId = trade.id;
    event.target.reset();
    $("#symbolInput").value = symbol;
    $("#dateInput").value = date;
    $("#amountInput").value = shares;
    $("#priceInput").value = price;
    saveUserState();
    render();
  });

  $("#tradeTable").addEventListener("click", (event) => {
    const button = event.target.closest("[data-delete]");
    if (!button) return;
    trades = trades.filter((trade) => trade.id !== button.dataset.delete);
    saveUserState();
    render();
  });
}

renderControls();
bindEvents();
render();
loadDailyMarketData();

setInterval(() => {
  const cache = readCache();
  const today = new Date().toISOString().slice(0, 10);
  if (!cache || cache.cachedAt !== today) {
    loadDailyMarketData();
  }
}, 60 * 60 * 1000);
