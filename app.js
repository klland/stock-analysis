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
const CACHE_KEY = "decision-ledger-twse-cache-v1";
const USER_STATE_KEY = "decision-ledger-user-state-v1";
const usEtfSymbols = new Set(["SPY", "QQQ", "VOO", "VTI", "IVV", "SCHD", "VGT", "XLK", "SMH", "SOXX", "DIA", "IWM", "TLT", "BND", "AGG", "IBIT"]);
let userStateLoaded = false;

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
    ["benchmark", "riskSymbol", "comparePeriod", "rankingPeriod", "sector", "dcaFrequency", "dcaStartDate", "dcaEndDate"].forEach((key) => {
      if (saved[key]) state[key] = saved[key];
    });
    if (Number.isFinite(Number(saved.dcaAmount))) state.dcaAmount = Number(saved.dcaAmount);
    if (Number.isFinite(Number(saved.portfolioAmount))) state.portfolioAmount = Number(saved.portfolioAmount);
    if (saved.portfolioWeights && typeof saved.portfolioWeights === "object") state.portfolioWeights = saved.portfolioWeights;
    if (saved.selectedTradeId) state.selectedTradeId = saved.selectedTradeId;
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
    const generatedPrices = stocks[symbol]?.prices?.length
      ? [...stocks[symbol].prices.slice(0, -1), close]
      : [previous * 0.76, previous * 0.82, previous * 0.88, previous * 0.92, previous * 0.97, previous, close, close * 0.99, close * 1.01, previous, close];

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
      tradeValue: toNumber(row.TradeValue),
      prices: generatedPrices.slice(-priceDates.length),
    };
    added += 1;
  });

  state.marketDataSource = "twse";
  state.marketDataDate = payload.date || payload.fetchedAt.slice(0, 10);
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
      tradeValue: toNumber(row["成交金額"]),
      prices: stocks[symbol]?.prices?.length
        ? [...stocks[symbol].prices.slice(0, -1), close]
        : [previous * 0.78, previous * 0.83, previous * 0.88, previous * 0.93, previous * 0.97, previous, close, close * 0.99, close * 1.01, previous, close],
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
      prices: stocks[symbol]?.prices?.length
        ? [...stocks[symbol].prices.slice(0, -1), close]
        : [close * 0.72, close * 0.78, close * 0.84, close * 0.9, close * 0.96, previous, close * 0.98, close * 1.02, close * 0.99, previous, close],
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

function priceAtIndex(symbol, index) {
  return stocks[symbol].prices[Math.max(0, Math.min(index, priceDates.length - 1))];
}

function latestPrice(symbol) {
  return stocks[symbol].prices.at(-1);
}

function periodReturn(symbol, period = "1y") {
  if ((period === "today" || period === "1d") && stocks[symbol]?.live) {
    return stocks[symbol].dailyReturn || 0;
  }
  const end = priceDates.length - 1;
  const start = Math.max(0, end - (periodSteps[period] ?? 1));
  const startPrice = priceAtIndex(symbol, start);
  return latestPrice(symbol) / startPrice - 1;
}

function allPeriodReturns(symbol) {
  return ["1d", "1w", "1m", "1y"].map((period) => periodReturn(symbol, period));
}

function sharesFor(trade, symbol = trade.symbol) {
  const entryPrice = trade.price || priceOnOrAfter(symbol, trade.date);
  return trade.amount / entryPrice;
}

function currentValueFor(trade, symbol = trade.symbol) {
  return sharesFor(trade, symbol) * latestPrice(symbol);
}

function returnFor(trade, symbol = trade.symbol) {
  return currentValueFor(trade, symbol) / trade.amount - 1;
}

function classForReturn(value) {
  return value >= 0 ? "return-positive" : "return-negative";
}

function buildEquitySeries(symbolMode) {
  return priceDates.map((date, index) => {
    const value = trades.reduce((sum, trade) => {
      if (trade.date > date) return sum;
      const symbol = symbolMode === "real" ? trade.symbol : symbolMode;
      const entryPrice = trade.price || priceOnOrAfter(symbol, trade.date);
      const currentPrice = priceAtIndex(symbol, index);
      return sum + (trade.amount / entryPrice) * currentPrice;
    }, 0);
    return { date, value };
  });
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

function summarizePortfolio() {
  const totalInvested = trades.reduce((sum, trade) => sum + trade.amount, 0);
  const currentValue = trades.reduce((sum, trade) => sum + currentValueFor(trade), 0);
  const series = buildEquitySeries("real");
  return {
    totalInvested,
    currentValue,
    totalReturn: totalInvested ? currentValue / totalInvested - 1 : 0,
    drawdown: maxDrawdown(series),
  };
}

function scenarioTotals(symbol) {
  const invested = trades.reduce((sum, trade) => sum + trade.amount, 0);
  const value = trades.reduce((sum, trade) => sum + currentValueFor(trade, symbol), 0);
  return { invested, value, returnRate: invested ? value / invested - 1 : 0 };
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
  let previousDate = "";
  return priceDates
    .map((date, index) => ({ date, index }))
    .filter((point) => point.date >= startDate && point.date <= endDate)
    .map(({ date, index }) => {
    const installments = previousDate ? countInstallments(previousDate, date, state.dcaFrequency) : 1;
    if (installments > 0) {
      const amount = state.dcaAmount * installments;
      invested += amount;
      shares += amount / priceAtIndex(symbol, index);
    }
    previousDate = date;
    return {
      date,
      invested,
      value: shares * priceAtIndex(symbol, index),
    };
  });
}

function countInstallments(previousDate, currentDate, frequency) {
  const previous = new Date(`${previousDate}T00:00:00`);
  const current = new Date(`${currentDate}T00:00:00`);
  if (!(current > previous)) return 0;
  const days = Math.max(1, Math.round((current - previous) / 86_400_000));
  if (frequency === "weekly") return Math.max(1, Math.floor(days / 7));
  return Math.max(
    1,
    (current.getFullYear() - previous.getFullYear()) * 12 + current.getMonth() - previous.getMonth(),
  );
}

function annualizedReturn(totalReturn, years) {
  if (years <= 0) return 0;
  return (1 + totalReturn) ** (1 / years) - 1;
}

function yearsBetween(startDate, endDate) {
  const start = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);
  if (!(end > start)) return 0;
  return (end - start) / 31_557_600_000;
}

function riskPosition(symbol) {
  const prices = stocks[symbol].prices;
  const latest = prices.at(-1);
  const high = Math.max(...prices);
  const low = Math.min(...prices);
  const sorted = [...prices].sort((a, b) => a - b);
  const belowCount = sorted.filter((price) => price <= latest).length;
  const percentile = ((belowCount - 1) / (sorted.length - 1)) * 100;
  const ma200 = prices.slice(-8).reduce((sum, price) => sum + price, 0) / Math.min(8, prices.length);
  const highDistance = latest / high - 1;
  const lowDistance = latest / low - 1;
  const maDistance = latest / ma200 - 1;
  const highLowScore = Math.round(Math.max(0, Math.min(100, percentile * 0.65 + (maDistance + 0.2) * 100 * 0.35)));
  return { latest, high, low, percentile, highDistance, lowDistance, maDistance, highLowScore };
}

function continuityScore(symbol) {
  const returns = allPeriodReturns(symbol);
  return Math.round(
    returns.reduce((sum, value) => sum + Math.max(0, Math.min(1, value + 0.15)), 0) * 25,
  );
}

function drawLineChart(canvas, seriesList, formatter = currency) {
  const ctx = canvas.getContext("2d");
  const values = seriesList.flatMap((item) => item.series.map((point) => point.value));
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
      const x = padding.left + width * (index / Math.max(1, series.length - 1));
      const y = padding.top + height * (1 - point.value / max);
      if (index === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
  });

  ctx.fillStyle = "#667068";
  ctx.fillText(priceDates[0], padding.left, canvas.height - 14);
  ctx.fillText(priceDates.at(-1), canvas.width - padding.right - 88, canvas.height - 14);
}

function drawBarChart(canvas, rows) {
  const ctx = canvas.getContext("2d");
  const padding = { top: 30, right: 112, bottom: 42, left: 210 };
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
    const valueWidth = ctx.measureText(value).width;
    if (row.returnRate >= 0) {
      ctx.textAlign = "left";
      const outsideX = x + barWidth + 10;
      ctx.fillText(value, Math.min(outsideX, canvas.width - padding.right + 8), y + 17);
    } else {
      ctx.textAlign = "right";
      const outsideX = x - 10;
      ctx.fillText(value, Math.max(outsideX, padding.left - valueWidth - 8), y + 17);
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
  $("#amountInput").value = 50000;
  $("#portfolioAmount").value = state.portfolioAmount;
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

function renderScenarioRanking() {
  if (trades.length === 0) {
    $("#scenarioRanking").innerHTML =
      `<div class="rank-row"><strong>尚未有交易</strong><small>新增一筆真實交易後，就能比較同日期同金額買其他標的的結果。</small></div>`;
    return;
  }

  $("#scenarioRanking").innerHTML = visibleSymbols()
    .map((symbol) => ({ symbol, ...scenarioTotals(symbol) }))
    .sort((a, b) => b.returnRate - a.returnRate)
    .map((item) => {
      const stock = stocks[item.symbol];
      return `
        <div class="rank-row">
          <div>
            <strong>${item.symbol} ${stock.name}</strong>
            <small>同日期同金額買入，目前 ${currency.format(item.value)}</small>
          </div>
          <strong class="${classForReturn(item.returnRate)}">${percent.format(item.returnRate)}</strong>
        </div>
      `;
    })
    .join("");
}

function renderTrades() {
  $("#tradeSelect").innerHTML = trades
    .map((trade) => `<option value="${trade.id}">${trade.date} ${trade.symbol} ${stocks[trade.symbol].name}</option>`)
    .join("");
  if (!trades.some((trade) => trade.id === state.selectedTradeId)) state.selectedTradeId = trades[0]?.id;
  $("#tradeSelect").value = state.selectedTradeId || "";

  if (trades.length === 0) {
    $("#tradeTable").innerHTML = `<tr><td colspan="7">尚未有交易，請先新增一筆買入紀錄。</td></tr>`;
    return;
  }

  $("#tradeTable").innerHTML = trades
    .map((trade) => {
      const entryPrice = trade.price || priceOnOrAfter(trade.symbol, trade.date);
      const value = currentValueFor(trade);
      const returnRate = value / trade.amount - 1;
      return `
        <tr>
          <td>${trade.date}</td>
          <td>${trade.symbol} ${stocks[trade.symbol].name}</td>
          <td>${currency.format(trade.amount)}</td>
          <td>${entryPrice.toFixed(2)}</td>
          <td>${currency.format(value)}</td>
          <td class="${classForReturn(returnRate)}">${percent.format(returnRate)}</td>
          <td><button class="delete-btn" type="button" data-delete="${trade.id}">刪除</button></td>
        </tr>
      `;
    })
    .join("");
}

function renderSingleTradeAnalysis() {
  const trade = trades.find((item) => item.id === state.selectedTradeId);
  if (!trade) {
    $("#singleTradeAnalysis").innerHTML = "<p>請先新增一筆交易。</p>";
    return;
  }

  $("#singleTradeAnalysis").innerHTML = visibleSymbols()
    .map((symbol) => ({ symbol, value: currentValueFor(trade, symbol), returnRate: returnFor(trade, symbol) }))
    .sort((a, b) => b.returnRate - a.returnRate)
    .map((item) => {
      const isReal = item.symbol === trade.symbol ? "真實買入" : "假設買入";
      return `
        <div class="analysis-card">
          <span>${isReal}</span>
          <strong>${item.symbol} ${stocks[item.symbol].name}</strong>
          <small>${currency.format(trade.amount)} 到 ${currency.format(item.value)}</small>
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

function renderCompareEngine() {
  const rows = state.compareSymbols
    .map((symbol) => ({ symbol, returnRate: periodReturn(symbol, state.comparePeriod) }))
    .sort((a, b) => b.returnRate - a.returnRate);
  drawBarChart($("#compareBarChart"), rows);
}

function renderDcaEngine() {
  if (state.dcaStartDate > state.dcaEndDate) {
    $("#dcaInvested").textContent = "--";
    $("#dcaValue").textContent = "--";
    $("#dcaReturn").textContent = "--";
    $("#dcaMdd").textContent = "--";
    drawLineChart($("#dcaChart"), []);
    $("#dcaLegend").innerHTML = "";
    $("#dcaResults").innerHTML = `<div class="rank-row"><strong>日期區間錯誤</strong><small>開始日期必須早於結束日期。</small></div>`;
    return;
  }

  const seriesBySymbol = state.dcaSymbols.map((symbol) => ({ symbol, series: buildDcaSeries(symbol) }));
  const primary = seriesBySymbol[0]?.series || [];
  const rows = seriesBySymbol.map(({ symbol, series }) => {
    const last = series.at(-1) || { invested: 0, value: 0 };
    const rate = last.invested ? last.value / last.invested - 1 : 0;
    return { symbol, series, last, rate, drawdown: maxDrawdown(series) };
  });
  const best = [...rows].sort((a, b) => b.rate - a.rate)[0] || { symbol: "", last: { invested: 0, value: 0 }, rate: 0, drawdown: 0 };
  const worstDrawdown = rows.reduce((worst, row) => Math.min(worst, row.drawdown), 0);
  const years = yearsBetween(state.dcaStartDate, state.dcaEndDate);
  $("#dcaInvested").textContent = currency.format(best.last.invested);
  $("#dcaValue").textContent = best.symbol ? `${best.symbol} ${currency.format(best.last.value)}` : "--";
  $("#dcaReturn").textContent = `${percent.format(best.rate)} / ${percent.format(annualizedReturn(best.rate, years))}`;
  $("#dcaReturn").className = classForReturn(best.rate);
  $("#dcaMdd").className = "return-negative";
  $("#dcaMdd").textContent = percent.format(worstDrawdown);
  const colors = ["#0f766e", "#a15c07", "#2563eb", "#7c3aed", "#b42318", "#475569", "#15803d"];
  drawLineChart($("#dcaChart"), [
    ...(primary.length ? [{ series: primary.map((point) => ({ date: point.date, value: point.invested })), color: "#667068" }] : []),
    ...seriesBySymbol.map((item, index) => ({ series: item.series, color: colors[index % colors.length] })),
  ]);
  $("#dcaLegend").innerHTML = [
    ...(primary.length ? [{ label: "累積投入", color: "#667068" }] : []),
    ...seriesBySymbol.map((item, index) => ({
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
    .sort((a, b) => b.rate - a.rate)
    .map(({ symbol, last, rate }) => {
      return `
        <div class="rank-row">
          <div>
            <strong>${symbol} ${stocks[symbol].name}</strong>
            <small>${state.dcaStartDate} 到 ${state.dcaEndDate}，累積投入 ${currency.format(last.invested)}</small>
          </div>
          <strong class="${classForReturn(rate)}">${percent.format(rate)}</strong>
        </div>
      `;
    })
    .join("");
}

function renderMarketRanking() {
  const isFullMarket = state.rankingPeriod === "today" || state.rankingPeriod === "1d";
  const sourceSymbols = isFullMarket ? Object.keys(stocks).filter(isTaiwanRankingSymbol) : visibleSymbols().filter(isTaiwanRankingSymbol);
  const sectors = ["全部", ...new Set(sourceSymbols.map((symbol) => stocks[symbol].sector).filter(Boolean))];
  $("#sectorFilter").innerHTML = sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("");
  if (!sectors.includes(state.sector)) state.sector = "全部";
  $("#sectorFilter").value = state.sector;
  $("#rankingNote").textContent = isFullMarket
    ? "今日排行使用證交所與櫃買中心每日全市場收盤資料，含上市櫃普通股、ETF / ETN。"
    : "週 / 月 / 年目前使用你的台股清單比較；全市場歷史排行需要每日歷史資料庫。";

  const rows = sourceSymbols
    .filter((symbol) => state.sector === "全部" || stocks[symbol].sector === state.sector)
    .map((symbol) => ({
      symbol,
      returnRate: periodReturn(symbol, state.rankingPeriod),
      score: continuityScore(symbol),
      marketCap: stocks[symbol].marketCap || 0,
    }))
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
    <div class="analysis-card score-card"><span>高低位分數</span><strong>${risk.highLowScore}</strong><small>0 偏低，100 偏高</small></div>
    <div class="analysis-card"><span>歷史價格百分位</span><strong>${risk.percentile.toFixed(0)}</strong><small>${symbol} ${stock.name} 最新價 ${risk.latest}</small></div>
    <div class="analysis-card"><span>距離 52 週高點</span><strong class="${classForReturn(risk.highDistance)}">${percent.format(risk.highDistance)}</strong><small>高點 ${risk.high}</small></div>
    <div class="analysis-card"><span>距離 52 週低點</span><strong class="${classForReturn(risk.lowDistance)}">${percent.format(risk.lowDistance)}</strong><small>低點 ${risk.low}</small></div>
    <div class="analysis-card"><span>與 200 日均線距離</span><strong class="${classForReturn(risk.maDistance)}">${percent.format(risk.maDistance)}</strong><small>示範版用 8 期均線近似</small></div>
  `;
}

function renderAdvancedMetrics() {
  const realSeries = buildEquitySeries("real");
  const benchmarkSeries = buildEquitySeries(state.benchmark);
  $("#advancedMetrics").innerHTML = `
    <div class="analysis-card"><span>最大跌幅 MDD</span><strong class="return-negative">${percent.format(maxDrawdown(realSeries))}</strong><small>真實風險</small></div>
    <div class="analysis-card"><span>勝率</span><strong>${percent.format(winRate(realSeries))}</strong><small>有多少時間賺錢</small></div>
    <div class="analysis-card"><span>波動率</span><strong>${percent.format(volatility(realSeries))}</strong><small>穩定度</small></div>
    <div class="analysis-card"><span>Sharpe Ratio</span><strong>${sharpeRatio(realSeries).toFixed(2)}</strong><small>報酬 / 風險</small></div>
    <div class="analysis-card"><span>相關性分析</span><strong>${correlation(realSeries, benchmarkSeries).toFixed(2)}</strong><small>和目前比較標的的同步程度</small></div>
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
    { series: buildEquitySeries("real"), color: "#0f766e" },
    { series: buildEquitySeries(state.benchmark), color: "#a15c07" },
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
    state.rankingPeriod = event.target.value;
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
    const amount = Number($("#amountInput").value);
    const price = Number($("#priceInput").value);
    if (!symbol || !date || !amount) return;

    const trade = { id: makeId(), symbol, date, amount, ...(price > 0 ? { price } : {}) };
    trades = [trade, ...trades];
    state.selectedTradeId = trade.id;
    event.target.reset();
    $("#symbolInput").value = symbol;
    $("#dateInput").value = date;
    $("#amountInput").value = amount;
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
