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
  hotMarket: "all",
  hotSector: "全部",
  hotMinTradeValue: 1,
  hotMinReturn: 0,
  hotLimit: 15,
  hotWeights: {
    value: 35,
    momentum: 25,
    high: 20,
    trades: 10,
    relative: 10,
  },
  stockTrendSymbol: "2330",
  stockTrendPeriod: "1m",
  dcaSymbols: ["0050", "006208"],
  dcaFrequency: "monthly",
  dcaAmount: 10000,
  dcaStartDate: "2024-01-15",
  dcaEndDate: "2026-05-01",
  conditionalSymbols: ["2330"],
  conditionalMode: "drop",
  conditionalThreshold: 3,
  conditionalAmount: 10000,
  conditionalStartDate: "2024-01-15",
  conditionalEndDate: "2026-05-01",
  riskSymbol: "2330",
  marketDataSource: "sample",
  marketDataDate: "",
};

const TWSE_DAILY_URL = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data";
const TWSE_COMPANY_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L";
const TPEX_DAILY_URL = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=data";
const TWSE_HISTORY_URL = "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX";
const TPEX_HISTORY_URL = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php";
const FINMIND_DATA_URL = "https://api.finmindtrade.com/api/v4/data";
const REMOTE_MARKET_MANIFEST_URL = "https://raw.githubusercontent.com/klland/stock-analysis/main/data/market-manifest.js";
const REMOTE_MARKET_DATA_URL = "https://raw.githubusercontent.com/klland/stock-analysis/main/data/market-data.js";
const REMOTE_MARKET_HISTORY_URL = "https://raw.githubusercontent.com/klland/stock-analysis/main/data/market-history.js";
const REMOTE_MARKET_HISTORY_CHUNK_BASE_URL = "https://raw.githubusercontent.com/klland/stock-analysis/main/data/history";
const LOCAL_MARKET_HISTORY_URL = "data/market-history.js";
const LOCAL_MARKET_HISTORY_CHUNK_BASE_URL = "data/history";
const CACHE_KEY = "decision-ledger-twse-cache-v1";
const HISTORY_CLOSE_CACHE_KEY = "decision-ledger-history-close-cache-v1";
const ADJUSTED_CLOSE_CACHE_KEY = "decision-ledger-close-cache-v2";
const ADJUSTED_HISTORY_CACHE_KEY = "decision-ledger-close-history-cache-v2";
let marketHistoryPeriods = {};
let marketDcaSnapshots = {};
let marketDcaSeries = {};
let marketHistoryReady = false;
let marketHistoryLoadPromise = null;
let marketHistoryLatestDate = "";
let marketHistoryIsStale = false;
let marketHistoryRefreshing = false;
let marketHistoryAvailableSymbols = new Set();
let marketHistoryPreferRemote = false;
const marketHistoryChunkPromises = new Map();
const marketPriceSeriesCache = new Map();
let marketDataVersion = "";
let marketManifest = null;
const USER_STATE_KEY = "decision-ledger-user-state-v1";
const usEtfSymbols = new Set(["SPY", "QQQ", "VOO", "VTI", "IVV", "SCHD", "VGT", "XLK", "SMH", "SOXX", "DIA", "IWM", "TLT", "BND", "AGG", "IBIT"]);
let userStateLoaded = false;
let singleTradeRenderToken = 0;
let compareRenderToken = 0;
let dcaRenderToken = 0;
let scenarioRenderToken = 0;
let stockTrendRenderToken = 0;
let stockTrendChartState = { points: [], hoverIndex: null };
let interactionToastTimer = null;
let stockSearchActiveIndex = -1;
let stockSearchResults = [];
let stockTrendHoverFrame = 0;
let hotStocksRenderFrame = 0;
let historyRenderToken = 0;
let historySectionObserver = null;

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

const priceFormat = new Intl.NumberFormat("zh-TW", {
  maximumFractionDigits: 2,
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
  status.classList.toggle("is-busy", /正在|載入|取得|補齊/.test(message));
  if (kind !== "info") notifyInteraction(message, kind);
}

function notifyInteraction(message, kind = "info") {
  const toast = $("#interactionToast");
  const text = $("#interactionToastText");
  if (!toast || !text) return;
  text.textContent = message;
  toast.dataset.kind = kind;
  toast.classList.add("is-visible");
  clearTimeout(interactionToastTimer);
  interactionToastTimer = setTimeout(() => toast.classList.remove("is-visible"), kind === "warn" ? 5200 : 3600);
}

function setButtonBusy(button, busy, busyLabel = "處理中") {
  if (!button) return;
  if (busy) {
    button.dataset.originalLabel = button.textContent;
    button.textContent = busyLabel;
    button.disabled = true;
    button.setAttribute("aria-busy", "true");
    button.classList.add("is-busy");
    return;
  }
  if (button.dataset.originalLabel) button.textContent = button.dataset.originalLabel;
  button.disabled = false;
  button.removeAttribute("aria-busy");
  button.classList.remove("is-busy");
}

function showButtonResult(button, label = "已完成", kind = "ok") {
  if (!button?.isConnected) return;
  const originalLabel = button.dataset.originalLabel || button.textContent;
  const restoreDisabled = button.disabled;
  button.textContent = `${kind === "ok" ? "✓ " : ""}${label}`;
  button.classList.remove("is-busy", "is-confirmed", "is-rejected");
  button.classList.add(kind === "ok" ? "is-confirmed" : "is-rejected");
  button.disabled = true;
  window.setTimeout(() => {
    if (!button.isConnected) return;
    button.textContent = originalLabel;
    button.classList.remove("is-confirmed", "is-rejected");
    button.disabled = restoreDisabled;
    delete button.dataset.originalLabel;
  }, kind === "ok" ? 1100 : 1500);
}

function showControlError(control, message) {
  if (control) {
    control.classList.remove("interaction-invalid");
    void control.offsetWidth;
    control.classList.add("interaction-invalid");
    control.setAttribute("aria-invalid", "true");
    window.setTimeout(() => {
      control.classList.remove("interaction-invalid");
      control.removeAttribute("aria-invalid");
    }, 1800);
    control.focus();
  }
  notifyInteraction(message, "warn");
}

function setPanelBusy(selector, busy) {
  const panel = $(selector);
  if (!panel) return;
  panel.classList.toggle("is-calculating", busy);
  panel.setAttribute("aria-busy", String(busy));
}

function controlLabel(control) {
  const label = control.closest("label");
  const text = label?.childNodes?.[0]?.textContent?.trim();
  return text || control.getAttribute("aria-label") || control.name || "設定";
}

function bindInteractionFeedback() {
  document.addEventListener("pointerdown", (event) => {
    const control = event.target.closest("button, a, select, input");
    if (!control || control.disabled) return;
    control.classList.remove("interaction-press");
    void control.offsetWidth;
    control.classList.add("interaction-press");
  });

  ["pointerup", "pointercancel", "pointerleave"].forEach((eventName) => {
    document.addEventListener(eventName, (event) => {
      event.target.closest("button, a, select, input")?.classList.remove("interaction-press");
    });
  });

  document.addEventListener("change", (event) => {
    const control = event.target;
    if (!control.matches("select, input")) return;
    control.classList.add("interaction-updated");
    window.setTimeout(() => control.classList.remove("interaction-updated"), 900);
    notifyInteraction(`${controlLabel(control)} 已更新`, "ok");
  });

  document.addEventListener("click", (event) => {
    const button = event.target.closest("button");
    if (!button || button.disabled || button.classList.contains("is-busy")) return;
    button.classList.remove("interaction-clicked");
    void button.offsetWidth;
    button.classList.add("interaction-clicked");
    window.setTimeout(() => button.classList.remove("interaction-clicked"), 480);
  });
}

function bindNavigationState() {
  const links = [...document.querySelectorAll(".sidebar nav a[href^='#']")];
  const targets = links
    .map((link) => ({ link, section: document.querySelector(link.getAttribute("href")) }))
    .filter((item) => item.section);
  const activate = (activeLink) => {
    links.forEach((link) => {
      const active = link === activeLink;
      link.classList.toggle("active", active);
      if (active) link.setAttribute("aria-current", "location");
      else link.removeAttribute("aria-current");
    });
  };
  links.forEach((link) => link.addEventListener("click", (event) => {
    event.preventDefault();
    const section = document.querySelector(link.getAttribute("href"));
    if (!section) return;
    activate(link);
    window.history.pushState(null, "", link.getAttribute("href"));
    section.scrollIntoView({ behavior: "smooth", block: "start" });

    // Lazy sections above the target can expand while scrolling. Re-anchor after
    // those layouts settle so a navigation click always lands on its section.
    [280, 760].forEach((delay) => {
      window.setTimeout(() => {
        if (window.location.hash === link.getAttribute("href")) {
          section.scrollIntoView({ behavior: "auto", block: "start" });
        }
      }, delay);
    });
  }));
  let scheduled = false;
  window.addEventListener("scroll", () => {
    if (scheduled) return;
    scheduled = true;
    window.requestAnimationFrame(() => {
      const marker = window.scrollY + Math.min(220, window.innerHeight * 0.32);
      const current = targets.filter(({ section }) => section.offsetTop <= marker).at(-1) || targets[0];
      if (current) activate(current.link);
      scheduled = false;
    });
  }, { passive: true });
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

function matchingStocks(value) {
  const query = String(value || "").trim().toLowerCase();
  if (!query) return [];
  return Object.entries(stocks)
    .filter(([symbol, stock]) => isTaiwanSymbol(symbol) || stock.market === "US")
    .map(([symbol, stock]) => {
      const code = symbol.toLowerCase();
      const name = stock.name.toLowerCase();
      let score = Number.POSITIVE_INFINITY;
      if (code === query) score = 0;
      else if (code.startsWith(query)) score = 1;
      else if (name === query) score = 2;
      else if (name.startsWith(query)) score = 3;
      else if (name.includes(query) || `${code} ${name}`.includes(query)) score = 4;
      return { symbol, stock, score };
    })
    .filter((result) => Number.isFinite(result.score))
    .sort(
      (a, b) =>
        a.score - b.score ||
        Number(state.watchlist.includes(a.symbol)) - Number(state.watchlist.includes(b.symbol)) ||
        (b.stock.marketCap || 0) - (a.stock.marketCap || 0) ||
        a.symbol.localeCompare(b.symbol),
    )
    .slice(0, 8);
}

function closeStockSearchSuggestions() {
  const input = $("#stockSearchInput");
  const suggestions = $("#stockSearchSuggestions");
  if (!input || !suggestions) return;
  suggestions.hidden = true;
  suggestions.replaceChildren();
  input.setAttribute("aria-expanded", "false");
  input.removeAttribute("aria-activedescendant");
  stockSearchActiveIndex = -1;
  stockSearchResults = [];
}

function setStockSearchActiveIndex(index) {
  const input = $("#stockSearchInput");
  const options = [...document.querySelectorAll("[data-stock-suggestion]")];
  if (!input || options.length === 0) return;
  stockSearchActiveIndex = (index + options.length) % options.length;
  options.forEach((option, optionIndex) => {
    const active = optionIndex === stockSearchActiveIndex;
    option.classList.toggle("is-active", active);
    option.setAttribute("aria-selected", String(active));
    if (active) {
      input.setAttribute("aria-activedescendant", option.id);
      option.scrollIntoView({ block: "nearest" });
    }
  });
}

function renderStockSearchSuggestions() {
  const input = $("#stockSearchInput");
  const suggestions = $("#stockSearchSuggestions");
  const hint = $("#stockSearchHint");
  if (!input || !suggestions) return;
  const query = input.value.trim();
  if (!query) {
    closeStockSearchSuggestions();
    if (hint) hint.textContent = "";
    return;
  }

  stockSearchResults = matchingStocks(query);
  stockSearchActiveIndex = -1;
  suggestions.replaceChildren();
  if (stockSearchResults.length === 0) {
    const empty = document.createElement("div");
    empty.className = "stock-search-empty";
    empty.textContent = "找不到符合的股票";
    suggestions.append(empty);
  } else {
    stockSearchResults.forEach(({ symbol, stock }, index) => {
      const option = document.createElement("button");
      option.type = "button";
      option.id = `stock-search-option-${index}`;
      option.className = "stock-search-option";
      option.dataset.stockSuggestion = symbol;
      option.setAttribute("role", "option");
      option.setAttribute("aria-selected", "false");

      const identity = document.createElement("span");
      identity.className = "stock-search-identity";
      const title = document.createElement("strong");
      title.textContent = `${symbol} ${stock.name}`;
      const detail = document.createElement("small");
      detail.textContent = `${stock.market || "市場"} · ${stock.sector || "股票"}`;
      identity.append(title, detail);

      const status = document.createElement("span");
      status.className = "stock-search-option-status";
      status.textContent = state.watchlist.includes(symbol) ? "已加入" : "選擇";
      option.append(identity, status);
      suggestions.append(option);
    });
  }
  suggestions.hidden = false;
  input.setAttribute("aria-expanded", "true");
  if (hint) hint.textContent = stockSearchResults.length ? `找到 ${stockSearchResults.length} 筆股票` : "找不到符合的股票";
}

function selectStockSearchSuggestion(symbol) {
  const input = $("#stockSearchInput");
  const stock = stocks[symbol];
  if (!input || !stock) return;
  input.value = `${symbol} ${stock.name}`;
  closeStockSearchSuggestions();
  input.focus();
  notifyInteraction(`已選擇 ${symbol} ${stock.name}`, "ok");
}

function bindStockSearch() {
  const form = $("#watchlistForm");
  const input = $("#stockSearchInput");
  const suggestions = $("#stockSearchSuggestions");
  if (!form || !input || !suggestions) return;

  input.addEventListener("input", renderStockSearchSuggestions);
  input.addEventListener("focus", () => {
    if (input.value.trim()) renderStockSearchSuggestions();
  });
  input.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeStockSearchSuggestions();
      event.preventDefault();
      return;
    }
    if (event.key === "ArrowDown" && stockSearchResults.length) {
      setStockSearchActiveIndex(stockSearchActiveIndex + 1);
      event.preventDefault();
      return;
    }
    if (event.key === "ArrowUp" && stockSearchResults.length) {
      setStockSearchActiveIndex(stockSearchActiveIndex < 0 ? stockSearchResults.length - 1 : stockSearchActiveIndex - 1);
      event.preventDefault();
      return;
    }
    if (event.key === "Enter" && stockSearchActiveIndex >= 0) {
      selectStockSearchSuggestion(stockSearchResults[stockSearchActiveIndex].symbol);
      event.preventDefault();
    }
  });

  suggestions.addEventListener("pointerdown", (event) => event.preventDefault());
  suggestions.addEventListener("pointermove", (event) => {
    const option = event.target.closest("[data-stock-suggestion]");
    if (!option) return;
    const index = [...suggestions.querySelectorAll("[data-stock-suggestion]")].indexOf(option);
    if (index >= 0 && index !== stockSearchActiveIndex) setStockSearchActiveIndex(index);
  });
  suggestions.addEventListener("click", (event) => {
    const option = event.target.closest("[data-stock-suggestion]");
    if (option) selectStockSearchSuggestion(option.dataset.stockSuggestion);
  });

  form.addEventListener("focusout", () => {
    window.setTimeout(() => {
      if (!form.contains(document.activeElement)) closeStockSearchSuggestions();
    }, 0);
  });
  document.addEventListener("pointerdown", (event) => {
    if (!form.contains(event.target)) closeStockSearchSuggestions();
  });
}

function loadUserState() {
  try {
    const saved = JSON.parse(localStorage.getItem(USER_STATE_KEY) || "null");
    if (!saved) return;
    if (Array.isArray(saved.trades)) trades = saved.trades.filter((trade) => trade.id && stocks[trade.symbol]);
    ["watchlist", "portfolioSymbols", "compareSymbols", "dcaSymbols", "conditionalSymbols"].forEach((key) => {
      if (Array.isArray(saved[key])) state[key] = uniqueSymbols(saved[key]);
    });
    ["benchmark", "riskSymbol", "comparePeriod", "compareStartDate", "compareEndDate", "rankingPeriod", "sector", "hotMarket", "hotSector", "stockTrendSymbol", "stockTrendPeriod", "dcaFrequency", "dcaStartDate", "dcaEndDate", "conditionalMode", "conditionalStartDate", "conditionalEndDate"].forEach((key) => {
      if (saved[key]) state[key] = saved[key];
    });
    ["hotMinTradeValue", "hotMinReturn", "hotLimit"].forEach((key) => {
      if (Number.isFinite(Number(saved[key]))) state[key] = Number(saved[key]);
    });
    if (saved.hotWeights && typeof saved.hotWeights === "object") {
      state.hotWeights = { ...state.hotWeights, ...saved.hotWeights };
    }
    if (Number.isFinite(Number(saved.dcaAmount))) state.dcaAmount = Number(saved.dcaAmount);
    if (Number.isFinite(Number(saved.conditionalAmount))) state.conditionalAmount = Number(saved.conditionalAmount);
    if (Number.isFinite(Number(saved.conditionalThreshold))) state.conditionalThreshold = Number(saved.conditionalThreshold);
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
        conditionalSymbols: state.conditionalSymbols,
        benchmark: state.benchmark,
        selectedTradeId: state.selectedTradeId,
        comparePeriod: state.comparePeriod,
        compareStartDate: state.compareStartDate,
        compareEndDate: state.compareEndDate,
        rankingPeriod: state.rankingPeriod,
        sector: state.sector,
        hotMarket: state.hotMarket,
        hotSector: state.hotSector,
        hotMinTradeValue: state.hotMinTradeValue,
        hotMinReturn: state.hotMinReturn,
        hotLimit: state.hotLimit,
        hotWeights: state.hotWeights,
        stockTrendSymbol: state.stockTrendSymbol,
        stockTrendPeriod: state.stockTrendPeriod,
        dcaFrequency: state.dcaFrequency,
        dcaAmount: state.dcaAmount,
        dcaStartDate: state.dcaStartDate,
        dcaEndDate: state.dcaEndDate,
        conditionalMode: state.conditionalMode,
        conditionalThreshold: state.conditionalThreshold,
        conditionalAmount: state.conditionalAmount,
        conditionalStartDate: state.conditionalStartDate,
        conditionalEndDate: state.conditionalEndDate,
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
  state.conditionalSymbols = uniqueSymbols(state.conditionalSymbols.filter((symbol) => list.includes(symbol)));
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

function marketPayloadDate(payload) {
  if (!payload) return "";
  return payload.history?.latestDate || (String(payload.date || "").includes("-") ? payload.date : rocDateToIso(payload.date));
}

function isNewerMarketPayload(candidate, current) {
  const candidateDate = marketPayloadDate(candidate);
  const currentDate = marketPayloadDate(current);
  if (candidateDate !== currentDate) return candidateDate > currentDate;
  return String(candidate?.fetchedAt || "") > String(current?.fetchedAt || "");
}

function parseMarketDataScript(text) {
  const match = String(text || "").match(/^window\.TWSE_MARKET_DATA = (.*);\s*$/s);
  return normalizeMarketPayload(match ? JSON.parse(match[1]) : null);
}

function parseMarketHistoryScript(text) {
  const match = String(text || "").match(/^window\.TWSE_MARKET_HISTORY = (.*);\s*$/s);
  return match ? JSON.parse(match[1]) : null;
}

function parseMarketHistoryChunkScript(text) {
  const match = String(text || "").match(/window\.TWSE_MARKET_HISTORY_CHUNKS\[("[^"]+")\] = (.*);\s*$/s);
  if (!match) return null;
  return { symbol: JSON.parse(match[1]), series: JSON.parse(match[2]) };
}

function parseMarketManifestScript(text) {
  const match = String(text || "").match(/^window\.TWSE_MARKET_MANIFEST = (.*);\s*$/s);
  const manifest = match ? JSON.parse(match[1]) : null;
  return manifest?.version && manifest?.summary && manifest?.history ? manifest : null;
}

async function fetchRemoteMarketManifest() {
  const response = await fetch(`${REMOTE_MARKET_MANIFEST_URL}?t=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Remote market manifest request failed: ${response.status}`);
  const manifest = parseMarketManifestScript(await response.text());
  if (!manifest) throw new Error("Remote market manifest is invalid");
  return manifest;
}

async function fetchRemoteMarketData(manifest) {
  const response = await fetch(`${REMOTE_MARKET_DATA_URL}?v=${encodeURIComponent(manifest?.version || Date.now())}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Remote market data request failed: ${response.status}`);
  const payload = parseMarketDataScript(await response.text());
  if (manifest?.version && payload?.fetchedAt !== manifest.version) throw new Error("Remote market summary version does not match manifest");
  return payload;
}

async function fetchRemoteMarketHistory(manifest) {
  const response = await fetch(`${REMOTE_MARKET_HISTORY_URL}?v=${encodeURIComponent(manifest?.version || Date.now())}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Remote market history request failed: ${response.status}`);
  const payload = parseMarketHistoryScript(await response.text());
  if (manifest?.version && payload?.fetchedAt !== manifest.version) throw new Error("Remote market history version does not match manifest");
  return payload;
}

async function fetchRemoteMarketHistoryChunk(symbol) {
  const response = await fetch(
    `${REMOTE_MARKET_HISTORY_CHUNK_BASE_URL}/${encodeURIComponent(symbol)}.js?v=${encodeURIComponent(marketDataVersion || Date.now())}`,
    { cache: "force-cache" },
  );
  if (!response.ok) throw new Error(`Remote history chunk request failed for ${symbol}: ${response.status}`);
  const chunk = parseMarketHistoryChunkScript(await response.text());
  if (!chunk || chunk.symbol !== symbol || !chunk.series?.points?.length) {
    throw new Error(`Remote history chunk is invalid for ${symbol}`);
  }
  return chunk.series;
}

async function fetchVerifiedRemoteMarketData() {
  let lastError;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const manifest = await fetchRemoteMarketManifest();
      const payload = await fetchRemoteMarketData(manifest);
      return { manifest, payload };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Remote market data is unavailable");
}

function loadLocalMarketHistoryScript() {
  if (window.TWSE_MARKET_HISTORY) return Promise.resolve(window.TWSE_MARKET_HISTORY);
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = `${LOCAL_MARKET_HISTORY_URL}?v=${encodeURIComponent(marketDataVersion || Date.now())}`;
    script.async = true;
    script.onload = () => resolve(window.TWSE_MARKET_HISTORY || null);
    script.onerror = () => reject(new Error("Local market history script unavailable"));
    document.head.append(script);
  });
}

function loadLocalMarketHistoryChunkScript(symbol) {
  const existing = window.TWSE_MARKET_HISTORY_CHUNKS?.[symbol];
  if (existing?.points?.length) return Promise.resolve(existing);
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = `${LOCAL_MARKET_HISTORY_CHUNK_BASE_URL}/${encodeURIComponent(symbol)}.js?v=${encodeURIComponent(marketDataVersion || Date.now())}`;
    script.async = true;
    script.onload = () => {
      const series = window.TWSE_MARKET_HISTORY_CHUNKS?.[symbol];
      if (series?.points?.length) resolve(series);
      else reject(new Error(`Local history chunk is invalid for ${symbol}`));
      script.remove();
    };
    script.onerror = () => {
      script.remove();
      reject(new Error(`Local history chunk unavailable for ${symbol}`));
    };
    document.head.append(script);
  });
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

function normalizeTwseDailyRows(rows) {
  return rows
    .map((row) => ({
      Date: row.Date || row["日期"] || "",
      Code: row.Code || row["證券代號"] || "",
      Name: row.Name || row["證券名稱"] || "",
      TradeVolume: row.TradeVolume || row["成交股數"] || "",
      TradeValue: row.TradeValue || row["成交金額"] || "",
      OpeningPrice: row.OpeningPrice || row["開盤價"] || "",
      HighestPrice: row.HighestPrice || row["最高價"] || "",
      LowestPrice: row.LowestPrice || row["最低價"] || "",
      ClosingPrice: row.ClosingPrice || row["收盤價"] || "",
      Change: row.Change || row["漲跌價差"] || "",
      Transaction: row.Transaction || row["成交筆數"] || "",
    }))
    .filter((row) => row.Date && row.Code && row.ClosingPrice);
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

  const bundledCloses = Object.fromEntries(
    currentHistorySymbols()
      .map((symbol) => [symbol, pricePointOnOrBefore(marketDcaSeries[symbol]?.points || [], targetDate)])
      .filter(([, point]) => point?.value > 0 && point.date >= addDaysIso(targetDate, -14))
      .map(([symbol, point]) => [symbol, point.value]),
  );
  if (Object.keys(bundledCloses).length) {
    return { date: targetDate, closes: bundledCloses, source: "bundled" };
  }

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

function addYearsIso(isoDate, years) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(Date.UTC(year + years, month - 1, day));
  return date.toISOString().slice(0, 10);
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
      value: toNumber(closes[index]) || toNumber(adjusted[index]),
    }))
    .filter((point) => point.date <= targetDate && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  const point = points.at(-1) || null;
  if (point) {
    cache[cacheKey] = { ...point, source: "close" };
    writeAdjustedCloseCache(cache);
    return cache[cacheKey];
  }
  return null;
}

async function adjustedHistoryFor(symbol, startDate, endDate) {
  const cache = readAdjustedHistoryCache();
  const cacheKey = `v4:${symbol}:${startDate}:${endDate}`;
  if (Array.isArray(cache[cacheKey])) return cache[cacheKey];

  if (isTaiwanSymbol(symbol)) {
    try {
      const points = [];
      let cursor = startDate;
      while (cursor <= endDate) {
        const chunkEnd = [addDaysIso(addYearsIso(cursor, 5), -1), endDate].sort()[0];
        const params = new URLSearchParams({
          dataset: "TaiwanStockPrice",
          data_id: symbol,
          start_date: cursor,
          end_date: chunkEnd,
        });
        const result = await fetchJson(`${FINMIND_DATA_URL}?${params}`);
        points.push(
          ...(Array.isArray(result?.data) ? result.data : []).map((row) => ({
            date: row.date,
            value: toNumber(row.close),
            volume: toNumber(row.Trading_Volume),
            source: "finmind",
          })),
        );
        cursor = addDaysIso(chunkEnd, 1);
      }
      const uniquePoints = Object.values(Object.fromEntries(
        points
          .filter((point) => point.date >= startDate && point.date <= endDate && point.value > 0)
          .map((point) => [point.date, point]),
      )).sort((a, b) => a.date.localeCompare(b.date));
      if (uniquePoints.length >= 2) {
        cache[cacheKey] = uniquePoints;
        writeAdjustedHistoryCache(cache);
        return uniquePoints;
      }
    } catch (error) {
      console.warn(`FinMind trend history unavailable for ${symbol}: ${error.message}`);
    }
  }

  const yahooSymbol = yahooSymbolFor(symbol);
  if (!yahooSymbol) return [];

  const start = Math.max(0, Math.floor(new Date(`${addDaysIso(startDate, -10)}T00:00:00Z`).getTime() / 1000));
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
      value: toNumber(closes[index]) || toNumber(adjusted[index]),
    }))
    .filter((point) => point.date >= addDaysIso(startDate, -10) && point.date <= addDaysIso(endDate, 3) && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));

  cache[cacheKey] = points;
  writeAdjustedHistoryCache(cache);
  return points;
}

async function ensureFullHistoryFor(symbol) {
  if (!stocks[symbol]) return false;
  await loadHistorySymbols([symbol]);
  const endDate = state.marketDataDate || priceDates.at(-1);
  const localPoints = localStockTrendPoints(symbol, "2000-01-01", endDate);
  if (!isTaiwanSymbol(symbol)) return localPoints.length >= 2;
  if (localPoints[0]?.date && localPoints[0].date <= "2001-01-31") return true;
  try {
    const points = await adjustedHistoryFor(symbol, "2000-01-01", endDate);
    return points[0]?.date && points[0].date <= "2001-01-31";
  } catch (error) {
    console.warn(`Full history prefetch unavailable for ${symbol}: ${error.message}`);
    return false;
  }
}

async function ensureWatchlistFullHistories(symbols = state.watchlist) {
  const targets = uniqueSymbols(symbols).filter((symbol) => stocks[symbol] && isTaiwanSymbol(symbol));
  for (const symbol of targets) {
    await ensureFullHistoryFor(symbol);
  }
}

function pricePointOnOrAfter(points, date) {
  if (!points.length) return null;
  let low = 0;
  let high = points.length - 1;
  let result = points.length;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (points[middle].date >= date) {
      result = middle;
      high = middle - 1;
    } else {
      low = middle + 1;
    }
  }
  return points[result] || points.at(-1) || null;
}

function pricePointOnOrBefore(points, date) {
  if (!points.length) return null;
  let low = 0;
  let high = points.length - 1;
  let result = -1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (points[middle].date <= date) {
      result = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return result >= 0 ? points[result] : null;
}

function sortedPricePoints(points) {
  return (points || [])
    .filter((point) => point?.date && toNumber(point.value) > 0)
    .map((point) => ({ date: point.date, value: toNumber(point.value), source: point.source || "history" }))
    .sort((a, b) => a.date.localeCompare(b.date));
}

function marketPriceSeries(symbol) {
  if (marketPriceSeriesCache.has(symbol)) return marketPriceSeriesCache.get(symbol);
  const bundled = sortedPricePoints(marketDcaSeries[symbol]?.points || []);
  if (bundled.length) {
    const currentDate = state.marketDataDate || priceDates.at(-1);
    const currentValue = toNumber(stocks[symbol]?.prices?.at(-1));
    if (currentDate && currentValue > 0 && currentDate > bundled.at(-1).date) {
      const series = [...bundled, { date: currentDate, value: currentValue, source: "current" }];
      marketPriceSeriesCache.set(symbol, series);
      return series;
    }
    marketPriceSeriesCache.set(symbol, bundled);
    return bundled;
  }

  const snapshotPoints = Object.values(marketHistoryPeriods)
    .filter((snapshot) => snapshot?.date && toNumber(snapshot.closes?.[symbol]) > 0)
    .map((snapshot) => ({ date: snapshot.date, value: toNumber(snapshot.closes[symbol]), source: "snapshot" }));
  const currentDate = state.marketDataDate || priceDates.at(-1);
  const current = stocks[symbol]?.prices?.at(-1);
  if (currentDate && current > 0) snapshotPoints.push({ date: currentDate, value: current, source: "current" });
  if (snapshotPoints.length) {
    const series = sortedPricePoints(snapshotPoints);
    marketPriceSeriesCache.set(symbol, series);
    return series;
  }

  const series = priceDates
    .map((date, index) => ({ date, value: toNumber(stocks[symbol]?.prices?.[index]), source: "sample" }))
    .filter((point) => point.value > 0);
  marketPriceSeriesCache.set(symbol, series);
  return series;
}

function marketPricePointOnOrBefore(symbol, date) {
  return pricePointOnOrBefore(marketPriceSeries(symbol), date);
}

function marketPricePointNearOrBefore(symbol, date, maxLookbackDays = 14) {
  const point = marketPricePointOnOrBefore(symbol, date);
  if (!point?.date) return null;
  return point.date >= addDaysIso(date, -maxLookbackDays) ? point : null;
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
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    hourCycle: "h23",
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  const date = new Date(`${values.year}-${values.month}-${values.day}T00:00:00+08:00`);
  const hour = Number(values.hour);
  if (hour < 16) date.setDate(date.getDate() - 1);
  while (date.getDay() === 0 || date.getDay() === 6) {
    date.setDate(date.getDate() - 1);
  }
  const tradingParts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const tradingValues = Object.fromEntries(tradingParts.map((part) => [part.type, part.value]));
  return `${tradingValues.year}-${tradingValues.month}-${tradingValues.day}`;
}

function isFreshMarketPayload(payload) {
  return marketPayloadDate(payload) >= latestTradingDate();
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
  const daily = normalizeTwseDailyRows(parseCsv(await dailyResponse.text()));
  const companies = await companyResponse.json();
  const tpexDaily = parseCsv(await tpexResponse.text());
  return { daily, companies, tpexDaily, fetchedAt: new Date().toISOString(), date: rocDateToIso(daily[0]?.Date) };
}

function applyTwseMarketData(payload) {
  const companyMap = new Map(payload.companies.map((item) => [item["公司代號"], item]));
  marketHistoryPeriods = payload.history?.periods || marketHistoryPeriods;
  marketDcaSnapshots = payload.history?.dcaMonthly || marketDcaSnapshots;
  if (Object.keys(payload.history?.dcaSeries || {}).length) {
    marketDcaSeries = payload.history.dcaSeries;
    marketHistoryReady = true;
  }
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
      transactions: toNumber(row.Transaction),
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
      transactions: toNumber(row["成交筆數"] || row["筆數"]),
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

function applyMarketPayload(payload) {
  if (payload?.fetchedAt && payload.fetchedAt !== marketDataVersion) {
    marketHistoryReady = false;
    marketDcaSeries = {};
    marketDcaSnapshots = {};
    marketHistoryAvailableSymbols = new Set();
    marketHistoryChunkPromises.clear();
    marketPriceSeriesCache.clear();
  }
  marketDataVersion = payload?.fetchedAt || marketDataVersion;
  const added = applyTwseMarketData(payload);
  const tpexAdded = applyTpexMarketData(payload);
  const usAdded = applyUsMarketData(payload);
  hydrateUserStateOnce();
  renderControls();
  render();
  renderDataSummary();
  return { added, tpexAdded, usAdded };
}

function applyMarketHistory(payload, { allowStale = false, refreshing = false } = {}) {
  const history = payload?.history || payload;
  const isCurrentVersion = !marketDataVersion || !payload?.fetchedAt || payload.fetchedAt === marketDataVersion;
  if (!isCurrentVersion && !allowStale) return false;
  if (!history) return false;
  marketHistoryPeriods = { ...marketHistoryPeriods, ...(history.periods || {}) };
  marketDcaSnapshots = { ...marketDcaSnapshots, ...(history.dcaMonthly || {}) };
  marketDcaSeries = { ...marketDcaSeries, ...(history.dcaSeries || {}) };
  marketPriceSeriesCache.clear();
  marketHistoryAvailableSymbols = new Set([
    ...marketHistoryAvailableSymbols,
    ...(history.symbols || []),
    ...Object.keys(history.dcaSeries || {}),
  ]);
  marketHistoryLatestDate = history.latestDate || "";
  marketHistoryIsStale = !isCurrentVersion;
  marketHistoryRefreshing = marketHistoryIsStale && refreshing;
  return true;
}

function currentHistorySymbols() {
  return uniqueSymbols([
    ...state.watchlist,
    ...state.portfolioSymbols,
    ...state.compareSymbols,
    ...state.dcaSymbols,
    ...state.conditionalSymbols,
    state.benchmark,
    state.riskSymbol,
    state.stockTrendSymbol,
    ...trades.map((trade) => trade.symbol),
  ]);
}

async function loadHistorySymbol(symbol) {
  if (!symbol || marketDcaSeries[symbol]?.points?.length) return true;
  if (marketHistoryAvailableSymbols.size && !marketHistoryAvailableSymbols.has(symbol)) return false;
  if (marketHistoryChunkPromises.has(symbol)) return marketHistoryChunkPromises.get(symbol);

  const promise = (async () => {
    const loaders = marketHistoryPreferRemote
      ? [() => fetchRemoteMarketHistoryChunk(symbol), () => loadLocalMarketHistoryChunkScript(symbol)]
      : [() => loadLocalMarketHistoryChunkScript(symbol), () => fetchRemoteMarketHistoryChunk(symbol)];
    for (const load of loaders) {
      try {
        const series = await load();
        if (series?.points?.length) {
          marketDcaSeries[symbol] = series;
          marketPriceSeriesCache.delete(symbol);
          return true;
        }
      } catch (error) {
        console.warn(error.message);
      }
    }
    return false;
  })().finally(() => marketHistoryChunkPromises.delete(symbol));
  marketHistoryChunkPromises.set(symbol, promise);
  return promise;
}

async function loadHistorySymbols(symbols) {
  const targets = uniqueSymbols(symbols).filter((symbol) => !marketHistoryAvailableSymbols.size || marketHistoryAvailableSymbols.has(symbol));
  if (targets.length === 0) return false;
  await mapWithConcurrency(targets, 4, loadHistorySymbol);
  return targets.every((symbol) => marketDcaSeries[symbol]?.points?.length);
}

function queueHistoryRenderTasks(tasks, token, schedule = window.requestAnimationFrame) {
  const runNext = () => {
    if (token !== historyRenderToken || tasks.length === 0) return;
    const result = tasks.shift()();
    if (result?.catch) result.catch((error) => console.warn(error.message));
    if (tasks.length) schedule(runNext);
  };
  schedule(runNext);
}

function scheduleHistoryDependentRender() {
  const token = ++historyRenderToken;
  historySectionObserver?.disconnect();
  historySectionObserver = null;

  queueHistoryRenderTasks([renderScenarioRanking, renderCharts], token);

  const lazySections = [
    ["#portfolio", renderPortfolioEngine],
    ["#compare", renderCompareEngine],
    ["#dca", renderDcaEngine],
    ["#conditional", renderConditionalEngine],
    ["#stock-trend", renderStockTrend],
    ["#advanced", renderAdvancedMetrics],
    ["#what-if", renderSingleTradeAnalysis],
  ]
    .map(([selector, task]) => [$(selector), task])
    .filter(([section]) => section);

  if (!("IntersectionObserver" in window)) {
    const idleSchedule = (callback) => {
      if ("requestIdleCallback" in window) {
        window.requestIdleCallback(callback, { timeout: 1200 });
      } else {
        window.setTimeout(callback, 80);
      }
    };
    queueHistoryRenderTasks(lazySections.map(([, task]) => task), token, idleSchedule);
    return;
  }

  const taskBySection = new Map(lazySections);
  historySectionObserver = new IntersectionObserver((entries, observer) => {
    if (token !== historyRenderToken) {
      observer.disconnect();
      return;
    }
    const tasks = entries
      .filter((entry) => entry.isIntersecting)
      .map((entry) => {
        observer.unobserve(entry.target);
        const task = taskBySection.get(entry.target);
        taskBySection.delete(entry.target);
        return task;
      })
      .filter(Boolean);
    if (tasks.length) queueHistoryRenderTasks(tasks, token);
  }, { rootMargin: "600px 0px" });
  lazySections.forEach(([section]) => historySectionObserver.observe(section));
}

function scheduleMarketHistoryLoad() {
  if (marketHistoryReady || marketHistoryLoadPromise) return;
  const start = () => {
    void loadMarketHistory();
  };
  if ("requestIdleCallback" in window) {
    window.requestIdleCallback(start, { timeout: 1800 });
  } else {
    window.setTimeout(start, 450);
  }
}

async function loadMarketHistory() {
  if (marketHistoryReady) return true;
  if (marketHistoryLoadPromise) return marketHistoryLoadPromise;
  marketHistoryLoadPromise = (async () => {
    let localApplied = false;
    try {
      try {
        const local = await loadLocalMarketHistoryScript();
        const localIsCurrent = !marketDataVersion || !local?.fetchedAt || local.fetchedAt === marketDataVersion;
        localApplied = applyMarketHistory(local, { allowStale: true, refreshing: !localIsCurrent });
        marketHistoryPreferRemote = !localIsCurrent;
      } catch (error) {
        console.warn(`Local market history unavailable: ${error.message}`);
      }
      if (!localApplied) {
        try {
          const manifest = marketManifest || await fetchRemoteMarketManifest();
          const remote = await fetchRemoteMarketHistory(manifest);
          if (applyMarketHistory(remote)) marketHistoryPreferRemote = true;
        } catch (error) {
          console.warn(`Remote market history unavailable: ${error.message}`);
          marketHistoryPreferRemote = false;
        }
      } else {
        marketHistoryPreferRemote = false;
      }
      const targets = currentHistorySymbols();
      const loaded = await loadHistorySymbols(targets);
      marketHistoryReady = loaded && targets.length > 0;
      marketHistoryIsStale = Boolean(marketHistoryLatestDate && state.marketDataDate && marketHistoryLatestDate < state.marketDataDate);
      marketHistoryRefreshing = false;
      renderDataSummary();
      scheduleHistoryDependentRender();
      return marketHistoryReady;
    } finally {
      marketHistoryLoadPromise = null;
    }
  })();
  return marketHistoryLoadPromise;
}

function renderDataSummary() {
  const date = state.marketDataDate || "--";
  const historyState = !marketHistoryReady
    ? "歷史日線背景載入中"
    : marketHistoryRefreshing
      ? `歷史日線至 ${marketHistoryLatestDate || "最近交易日"}，正在補齊`
      : marketHistoryIsStale
        ? `歷史日線至 ${marketHistoryLatestDate || "最近交易日"}`
        : "歷史日線已就緒";
  const freshness = $("#dataFreshness");
  const detail = $("#dataDetail");
  const sidebarNote = $("#sidebarDataNote");
  if (freshness) freshness.textContent = `資料日 ${date}`;
  if (detail) detail.textContent = `${historyState} · 台股與美股收盤價`;
  if (sidebarNote) sidebarNote.textContent = `${date} · ${historyState}`;
}

async function loadDailyMarketData() {
  const today = new Date().toISOString().slice(0, 10);
  const bundled = normalizeMarketPayload(window.TWSE_MARKET_DATA);
  const cache = readCache();
  const initial = [bundled, cache].filter(Boolean).sort((a, b) => (isNewerMarketPayload(a, b) ? -1 : 1))[0];
  const initialNeedsApply = initial && (!marketDataVersion || initial.fetchedAt !== marketDataVersion || !state.marketDataDate);
  if (initialNeedsApply) {
    const { added, tpexAdded, usAdded } = applyMarketPayload(initial);
    setMarketStatus(`已先顯示 ${state.marketDataDate} 的可用資料，正在背景確認今日更新。TWSE ${added}、TPEx ${tpexAdded}、美股 ${usAdded} 檔。`, "ok");
  }
  try {
    setMarketStatus("正在確認 GitHub 的最新收盤資料...");
    const remoteResult = await fetchVerifiedRemoteMarketData();
    marketManifest = remoteResult.manifest;
    const remote = remoteResult.payload;
    if (remote?.fetchedAt && (!initial?.fetchedAt || isNewerMarketPayload(remote, initial))) {
      const { added, tpexAdded, usAdded } = applyMarketPayload(remote);
      writeCache({ ...remote, cachedAt: today });
      setMarketStatus(`已更新至 ${state.marketDataDate} 收盤資料。TWSE ${added}、TPEx ${tpexAdded}、美股 ${usAdded} 檔；歷史分析會在背景完成。`, "ok");
      scheduleMarketHistoryLoad();
      return;
    }
    if (initial) {
      setMarketStatus(`資料已是最新可用版本：${state.marketDataDate}。歷史分析會在背景完成。`, "ok");
      scheduleMarketHistoryLoad();
      return;
    }
  } catch (error) {
    console.warn(`Remote market data unavailable: ${error.message}`);
    if (initial) {
      setMarketStatus(`已顯示本地最新可用資料：${state.marketDataDate}。背景更新暫時無法連線。`, "warn");
      scheduleMarketHistoryLoad();
      return;
    }
  }

  if (!initial && bundled?.fetchedAt && isFreshMarketPayload(bundled)) {
    const { added, tpexAdded, usAdded } = applyMarketPayload(bundled);
    setMarketStatus(`已載入本地每日資料 ${state.marketDataDate}，TWSE ${added} 檔、TPEx ${tpexAdded} 檔、美股 ${usAdded} 檔；排程會每天收盤後更新。`, "ok");
    scheduleMarketHistoryLoad();
    return;
  }

  if (!initial && cache?.cachedAt === today) {
    const { added, tpexAdded, usAdded } = applyMarketPayload(cache);
    setMarketStatus(`已載入每日快取 ${state.marketDataDate}，TWSE ${added} 檔、TPEx ${tpexAdded} 檔、美股 ${usAdded} 檔。`, "ok");
    scheduleMarketHistoryLoad();
    return;
  }

  try {
    setMarketStatus("正在從證交所與櫃買中心取得每日收盤行情...");
    const payload = await fetchMarketData();
    const cachePayload = { ...payload, usDaily: bundled?.usDaily || cache?.usDaily || [], cachedAt: today };
    writeCache(cachePayload);
    const { added, tpexAdded, usAdded } = applyMarketPayload(cachePayload);
    setMarketStatus(`已更新每日資料 ${state.marketDataDate}，TWSE ${added} 檔、TPEx ${tpexAdded} 檔；美股 ${usAdded} 檔沿用本地快照，會由每日更新腳本刷新。`, "ok");
    scheduleMarketHistoryLoad();
  } catch (error) {
    console.warn(error);
    if (bundled?.fetchedAt) {
      const { added, tpexAdded, usAdded } = applyMarketPayload(bundled);
      setMarketStatus(`即時資料暫時無法載入，先使用本地快照 ${state.marketDataDate}：TWSE ${added} 檔、TPEx ${tpexAdded} 檔、美股 ${usAdded} 檔。`, "warn");
      scheduleMarketHistoryLoad();
      return;
    }
    hydrateUserStateOnce();
    setMarketStatus("每日資料暫時無法載入，已改用內建樣本；重新整理時會再自動嘗試。", "warn");
    scheduleMarketHistoryLoad();
  }
}

function seriesFor(symbol) {
  return priceDates.map((date, index) => [date, stocks[symbol].prices[index]]);
}

function priceOnOrAfter(symbol, date) {
  const point = pricePointOnOrAfter(marketPriceSeries(symbol), date);
  return point?.value || 0;
}

function previousClose(symbol) {
  const stock = stocks[symbol];
  const dailyReturn = stock?.dailyReturn;
  const latest = latestPrice(symbol);
  if (!stock?.live || !Number.isFinite(dailyReturn) || dailyReturn <= -1) return 0;
  return latest / (1 + dailyReturn);
}

function priceOnOrBefore(symbol, date) {
  const point = marketPricePointOnOrBefore(symbol, date);
  if (point?.value > 0) return point.value;
  if (stocks[symbol]?.live && state.marketDataDate && date < state.marketDataDate) {
    const previous = previousClose(symbol);
    if (previous > 0) return previous;
  }
  return 0;
}

function entryPriceForTrade(trade, symbol = trade.symbol) {
  if (symbol === trade.symbol && Number.isFinite(Number(trade.price))) return Number(trade.price);
  return priceOnOrBefore(symbol, trade.date);
}

function priceAtIndex(symbol, index) {
  const date = priceDates[Math.max(0, Math.min(index, priceDates.length - 1))];
  return priceOnOrBefore(symbol, date) || toNumber(stocks[symbol]?.prices?.[Math.max(0, Math.min(index, priceDates.length - 1))]);
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
  const current = toNumber(stocks[symbol]?.prices?.at(-1));
  if (current > 0) return current;
  const latestDate = state.marketDataDate || priceDates.at(-1);
  const point = marketPricePointOnOrBefore(symbol, latestDate);
  return point?.value || 0;
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

function comparePricePoint(symbol, targetDate, snapshot) {
  const bundled = marketPricePointNearOrBefore(symbol, targetDate, 14);
  if (bundled?.value > 0) return bundled;
  const snapshotClose = snapshot?.latest ? latestPrice(symbol) : closeFromHistoricalSnapshot(symbol, snapshot);
  return snapshotClose > 0 ? { date: snapshot.date || targetDate, value: snapshotClose, source: "snapshot" } : null;
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
  if (bundledPoint?.value > 0 && bundledPoint.date <= targetDate) return { date: bundledPoint.date, value: bundledPoint.value, source: "close" };
  if (marketHistoryReady) return null;
  try {
    const adjusted = await adjustedCloseOnOrBefore(symbol, targetDate);
    if (adjusted?.value > 0) return adjusted;
  } catch (error) {
    console.warn(`Scenario close unavailable for ${symbol}: ${error.message}`);
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
  if (bundledPoint?.value > 0 && bundledPoint.date <= targetDate) return { date: bundledPoint.date, value: bundledPoint.value, source: "close" };
  if (marketHistoryReady) {
    const raw = toNumber(snapshot?.closes?.[symbol]);
    return raw > 0 ? { date: snapshot.date, value: raw, source: "raw" } : null;
  }
  try {
    const adjusted = await adjustedCloseOnOrBefore(symbol, targetDate);
    if (adjusted?.value > 0) return adjusted;
  } catch (error) {
    console.warn(`Close unavailable for ${symbol}: ${error.message}`);
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
  const symbols = symbolMode === "real" ? uniqueSymbols(trades.map((trade) => trade.symbol)) : [symbolMode];
  const dates = [...new Set(symbols.flatMap((symbol) => marketPriceSeries(symbol).map((point) => point.date)).filter((date) => trades.some((trade) => trade.date <= date)))]
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
  return marketPriceSeries(symbol);
}

function closeFromSnapshot(symbol, point) {
  if (point.source === "current") return latestPrice(symbol);
  return toNumber(point.closes?.[symbol]);
}

function portfolioSnapshotSeries(symbolMode = "real") {
  const dates = [
    ...new Set(
      trades
        .flatMap((trade) => marketPriceSeries(symbolMode === "real" ? trade.symbol : symbolMode).map((point) => point.date))
        .filter((date) => trades.some((trade) => trade.date <= date)),
    ),
  ].sort((a, b) => a.localeCompare(b));

  return dates
    .map((date) => {
      const value = trades.reduce((sum, trade) => {
        if (trade.date > date) return sum;
        const symbol = symbolMode === "real" ? trade.symbol : symbolMode;
        const price = priceOnOrBefore(symbol, date);
        if (!(price > 0)) return sum;
        return sum + sharesFor(trade, symbol) * price;
      }, 0);
      return { date, value };
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
  const series = buildScenarioEquitySeries("real");
  return {
    totalInvested,
    currentValue,
    totalReturn: totalInvested ? currentValue / totalInvested - 1 : 0,
    drawdown: maxDrawdown(series),
  };
}

function portfolioWeights() {
  const inputs = [...document.querySelectorAll("[data-weight]")];
  return inputs.map((input) => {
    const percentValue = Math.max(0, Math.min(100, Number(input.value) || 0));
    return { symbol: input.dataset.weight, percent: percentValue, weight: percentValue / 100 };
  });
}

function syncPortfolioWeightsFromInputs() {
  state.portfolioWeights = Object.fromEntries(
    [...document.querySelectorAll("[data-weight]")].map((input) => [input.dataset.weight, Number(input.value) || 0]),
  );
}

function portfolioWeightTotal(weights = portfolioWeights()) {
  return weights.reduce((sum, item) => sum + item.weight, 0);
}

function portfolioStartDate(weights) {
  const firstDates = weights
    .filter((item) => item.weight > 0)
    .map((item) => marketPriceSeries(item.symbol).find((point) => point.value > 0)?.date)
    .filter(Boolean);
  return firstDates.length ? firstDates.sort((a, b) => b.localeCompare(a))[0] : "";
}

function portfolioDetailRows(amount, weights, startDate, endDate) {
  return weights
    .filter((item) => item.weight > 0)
    .map((item) => {
      const startPoint = marketPricePointOnOrBefore(item.symbol, startDate);
      const endPoint = marketPricePointOnOrBefore(item.symbol, endDate);
      const startPrice = toNumber(startPoint?.value);
      const endPrice = toNumber(endPoint?.value);
      const allocated = amount * item.weight;
      const value = startPrice > 0 && endPrice > 0 ? allocated * (endPrice / startPrice) : 0;
      return {
        ...item,
        startDate: startPoint?.date || startDate,
        endDate: endPoint?.date || endDate,
        startPrice,
        endPrice,
        allocated,
        value,
        returnRate: allocated > 0 && value > 0 ? value / allocated - 1 : NaN,
      };
    });
}

function buildWeightedPortfolio(amount) {
  const weights = portfolioWeights();
  const weightSum = Math.min(1, portfolioWeightTotal(weights));
  const startDate = portfolioStartDate(weights);
  if (!startDate) return { series: [], weights, weightSum, startDate: "", endDate: "", details: [] };
  const dates = [
    ...new Set(weights.flatMap((item) => marketPriceSeries(item.symbol).map((point) => point.date)).filter((date) => date >= startDate)),
  ].sort((a, b) => a.localeCompare(b));
  const cash = amount * Math.max(0, 1 - weightSum);
  const series = dates.map((date) => {
    const value = weights.reduce((sum, item) => {
      if (item.weight <= 0) return sum;
      const startPrice = priceOnOrBefore(item.symbol, dates[0]);
      const currentPrice = priceOnOrBefore(item.symbol, date);
      if (!(startPrice > 0) || !(currentPrice > 0)) return sum;
      return sum + amount * item.weight * (currentPrice / startPrice);
    }, cash);
    return { date, value };
  }).filter((point) => point.value > 0);
  const endDate = series.at(-1)?.date || dates.at(-1) || "";
  return { series, weights, weightSum, startDate, endDate, details: portfolioDetailRows(amount, weights, startDate, endDate), cash };
}

function buildWeightedPortfolioSeries(amount) {
  return buildWeightedPortfolio(amount).series;
}

function enforcePortfolioWeightLimit(changedInput) {
  const inputs = [...document.querySelectorAll("[data-weight]")];
  if (!changedInput) return;
  const otherTotal = inputs
    .filter((input) => input !== changedInput)
    .reduce((sum, input) => sum + Math.max(0, Math.min(100, Number(input.value) || 0)), 0);
  const maxAllowed = Math.max(0, 100 - otherTotal);
  const nextValue = Math.max(0, Math.min(maxAllowed, Number(changedInput.value) || 0));
  changedInput.value = Number.isInteger(nextValue) ? String(nextValue) : nextValue.toFixed(2);
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

function conditionalModeText() {
  return state.conditionalMode === "rise" ? "單日上漲" : "單日下跌";
}

function conditionalTriggered(dailyReturn) {
  const threshold = Math.max(0, Number(state.conditionalThreshold) || 0) / 100;
  return state.conditionalMode === "rise" ? dailyReturn >= threshold : dailyReturn <= -threshold;
}

function buildConditionalStrategy(symbol) {
  const points = marketPriceSeries(symbol)
    .filter((point) => point.date >= state.conditionalStartDate && point.date <= state.conditionalEndDate)
    .sort((a, b) => a.date.localeCompare(b.date));
  if (points.length < 2) {
    return { symbol, points, series: [], triggers: [], dailyReturns: [], error: "這段時間沒有足夠日線資料" };
  }

  let invested = 0;
  let shares = 0;
  const amount = Math.max(0, Number(state.conditionalAmount) || 0);
  const triggers = [];
  const series = [];
  const dailyReturns = [];
  const cashflows = [];

  for (let index = 1; index < points.length; index += 1) {
    const previous = points[index - 1];
    const point = points[index];
    const dailyReturn = previous.value > 0 ? point.value / previous.value - 1 : 0;
    dailyReturns.push({ date: point.date, value: dailyReturn, close: point.value });
    if (amount > 0 && conditionalTriggered(dailyReturn)) {
      const boughtShares = amount / point.value;
      invested += amount;
      shares += boughtShares;
      cashflows.push({ date: point.date, amount: -amount });
      triggers.push({
        date: point.date,
        close: point.value,
        dailyReturn,
        amount,
        shares: boughtShares,
      });
    }
    if (invested > 0) series.push({ date: point.date, invested, value: shares * point.value });
  }

  const last = series.at(-1) || { date: points.at(-1).date, invested: 0, value: 0 };
  if (last.value > 0) cashflows.push({ date: last.date, amount: last.value });
  const periodReturn = points[0].value > 0 ? points.at(-1).value / points[0].value - 1 : NaN;
  const bestDay = dailyReturns.reduce((best, item) => (item.value > best.value ? item : best), dailyReturns[0]);
  const worstDay = dailyReturns.reduce((worst, item) => (item.value < worst.value ? item : worst), dailyReturns[0]);
  return {
    symbol,
    points,
    series,
    triggers,
    dailyReturns,
    cashflows,
    last,
    periodReturn,
    bestDay,
    worstDay,
    returnRate: last.invested > 0 ? last.value / last.invested - 1 : NaN,
    annualized: xirr(cashflows),
    error: triggers.length ? "" : "這段期間沒有觸發買入",
  };
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

function clamp(value, min = 0, max = 1) {
  return Math.max(min, Math.min(max, value));
}

function hotStockUniverse() {
  return Object.keys(stocks)
    .filter(isTaiwanRankingSymbol)
    .filter((symbol) => state.hotMarket === "all" || stocks[symbol].market === state.hotMarket)
    .filter((symbol) => state.hotSector === "全部" || stocks[symbol].sector === state.hotSector);
}

function hotStockSectors() {
  const sourceSymbols = Object.keys(stocks)
    .filter(isTaiwanRankingSymbol)
    .filter((symbol) => state.hotMarket === "all" || stocks[symbol].market === state.hotMarket);
  return ["全部", ...new Set(sourceSymbols.map((symbol) => stocks[symbol].sector).filter(Boolean))];
}

function hotShortReturn(symbol) {
  const points = marketPriceSeries(symbol);
  if (points.length >= 6) {
    const start = points.at(-6);
    const end = points.at(-1);
    return start?.value > 0 ? end.value / start.value - 1 : 0;
  }
  const weekly = rankingReturn(symbol, "1w");
  return Number.isFinite(weekly) ? weekly : stocks[symbol]?.dailyReturn || 0;
}

function hotHighPosition(symbol) {
  const points = marketPriceSeries(symbol).slice(-20);
  const latest = latestPrice(symbol);
  const high = Math.max(...points.map((point) => point.value), latest);
  if (!(high > 0) || !(latest > 0)) return 0;
  return clamp(latest / high);
}

function hotRelativeStrength(symbol) {
  const base = hotShortReturn(state.benchmark) || 0;
  return hotShortReturn(symbol) - base;
}

function hotWeightTotal() {
  return Object.values(state.hotWeights).reduce((sum, value) => sum + Math.max(0, Number(value) || 0), 0) || 1;
}

function scoreFromTradeValue(value, maxValue) {
  if (!(value > 0) || !(maxValue > 0)) return 0;
  return clamp(Math.log10(value + 1) / Math.log10(maxValue + 1));
}

function scoreFromTransactionCount(value, maxValue) {
  if (!(value > 0) || !(maxValue > 0)) return 0;
  return clamp(Math.log10(value + 1) / Math.log10(maxValue + 1));
}

function scoreFromReturn(value) {
  return clamp((value + 0.02) / 0.14);
}

function scoreFromRelative(value) {
  return clamp((value + 0.03) / 0.12);
}

function hotStockRows() {
  const symbols = hotStockUniverse();
  const maxTradeValue = Math.max(...symbols.map((symbol) => stocks[symbol]?.tradeValue || 0), 1);
  const maxTransactions = Math.max(...symbols.map((symbol) => stocks[symbol]?.transactions || 0), 1);
  const weightTotal = hotWeightTotal();
  const minTradeValue = Math.max(0, Number(state.hotMinTradeValue) || 0) * 100_000_000;
  const minReturn = (Number(state.hotMinReturn) || 0) / 100;

  return symbols
    .map((symbol) => {
      const stock = stocks[symbol];
      const shortReturn = hotShortReturn(symbol);
      const relativeStrength = hotRelativeStrength(symbol);
      const highPosition = hotHighPosition(symbol);
      const valueScore = scoreFromTradeValue(stock.tradeValue || 0, maxTradeValue);
      const momentumScore = scoreFromReturn(shortReturn);
      const highScore = highPosition;
      const tradesScore = scoreFromTransactionCount(stock.transactions || 0, maxTransactions);
      const relativeScore = scoreFromRelative(relativeStrength);
      const weighted =
        valueScore * state.hotWeights.value +
        momentumScore * state.hotWeights.momentum +
        highScore * state.hotWeights.high +
        tradesScore * state.hotWeights.trades +
        relativeScore * state.hotWeights.relative;
      return {
        symbol,
        shortReturn,
        relativeStrength,
        highPosition,
        tradeValue: stock.tradeValue || 0,
        valueScore,
        transactions: stock.transactions || 0,
        score: Math.round((weighted / weightTotal) * 100),
        hasDailySeries: (marketDcaSeries[symbol]?.points || []).length >= 6,
      };
    })
    .filter((row) => row.tradeValue >= minTradeValue && row.shortReturn >= minReturn)
    .sort((a, b) => b.score - a.score || b.tradeValue - a.tradeValue)
    .slice(0, Math.max(5, Math.min(50, Number(state.hotLimit) || 15)));
}

function stockTrendPeriodLabel(period = state.stockTrendPeriod) {
  return {
    "1d": "1日",
    "1w": "1週",
    "1m": "1月",
    "3m": "3月",
    "6m": "半年",
    "1y": "1年",
    "5y": "5年",
    all: "全部",
  }[period] || period;
}

function stockTrendWindow(period = state.stockTrendPeriod) {
  const endDate = state.marketDataDate || priceDates.at(-1);
  if (period === "all") return { startDate: "2000-01-01", endDate, label: "全部" };
  const lookbackDays = {
    "1d": 1,
    "1w": 7,
    "1m": 30,
    "3m": 90,
    "6m": 183,
    "1y": 365,
    "5y": 365 * 5,
  }[period] || 30;
  return { startDate: addDaysIso(endDate, -lookbackDays), endDate, label: stockTrendPeriodLabel(period) };
}

function localStockTrendPoints(symbol, startDate, endDate) {
  const points = marketPriceSeries(symbol)
    .filter((point) => point.date >= addDaysIso(startDate, -14) && point.date <= endDate && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));

  if (points.length >= 2) return points.filter((point) => point.date >= startDate);

  if (state.marketDataDate && stocks[symbol]?.live && (state.stockTrendPeriod === "1d" || state.stockTrendPeriod === "1w")) {
    const previous = previousClose(symbol);
    const latest = latestPrice(symbol);
    if (previous > 0 && latest > 0) {
      return [
        { date: "前一交易日", value: previous, source: "daily" },
        { date: state.marketDataDate, value: latest, source: "daily" },
      ];
    }
  }

  return points;
}

async function stockTrendPoints(symbol, period = state.stockTrendPeriod) {
  await loadHistorySymbols([symbol]);
  const { startDate, endDate } = stockTrendWindow(period);
  const localPoints = localStockTrendPoints(symbol, startDate, endDate);
  const needsLongHistory = period === "3m" || period === "6m" || period === "1y" || period === "5y" || period === "all";
  const localCoverageStart = localPoints[0]?.date || "";
  const hasEnoughCoverage = localCoverageStart && localCoverageStart <= addDaysIso(startDate, 14);
  if ((period === "5y" || period === "all") && !hasEnoughCoverage) {
    // Force a full Yahoo history fetch; the bundled local series only keeps a compact recent window.
  } else if (localPoints.length >= 6 || (!needsLongHistory && localPoints.length >= 2)) {
    return { points: localPoints, startDate, endDate, source: localPoints.some((point) => point.source === "history") ? "本地日線" : "本地快照" };
  }

  try {
    const fetched = await adjustedHistoryFor(symbol, startDate, endDate);
    const points = fetched.filter((point) => point.date >= startDate && point.date <= endDate && point.value > 0);
    if (points.length >= 2) return { points, startDate, endDate, source: "Yahoo 日線" };
  } catch (error) {
    console.warn(`Trend history unavailable for ${symbol}: ${error.message}`);
  }

  return { points: localPoints, startDate, endDate, source: localPoints.length ? "本地快照" : "無資料" };
}

function stockTrendTimeLabel(symbol) {
  return stocks[symbol]?.market === "US" ? "16:00 收盤" : "13:30 收盤";
}

function drawPriceChart(canvas, points, options = {}) {
  const ctx = canvas.getContext("2d");
  const values = points.map((point) => point.value).filter((value) => value > 0);
  canvas.height = 360;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.font = "13px system-ui";

  if (values.length < 2) {
    ctx.fillStyle = "#667068";
    ctx.fillText("這個區間沒有足夠價格資料。", 24, 44);
    return;
  }

  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = Math.max(0.01, max - min);
  const padding = { top: 18, right: 18, bottom: 34, left: 18 };
  const width = canvas.width - padding.left - padding.right;
  const height = canvas.height - padding.top - padding.bottom;
  const isPositive = points.at(-1).value >= points[0].value;
  const lineColor = isPositive ? "#0f8f4f" : "#d93025";
  const hoverIndex = Number.isInteger(options.hoverIndex) ? Math.max(0, Math.min(points.length - 1, options.hoverIndex)) : null;

  ctx.strokeStyle = "#eef1ee";
  ctx.lineWidth = 1;
  for (let index = 1; index <= 3; index += 1) {
    const y = padding.top + height * (index / 4);
    ctx.beginPath();
    ctx.moveTo(padding.left, y);
    ctx.lineTo(canvas.width - padding.right, y);
    ctx.stroke();
  }

  const coords = points.map((point, index) => {
    const x = padding.left + width * (index / Math.max(1, points.length - 1));
    const y = padding.top + height * (1 - (point.value - min) / range);
    return { x, y };
  });

  const gradient = ctx.createLinearGradient(0, padding.top, 0, canvas.height - padding.bottom);
  gradient.addColorStop(0, isPositive ? "rgba(15, 143, 79, 0.18)" : "rgba(217, 48, 37, 0.18)");
  gradient.addColorStop(1, "rgba(255, 255, 255, 0)");
  ctx.beginPath();
  coords.forEach((point, index) => {
    if (index === 0) ctx.moveTo(point.x, point.y);
    else ctx.lineTo(point.x, point.y);
  });
  ctx.lineTo(coords.at(-1).x, canvas.height - padding.bottom);
  ctx.lineTo(coords[0].x, canvas.height - padding.bottom);
  ctx.closePath();
  ctx.fillStyle = gradient;
  ctx.fill();

  ctx.strokeStyle = lineColor;
  ctx.lineWidth = 3.5;
  ctx.lineJoin = "round";
  ctx.lineCap = "round";
  ctx.beginPath();
  coords.forEach((point, index) => {
    if (index === 0) ctx.moveTo(point.x, point.y);
    else ctx.lineTo(point.x, point.y);
  });
  ctx.stroke();

  ctx.fillStyle = "#667068";
  ctx.fillText(points[0].date, padding.left, canvas.height - 10);
  ctx.textAlign = "right";
  ctx.fillText(points.at(-1).date, canvas.width - padding.right, canvas.height - 10);
  ctx.textAlign = "left";

  if (hoverIndex !== null) {
    const point = points[hoverIndex];
    const coord = coords[hoverIndex];
    const baseValue = options.baseValue || points[0].value;
    const pointReturn = baseValue > 0 ? point.value / baseValue - 1 : NaN;
    const tooltipLines = [
      `${point.date} ${stockTrendTimeLabel(options.symbol)}`,
      `${priceFormat.format(point.value)}  ${Number.isFinite(pointReturn) ? percent.format(pointReturn) : "--"}`,
    ];
    const boxWidth = 190;
    const boxHeight = 58;
    const boxX = Math.max(12, Math.min(canvas.width - boxWidth - 12, coord.x - boxWidth / 2));
    const boxY = coord.y > 92 ? coord.y - boxHeight - 18 : coord.y + 18;

    ctx.strokeStyle = "rgba(24, 32, 27, 0.28)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(coord.x, padding.top);
    ctx.lineTo(coord.x, canvas.height - padding.bottom);
    ctx.stroke();

    ctx.fillStyle = lineColor;
    ctx.beginPath();
    ctx.arc(coord.x, coord.y, 4.5, 0, Math.PI * 2);
    ctx.fill();

    ctx.fillStyle = "rgba(24, 32, 27, 0.92)";
    ctx.beginPath();
    ctx.roundRect(boxX, boxY, boxWidth, boxHeight, 10);
    ctx.fill();
    ctx.fillStyle = "#ffffff";
    ctx.font = "12px system-ui";
    ctx.fillText(tooltipLines[0], boxX + 12, boxY + 22);
    ctx.font = "700 15px system-ui";
    ctx.fillText(tooltipLines[1], boxX + 12, boxY + 43);
  }
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
    ctx.fillText(options.emptyMessage || "這個日期區間沒有可用資料。", 24, 44);
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
    .map((symbol) => {
      const stock = stocks[symbol];
      const returnRate = stock.dailyReturn;
      return `
        <article class="watchlist-card">
          <button class="watchlist-stock" type="button" data-select-stock="${symbol}">
            <strong>${symbol} ${stock.name}</strong>
            <small>${stock.market} · ${stock.sector}</small>
          </button>
          <div class="watchlist-quote">
            <strong>${priceFormat.format(latestPrice(symbol))}</strong>
            <small class="${classForReturn(returnRate)}">${returnRate >= 0 ? "+" : ""}${percent.format(returnRate)} 今日</small>
          </div>
          <button class="watchlist-remove" type="button" aria-label="刪除 ${symbol}" data-remove-watchlist="${symbol}">×</button>
        </article>
      `;
    })
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

function renderConditionalList() {
  $("#conditionalChips").innerHTML = state.conditionalSymbols
    .map(
      (symbol) => `
        <span class="selected-chip">
          ${symbol} ${stocks[symbol].name}
          <button type="button" aria-label="刪除 ${symbol}" data-remove-conditional="${symbol}">×</button>
        </span>
      `,
    )
    .join("");
}

function renderControls() {
  syncSelections();
  normalizeRankingPeriod();
  const list = visibleSymbols();
  const options = optionHtml(list);
  $("#benchmarkSelect").innerHTML = options;
  $("#symbolInput").innerHTML = options;
  $("#riskSymbol").innerHTML = options;
  $("#stockTrendSymbol").innerHTML = options;
  $("#portfolioAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.portfolioSymbols.includes(symbol)));
  $("#compareAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.compareSymbols.includes(symbol)));
  $("#dcaAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.dcaSymbols.includes(symbol)));
  $("#conditionalAddSelect").innerHTML = optionHtml(list.filter((symbol) => !state.conditionalSymbols.includes(symbol)));
  $("#portfolioAddForm button").disabled = !$("#portfolioAddSelect").value;
  $("#compareAddForm button").disabled = !$("#compareAddSelect").value;
  $("#dcaAddForm button").disabled = !$("#dcaAddSelect").value;
  $("#conditionalAddForm button").disabled = !$("#conditionalAddSelect").value;
  $("#benchmarkSelect").value = state.benchmark;
  $("#riskSymbol").value = state.riskSymbol;
  if (!stocks[state.stockTrendSymbol]) state.stockTrendSymbol = list[0] || "2330";
  $("#stockTrendSymbol").value = state.stockTrendSymbol;
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
  $("#conditionalMode").value = state.conditionalMode;
  $("#conditionalThreshold").value = state.conditionalThreshold;
  $("#conditionalAmount").value = state.conditionalAmount;
  $("#conditionalStartDate").value = state.conditionalStartDate;
  $("#conditionalEndDate").value = state.conditionalEndDate;

  const sectors = ["全部", ...new Set(Object.values(stocks).map((stock) => stock.sector).filter(Boolean))];
  $("#sectorFilter").innerHTML = sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("");
  $("#sectorFilter").value = sectors.includes(state.sector) ? state.sector : "全部";

  renderWatchlist();
  renderCompareList();
  renderDcaList();
  renderConditionalList();

  const defaultWeights = { "2330": 35, "0050": 35, "2454": 20, "2881": 10, "2317": 10 };
  let assignedWeight = 0;
  $("#portfolioWeights").innerHTML = state.portfolioSymbols
    .map((symbol) => {
      const rawWeight = Math.max(0, Math.min(100, Number(state.portfolioWeights[symbol] ?? defaultWeights[symbol] ?? 0)));
      const weight = Math.min(rawWeight, Math.max(0, 100 - assignedWeight));
      assignedWeight += weight;
      state.portfolioWeights[symbol] = weight;
      return `
        <label class="weight-item">
          <span>${symbol} ${stocks[symbol].name}</span>
          <input data-weight="${symbol}" type="number" min="0" max="100" step="5" value="${weight}" />
          <button class="delete-btn" type="button" data-remove-portfolio="${symbol}">刪除</button>
        </label>
      `;
    })
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

  $("#scenarioRanking").innerHTML = `<div class="rank-row"><strong>正在重算替代結果</strong><small>使用每筆交易日期的每日收盤價買入，而不是範例價格序列。</small></div>`;
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
          : `${item.entryDate} ${item.entrySource === "close" ? "收盤價" : "原始收盤"} ${item.entryPrice.toFixed(2)}`;
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
  const portfolio = buildWeightedPortfolio(amount);
  const { series, weightSum, startDate, endDate, details, cash } = portfolio;
  if (!series.length) {
    $("#portfolioBasis").textContent = "請先加入至少一檔股票並設定權重。";
    $("#portfolioResults").innerHTML = "";
    return;
  }
  const finalValue = series.at(-1).value;
  const totalReturn = amount ? finalValue / amount - 1 : 0;
  const weightPercent = Math.round(weightSum * 1000) / 10;
  const unallocatedPercent = Math.max(0, 100 - weightPercent);
  $("#portfolioBasis").textContent =
    `計算區間：${startDate} 到 ${endDate}。權重加總上限為 100%，目前已配置 ${weightPercent}%` +
    `${unallocatedPercent ? `，未配置 ${unallocatedPercent}% 以現金保留` : ""}。總報酬 = 期末組合市值 / 投入金額 - 1。`;

  $("#portfolioResults").innerHTML = `
    <div class="analysis-card"><span>目前組合市值</span><strong>${currency.format(finalValue)}</strong><small>股票市值 + 現金 ${currency.format(cash || 0)}</small></div>
    <div class="analysis-card"><span>總報酬</span><strong class="${classForReturn(totalReturn)}">${percent.format(totalReturn)}</strong><small>${startDate} 到 ${endDate}</small></div>
    <div class="analysis-card"><span>最大跌幅</span><strong class="return-negative">${percent.format(maxDrawdown(series))}</strong><small>歷史資產高點到低點</small></div>
    <div class="analysis-card"><span>波動度</span><strong>${percent.format(volatility(series))}</strong><small>時間序列標準差</small></div>
    <div class="analysis-card"><span>勝率</span><strong>${percent.format(winRate(series))}</strong><small>有多少期間為正報酬</small></div>
    <div class="analysis-card"><span>Sharpe Ratio</span><strong>${sharpeRatio(series).toFixed(2)}</strong><small>報酬 / 風險</small></div>
    ${details
      .map(
        (row) => `
          <div class="analysis-card">
            <span>${row.percent}% · ${row.symbol}</span>
            <strong class="${classForReturn(row.returnRate)}">${metricValue(row.returnRate)}</strong>
            <small>${row.startDate} ${row.startPrice.toFixed(2)} → ${row.endDate} ${row.endPrice.toFixed(2)}</small>
          </div>
        `,
      )
      .join("")}
  `;
}

async function renderCompareEngine() {
  const token = ++compareRenderToken;
  setPanelBusy("#compare", true);
  const customFields = $("#compareCustomFields");
  customFields.hidden = state.comparePeriod !== "custom";
  if (state.comparePeriod === "custom" && state.compareStartDate > state.compareEndDate) {
    $("#compareBasis").textContent = "日期區間錯誤：開始日期必須早於結束日期。";
    drawBarChart($("#compareBarChart"), []);
    $("#compareBreakdown").innerHTML = "";
    setPanelBusy("#compare", false);
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
    setPanelBusy("#compare", false);
    return;
  }

  const rows = state.compareSymbols
    .map((symbol) => {
      const startPoint = comparePricePoint(symbol, startTarget, startSnapshot);
      const endPoint = comparePricePoint(symbol, endTarget, endSnapshot);
      const startPrice = toNumber(startPoint?.value);
      const endPrice = toNumber(endPoint?.value);
      return {
        symbol,
        startDate: startPoint?.date || startSnapshot.date,
        endDate: endPoint?.date || endSnapshot.date,
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
  setPanelBusy("#compare", false);
}

async function renderDcaEngine() {
  const token = ++dcaRenderToken;
  if (!marketHistoryReady) {
    $("#dcaBasis").textContent = "歷史日線正在背景載入；完成後會自動計算投入、年化報酬與最大跌幅。";
    ["#dcaInvested", "#dcaValue", "#dcaReturn", "#dcaMdd"].forEach((selector) => {
      $(selector).textContent = "載入中";
      $(selector).className = "";
    });
    drawLineChart($("#dcaChart"), []);
    $("#dcaLegend").innerHTML = "";
    $("#dcaResults").innerHTML = `<div class="rank-row"><strong>正在準備完整歷史日線</strong><small>每日報價已可使用，回測資料完成後會自動顯示。</small></div>`;
    return;
  }
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
    `正在取得 ${state.dcaStartDate} 到 ${state.dcaEndDate} 的每日收盤價，依每期投入金額計算股數與 XIRR 年化...`;
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
    "排定日若非交易日，使用往前最近交易日收盤價買入；總報酬 = 總資產 / 累積投入 - 1，年化 = 每期現金流 XIRR。";
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
            <div>
              <small>總報酬</small>
              <span class="${classForReturn(rate)}">${percent.format(rate)}</span>
            </div>
            <div>
              <small>年化</small>
              <span class="${classForReturn(annualized)}">${metricValue(annualized)}</span>
            </div>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderConditionalEngine() {
  if (!marketHistoryReady) {
    $("#conditionalBasis").textContent = "歷史日線正在背景載入；完成後會自動計算觸發買入策略。";
    ["#conditionalInvested", "#conditionalValue", "#conditionalReturn", "#conditionalTrades"].forEach((selector) => {
      $(selector).textContent = "載入中";
      $(selector).className = "";
    });
    drawLineChart($("#conditionalChart"), []);
    $("#conditionalLegend").innerHTML = "";
    $("#conditionalResults").innerHTML = `<div class="rank-row"><strong>正在準備完整歷史日線</strong><small>完成後會自動執行策略回測。</small></div>`;
    return;
  }
  if (state.conditionalStartDate > state.conditionalEndDate) {
    $("#conditionalBasis").textContent = "日期區間錯誤：開始日期必須早於結束日期。";
    $("#conditionalInvested").textContent = "--";
    $("#conditionalValue").textContent = "--";
    $("#conditionalReturn").textContent = "--";
    $("#conditionalTrades").textContent = "--";
    drawLineChart($("#conditionalChart"), []);
    $("#conditionalLegend").innerHTML = "";
    $("#conditionalResults").innerHTML = `<div class="rank-row"><strong>日期區間錯誤</strong><small>開始日期必須早於結束日期。</small></div>`;
    return;
  }

  const threshold = Math.max(0, Number(state.conditionalThreshold) || 0);
  const rows = state.conditionalSymbols.map(buildConditionalStrategy);
  const validRows = rows.filter((row) => row.last?.invested > 0);
  $("#conditionalBasis").textContent =
    `模擬區間：${state.conditionalStartDate} 到 ${state.conditionalEndDate}。` +
    `規則：每個交易日與前一交易日相比，${conditionalModeText()}達 ${threshold}% 時，以收盤價投入 ${currency.format(state.conditionalAmount)}。` +
    "策略報酬 = 期末資產 / 累積投入 - 1。";

  if (!state.conditionalSymbols.length) {
    $("#conditionalInvested").textContent = "--";
    $("#conditionalValue").textContent = "--";
    $("#conditionalReturn").textContent = "--";
    $("#conditionalTrades").textContent = "--";
    drawLineChart($("#conditionalChart"), []);
    $("#conditionalLegend").innerHTML = "";
    $("#conditionalResults").innerHTML = `<div class="rank-row"><strong>尚未選擇標的</strong><small>先從你的公司股票清單加入一檔股票或 ETF。</small></div>`;
    return;
  }

  if (!validRows.length) {
    $("#conditionalInvested").textContent = "--";
    $("#conditionalValue").textContent = "--";
    $("#conditionalReturn").textContent = "--";
    $("#conditionalTrades").textContent = "0";
    drawLineChart($("#conditionalChart"), []);
    $("#conditionalLegend").innerHTML = "";
    $("#conditionalResults").innerHTML = rows
      .map((row) => `<div class="rank-row"><strong>${row.symbol} ${stocks[row.symbol].name}</strong><small>${row.error || "沒有可用資料"}</small><strong>--</strong></div>`)
      .join("");
    return;
  }

  const totals = validRows.reduce(
    (sum, row) => ({
      invested: sum.invested + row.last.invested,
      value: sum.value + row.last.value,
      triggers: sum.triggers + row.triggers.length,
    }),
    { invested: 0, value: 0, triggers: 0 },
  );
  const totalReturn = totals.invested > 0 ? totals.value / totals.invested - 1 : NaN;
  $("#conditionalInvested").textContent = currency.format(totals.invested);
  $("#conditionalValue").textContent = currency.format(totals.value);
  $("#conditionalReturn").textContent = percent.format(totalReturn);
  $("#conditionalReturn").className = classForReturn(totalReturn);
  $("#conditionalTrades").textContent = `${totals.triggers} 次`;

  const colors = ["#0f766e", "#a15c07", "#2563eb", "#7c3aed", "#b42318", "#475569", "#15803d"];
  const allDates = [
    ...new Set(validRows.flatMap((row) => row.series.map((point) => point.date))),
  ].sort((a, b) => a.localeCompare(b));
  const totalInvestedSeries = allDates.map((date) => ({
    date,
    value: validRows.reduce((sum, row) => sum + (pricePointOnOrBefore(row.series, date)?.invested || 0), 0),
  })).filter((point) => point.value > 0);
  drawLineChart($("#conditionalChart"), [
    { series: totalInvestedSeries, color: "#667068" },
    ...validRows.map((row, index) => ({ series: row.series, color: colors[index % colors.length] })),
  ], currency, { startDate: state.conditionalStartDate, endDate: state.conditionalEndDate });
  $("#conditionalLegend").innerHTML = [
    { label: "合計累積投入", color: "#667068" },
    ...validRows.map((row, index) => ({ label: `${row.symbol} ${stocks[row.symbol].name}`, color: colors[index % colors.length] })),
  ]
    .map((item) => `<span><i style="background:${item.color}"></i>${item.label}</span>`)
    .join("");

  $("#conditionalResults").innerHTML = rows
    .sort((a, b) => {
      if (!Number.isFinite(a.returnRate) && Number.isFinite(b.returnRate)) return 1;
      if (Number.isFinite(a.returnRate) && !Number.isFinite(b.returnRate)) return -1;
      return (b.returnRate || -Infinity) - (a.returnRate || -Infinity);
    })
    .map((row) => {
      if (!row.last?.invested) {
        return `
          <div class="rank-row">
            <div>
              <strong>${row.symbol} ${stocks[row.symbol].name}</strong>
              <small>${row.error || "沒有可用資料"}</small>
            </div>
            <strong>--</strong>
          </div>
        `;
      }
      const visibleTriggers = row.triggers.slice(0, 8);
      return `
        <div class="rank-row dca-result-row conditional-result-row">
          <div class="conditional-result-main">
            <strong>${row.symbol} ${stocks[row.symbol].name}</strong>
            <small>
              單檔投入 ${currency.format(row.last.invested)} · 單檔資產 ${currency.format(row.last.value)} ·
              區間漲跌 ${metricValue(row.periodReturn)} · 觸發 ${row.triggers.length} 次
            </small>
            <small>
              最大單日漲 ${metricValue(row.bestDay?.value)} · 最大單日跌 ${metricValue(row.worstDay?.value)}
            </small>
            <div class="trigger-list" aria-label="${row.symbol} 觸發買入日期">
              ${visibleTriggers
                .map(
                  (trade) => `
                    <div class="trigger-item">
                      <span>${trade.date}</span>
                      <b class="${classForReturn(trade.dailyReturn)}">${percent.format(trade.dailyReturn)}</b>
                      <small>收盤 ${trade.close.toFixed(2)} · 投入 ${currency.format(trade.amount)}</small>
                    </div>
                  `,
                )
                .join("")}
            </div>
            ${row.triggers.length > visibleTriggers.length ? `<small>還有 ${row.triggers.length - visibleTriggers.length} 次未列出</small>` : ""}
          </div>
          <div class="dca-result-metrics">
            <div>
              <small>單檔報酬</small>
              <span class="${classForReturn(row.returnRate)}">${percent.format(row.returnRate)}</span>
            </div>
            <div>
              <small>單檔資產</small>
              <span>${currency.format(row.last.value)}</span>
            </div>
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
          <td><button class="stock-link" type="button" data-select-stock="${row.symbol}">${row.symbol} ${stock.name}</button><br><small>${stock.sector}</small></td>
          <td class="${classForReturn(row.returnRate)}">${percent.format(row.returnRate)}</td>
          <td>${formatVolume(stock.volume)}</td>
          <td>${formatMarketCap(row.marketCap)}</td>
          <td><span class="score-pill">${row.score}</span></td>
        </tr>
      `;
    })
    .join("");
}

function renderHotStocks() {
  $("#hotMarket").value = state.hotMarket;
  const sectors = hotStockSectors();
  $("#hotSector").innerHTML = sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("");
  if (!sectors.includes(state.hotSector)) state.hotSector = "全部";
  $("#hotSector").value = state.hotSector;
  $("#hotMinTradeValue").value = state.hotMinTradeValue;
  $("#hotMinReturn").value = state.hotMinReturn;
  $("#hotLimit").value = state.hotLimit;

  Object.entries(state.hotWeights).forEach(([key, value]) => {
    const input = $(`#hotWeight${key[0].toUpperCase()}${key.slice(1)}`);
    const label = $(`#hotWeight${key[0].toUpperCase()}${key.slice(1)}Label`);
    if (input) input.value = value;
    if (label) label.textContent = `${value}%`;
  });

  const rows = hotStockRows();
  const universeCount = hotStockUniverse().length;
  const dailySeriesCount = rows.filter((row) => row.hasDailySeries).length;
  $("#hotStockNote").textContent =
    `資料日 ${state.marketDataDate || priceDates.at(-1)}，從 ${universeCount} 檔上市櫃標的中篩選。` +
    `成交值與成交筆數使用當日官方行情；短線漲幅與 20 日位置在 ${dailySeriesCount} 檔上榜標的使用日線，其餘使用可用歷史快照降級估算。`;

  if (!rows.length) {
    $("#hotStockRows").innerHTML = `<tr><td colspan="8">目前條件沒有符合的股票，請降低成交值或漲幅門檻。</td></tr>`;
    return;
  }

  $("#hotStockRows").innerHTML = rows
    .map((row, index) => {
      const stock = stocks[row.symbol];
      return `
        <tr>
          <td>${index + 1}</td>
          <td><button class="stock-link" type="button" data-select-stock="${row.symbol}">${row.symbol} ${stock.name}</button><br><small>${stock.market} · ${stock.sector}</small></td>
          <td><span class="score-pill hot-score">${row.score}</span></td>
          <td>${formatMarketCap(row.tradeValue)}</td>
          <td>${Math.round(row.valueScore * 100)}%</td>
          <td class="${classForReturn(row.shortReturn)}">${percent.format(row.shortReturn)}</td>
          <td>${Math.round(row.highPosition * 100)}%</td>
          <td class="${classForReturn(row.relativeStrength)}">${percent.format(row.relativeStrength)}</td>
        </tr>
      `;
    })
    .join("");
}

function scheduleHotStocksRender() {
  if (hotStocksRenderFrame) return;
  hotStocksRenderFrame = window.requestAnimationFrame(() => {
    hotStocksRenderFrame = 0;
    renderHotStocks();
  });
}

async function renderStockTrend() {
  const token = ++stockTrendRenderToken;
  setPanelBusy("#stock-trend", true);
  const symbol = stocks[state.stockTrendSymbol] ? state.stockTrendSymbol : visibleSymbols()[0] || "2330";
  state.stockTrendSymbol = symbol;
  const stock = stocks[symbol];
  $("#stockTrendSymbol").value = symbol;
  $("#stockTrendTitle").textContent = stock.name;
  $("#stockTrendSubtitle").textContent = `${symbol} · ${stock.market} · ${stock.sector}`;
  [...document.querySelectorAll("[data-stock-period]")].forEach((button) => {
    button.classList.toggle("active", button.dataset.stockPeriod === state.stockTrendPeriod);
  });
  $("#stockTrendNote").textContent = `正在取得 ${symbol} ${stock.name} ${stockTrendPeriodLabel()} 走勢...`;
  $("#stockTrendReturn").textContent = "--";
  $("#stockTrendReturn").className = "apple-change";
  $("#stockTrendLatest").textContent = "--";
  $("#stockTrendHigh").textContent = "--";
  $("#stockTrendLow").textContent = "--";
  $("#stockTrendVolume").textContent = "--";
  $("#stockTrendTradeValue").textContent = "--";
  if (!marketHistoryReady) {
    const latest = latestPrice(symbol);
    $("#stockTrendLatest").textContent = latest > 0 ? priceFormat.format(latest) : "--";
    $("#stockTrendReturn").textContent = `${stock.dailyReturn >= 0 ? "+" : ""}${percent.format(stock.dailyReturn)} 今日`;
    $("#stockTrendReturn").className = `apple-change ${classForReturn(stock.dailyReturn)}`.trim();
    $("#stockTrendVolume").textContent = formatVolume(stock.volume);
    $("#stockTrendTradeValue").textContent = formatMarketCap(stock.tradeValue || 0);
    $("#stockTrendNote").textContent = "已顯示當日收盤價，完整日線正在背景載入。";
    drawPriceChart($("#stockTrendChart"), []);
    setPanelBusy("#stock-trend", false);
    return;
  }
  drawPriceChart($("#stockTrendChart"), []);
  stockTrendChartState = { points: [], hoverIndex: null, symbol, baseValue: 0 };

  const result = await stockTrendPoints(symbol);
  if (token !== stockTrendRenderToken) return;

  const points = result.points;
  const first = points[0];
  const last = points.at(-1);
  const values = points.map((point) => point.value);
  const returnRate = first?.value > 0 && last?.value > 0 ? last.value / first.value - 1 : NaN;
  const high = Math.max(...values);
  const low = Math.min(...values);

  $("#stockTrendReturn").textContent = Number.isFinite(returnRate)
    ? `${returnRate >= 0 ? "+" : ""}${percent.format(returnRate)} ${stockTrendPeriodLabel()}`
    : "--";
  $("#stockTrendReturn").className = `apple-change ${classForReturn(returnRate)}`.trim();
  $("#stockTrendLatest").textContent = last?.value > 0 ? priceFormat.format(last.value) : "--";
  $("#stockTrendHigh").textContent = Number.isFinite(high) ? priceFormat.format(high) : "--";
  $("#stockTrendLow").textContent = Number.isFinite(low) ? priceFormat.format(low) : "--";
  $("#stockTrendVolume").textContent = formatVolume(stock.volume);
  $("#stockTrendTradeValue").textContent = formatMarketCap(stock.tradeValue || 0);
  $("#stockTrendNote").textContent =
    points.length >= 2
      ? `${symbol} ${stock.name}，${result.source}，區間 ${first.date} 到 ${last.date}，共 ${points.length} 個交易日；滑鼠移到線上可看單日收盤時間與價格。`
      : `${symbol} ${stock.name} 在 ${stockTrendPeriodLabel()} 區間沒有足夠日線；請改選較短期間或稍後重試。`;
  stockTrendChartState = { points, hoverIndex: null, symbol, baseValue: first?.value || 0 };
  drawPriceChart($("#stockTrendChart"), points, stockTrendChartState);
  setPanelBusy("#stock-trend", false);
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
  const chartStatus = $("#equityChartStatus");
  if (!marketHistoryReady) {
    drawLineChart($("#equityChart"), [], currency, { emptyMessage: "正在載入歷史日線，完成後顯示每日資產曲線。" });
    $("#equityLegend").innerHTML = "";
    if (chartStatus) chartStatus.textContent = "歷史日線載入中，不使用快照直線代替。";
    return;
  }
  const colors = ["#a15c07", "#2563eb", "#7c3aed", "#b42318", "#475569", "#15803d", "#be185d", "#0891b2"];
  const alternatives = visibleSymbols()
    .map((symbol, index) => ({
      symbol,
      label: `${symbol} ${stocks[symbol].name}`,
      color: colors[index % colors.length],
      series: buildScenarioEquitySeries(symbol),
    }))
    .filter((item) => item.series.length);
  const real = { label: "真實投資", color: "#0f766e", series: buildScenarioEquitySeries("real") };
  drawLineChart($("#equityChart"), [
    ...(real.series.length ? [real] : []),
    ...alternatives.map((item) => ({ series: item.series, color: item.color })),
  ]);
  $("#equityLegend").innerHTML = [
    ...(real.series.length ? [{ label: real.label, color: real.color }] : []),
    ...alternatives.map((item) => ({ label: item.label, color: item.color })),
  ]
    .map((item) => `<span><i style="background:${item.color}"></i>${item.label}</span>`)
    .join("");
  if (chartStatus) {
    const dateNote = marketHistoryLatestDate ? `，日線至 ${marketHistoryLatestDate}` : "";
    const refreshNote = marketHistoryRefreshing ? "，正在背景補齊最新資料" : "";
    chartStatus.textContent = `以每日收盤估值，共 ${real.series.length.toLocaleString("zh-TW")} 個交易日${dateNote}${refreshNote}。`;
  }
}

function renderHistoryLoadingPlaceholders() {
  $("#scenarioRanking").innerHTML = `<div class="rank-row"><strong>正在準備替代結果</strong><small>完整日線會在背景載入，不影響目前股價與自選清單操作。</small></div>`;
  $("#singleTradeAnalysis").innerHTML = "<p>完整日線載入後會自動顯示單筆交易分析。</p>";
  $("#portfolioBasis").textContent = "正在背景載入投資組合需要的歷史日線。";
  $("#portfolioResults").innerHTML = "";
  $("#compareBasis").textContent = "正在背景載入比較標的的歷史日線。";
  $("#compareBreakdown").innerHTML = "";
  drawBarChart($("#compareBarChart"), []);
  $("#advancedMetrics").innerHTML = `<div class="analysis-card"><span>分析狀態</span><strong>載入中</strong><small>日線完成後自動計算。</small></div>`;
}

function render() {
  renderMetrics();
  renderTrades();
  renderMarketRanking();
  renderHotStocks();
  renderRiskPosition();
  if (marketHistoryReady) {
    scheduleHistoryDependentRender();
    return;
  }
  renderHistoryLoadingPlaceholders();
  renderDcaEngine();
  renderConditionalEngine();
  renderStockTrend();
  renderCharts();
}

function bindEvents() {
  bindInteractionFeedback();
  bindNavigationState();
  bindStockSearch();
  $("#refreshDataButton")?.addEventListener("click", async (event) => {
    const button = event.currentTarget;
    setButtonBusy(button, true, "更新中");
    setMarketStatus("正在重新確認當日收盤資料...");
    try {
      await new Promise((resolve) => window.requestAnimationFrame(() => window.requestAnimationFrame(resolve)));
      await loadDailyMarketData();
    } finally {
      setButtonBusy(button, false);
      showButtonResult(button, "檢查完成");
    }
  });
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
    const input = $("#stockSearchInput");
    const submitButton = event.currentTarget.querySelector("button[type='submit']");
    const symbol = resolveStockInput(input.value);
    closeStockSearchSuggestions();
    if (!symbol) {
      showControlError(input, "找不到這個股票代號或名稱，請從建議清單選擇。");
      return;
    }
    if (state.watchlist.includes(symbol)) {
      showControlError(input, `${symbol} 已經在自選清單中。`);
      return;
    }
    state.watchlist = uniqueSymbols([...state.watchlist, symbol]);
    input.value = "";
    saveUserState();
    renderControls();
    render();
    showButtonResult(submitButton, "已新增");
    notifyInteraction(`${symbol} ${stocks[symbol].name} 已加入自選清單`, "ok");
    void (async () => {
      setMarketStatus(`正在補齊 ${symbol} ${stocks[symbol]?.name || ""} 的完整歷史日線...`);
      const ready = await ensureFullHistoryFor(symbol);
      setMarketStatus(
        ready
          ? `${symbol} ${stocks[symbol]?.name || ""} 已完成完整歷史日線檢查。`
          : `${symbol} ${stocks[symbol]?.name || ""} 完整歷史日線暫時無法載入，會保留在清單並於下次開啟時再試。`,
        ready ? "ok" : "warn",
      );
      if (state.stockTrendSymbol === symbol) renderStockTrend();
    })();
  });

  $("#watchlistChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-watchlist]");
    if (!button) return;
    if (state.watchlist.length <= 1) {
      notifyInteraction("自選清單至少要保留一檔股票。", "warn");
      return;
    }
    const symbol = button.dataset.removeWatchlist;
    state.watchlist = state.watchlist.filter((item) => item !== symbol);
    syncSelections();
    saveUserState();
    renderControls();
    render();
    notifyInteraction(`${symbol} 已從自選清單移除`, "ok");
  });

  $("#portfolioAddForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const submitButton = event.currentTarget.querySelector("button[type='submit']");
    const symbol = $("#portfolioAddSelect").value;
    if (!symbol) return;
    const remaining = Math.max(0, 100 - portfolioWeights().reduce((sum, item) => sum + item.percent, 0));
    state.portfolioSymbols = uniqueSymbols([...state.portfolioSymbols, symbol]);
    state.portfolioWeights[symbol] = Math.min(remaining, 10);
    saveUserState();
    renderControls();
    renderPortfolioEngine();
    showButtonResult(submitButton, "已加入");
    notifyInteraction(`${symbol} 已加入投資組合`, "ok");
  });

  $("#compareAddForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const submitButton = event.currentTarget.querySelector("button[type='submit']");
    const symbol = $("#compareAddSelect").value;
    if (!symbol) return;
    state.compareSymbols = uniqueSymbols([...state.compareSymbols, symbol]);
    saveUserState();
    renderControls();
    renderCompareEngine();
    showButtonResult(submitButton, "已加入");
    notifyInteraction(`${symbol} 已加入比較`, "ok");
  });

  $("#compareChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-compare]");
    if (!button) return;
    state.compareSymbols = state.compareSymbols.filter((symbol) => symbol !== button.dataset.removeCompare);
    saveUserState();
    renderControls();
    renderCompareEngine();
    notifyInteraction(`${button.dataset.removeCompare} 已從比較移除`, "ok");
  });

  $("#portfolioAmount").addEventListener("input", () => {
    state.portfolioAmount = Number($("#portfolioAmount").value) || 0;
    saveUserState();
    renderPortfolioEngine();
  });
  $("#portfolioWeights").addEventListener("input", (event) => {
    if (event.target.matches("[data-weight]")) enforcePortfolioWeightLimit(event.target);
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
    notifyInteraction(`${button.dataset.removePortfolio} 已從投資組合移除`, "ok");
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
    const submitButton = event.currentTarget.querySelector("button[type='submit']");
    const symbol = $("#dcaAddSelect").value;
    if (!symbol) return;
    state.dcaSymbols = uniqueSymbols([...state.dcaSymbols, symbol]);
    saveUserState();
    renderControls();
    renderDcaEngine();
    showButtonResult(submitButton, "已加入");
    notifyInteraction(`${symbol} 已加入定期定額`, "ok");
  });
  $("#dcaChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-dca]");
    if (!button) return;
    state.dcaSymbols = state.dcaSymbols.filter((symbol) => symbol !== button.dataset.removeDca);
    saveUserState();
    renderControls();
    renderDcaEngine();
    notifyInteraction(`${button.dataset.removeDca} 已從定期定額移除`, "ok");
  });

  $("#conditionalMode").addEventListener("change", (event) => {
    state.conditionalMode = event.target.value;
    saveUserState();
    renderConditionalEngine();
  });
  $("#conditionalThreshold").addEventListener("input", (event) => {
    state.conditionalThreshold = Number(event.target.value) || 0;
    saveUserState();
    renderConditionalEngine();
  });
  $("#conditionalAmount").addEventListener("input", (event) => {
    state.conditionalAmount = Number(event.target.value) || 0;
    saveUserState();
    renderConditionalEngine();
  });
  $("#conditionalStartDate").addEventListener("change", (event) => {
    state.conditionalStartDate = event.target.value || priceDates[0];
    saveUserState();
    renderConditionalEngine();
  });
  $("#conditionalEndDate").addEventListener("change", (event) => {
    state.conditionalEndDate = event.target.value || priceDates.at(-1);
    saveUserState();
    renderConditionalEngine();
  });
  $("#conditionalAddForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const submitButton = event.currentTarget.querySelector("button[type='submit']");
    const symbol = $("#conditionalAddSelect").value;
    if (!symbol) return;
    state.conditionalSymbols = uniqueSymbols([...state.conditionalSymbols, symbol]);
    saveUserState();
    renderControls();
    renderConditionalEngine();
    showButtonResult(submitButton, "已加入");
    notifyInteraction(`${symbol} 已加入條件買入`, "ok");
  });
  $("#conditionalChips").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-conditional]");
    if (!button) return;
    state.conditionalSymbols = state.conditionalSymbols.filter((symbol) => symbol !== button.dataset.removeConditional);
    saveUserState();
    renderControls();
    renderConditionalEngine();
    notifyInteraction(`${button.dataset.removeConditional} 已從條件買入移除`, "ok");
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
  $("#hotMarket").addEventListener("change", (event) => {
    state.hotMarket = event.target.value;
    state.hotSector = "全部";
    saveUserState();
    renderHotStocks();
  });
  $("#hotSector").addEventListener("change", (event) => {
    state.hotSector = event.target.value;
    saveUserState();
    renderHotStocks();
  });
  [
    ["hotMinTradeValue", "hotMinTradeValue"],
    ["hotMinReturn", "hotMinReturn"],
    ["hotLimit", "hotLimit"],
  ].forEach(([selector, key]) => {
    $(`#${selector}`).addEventListener("input", (event) => {
      state[key] = Number(event.target.value) || 0;
      saveUserState();
      scheduleHotStocksRender();
    });
  });
  [
    ["hotWeightValue", "value"],
    ["hotWeightMomentum", "momentum"],
    ["hotWeightHigh", "high"],
    ["hotWeightTrades", "trades"],
    ["hotWeightRelative", "relative"],
  ].forEach(([selector, key]) => {
    $(`#${selector}`).addEventListener("input", (event) => {
      state.hotWeights[key] = Number(event.target.value) || 0;
      saveUserState();
      scheduleHotStocksRender();
    });
  });
  $("#stockTrendSymbol").addEventListener("change", (event) => {
    state.stockTrendSymbol = event.target.value;
    saveUserState();
    renderStockTrend();
  });
  $("#stockTrendPeriods").addEventListener("click", (event) => {
    const button = event.target.closest("[data-stock-period]");
    if (!button) return;
    state.stockTrendPeriod = button.dataset.stockPeriod;
    saveUserState();
    notifyInteraction(`已切換至 ${stockTrendPeriodLabel()} 走勢`, "ok");
    renderStockTrend();
  });
  $("#stockTrendChart").addEventListener("mousemove", (event) => {
    const canvas = event.currentTarget;
    const points = stockTrendChartState.points || [];
    if (points.length < 2) return;
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) * (canvas.width / rect.width);
    const padding = { left: 18, right: 18 };
    const ratio = clamp((x - padding.left) / Math.max(1, canvas.width - padding.left - padding.right));
    const hoverIndex = Math.round(ratio * (points.length - 1));
    if (stockTrendChartState.hoverIndex === hoverIndex || stockTrendHoverFrame) return;
    stockTrendHoverFrame = window.requestAnimationFrame(() => {
      stockTrendHoverFrame = 0;
      stockTrendChartState = { ...stockTrendChartState, hoverIndex };
      drawPriceChart(canvas, points, stockTrendChartState);
    });
  });
  $("#stockTrendChart").addEventListener("mouseleave", (event) => {
    if (stockTrendHoverFrame) window.cancelAnimationFrame(stockTrendHoverFrame);
    stockTrendHoverFrame = 0;
    const points = stockTrendChartState.points || [];
    stockTrendChartState = { ...stockTrendChartState, hoverIndex: null };
    drawPriceChart(event.currentTarget, points, stockTrendChartState);
  });
  document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-select-stock]");
    if (!button) return;
    const symbol = button.dataset.selectStock;
    if (!stocks[symbol]) return;
    state.stockTrendSymbol = symbol;
    saveUserState();
    notifyInteraction(`正在開啟 ${symbol} ${stocks[symbol].name} 走勢`, "info");
    renderStockTrend();
    document.querySelector("#stock-trend")?.scrollIntoView({ behavior: "smooth", block: "start" });
  });
  $("#riskSymbol").addEventListener("change", (event) => {
    state.riskSymbol = event.target.value;
    saveUserState();
    renderRiskPosition();
  });

  $("#tradeForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const submitButton = event.currentTarget.querySelector("button[type='submit']");
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
    showButtonResult(submitButton, "已記錄");
    notifyInteraction(`${symbol} 交易已加入投資紀錄`, "ok");
  });

  $("#tradeTable").addEventListener("click", (event) => {
    const button = event.target.closest("[data-delete]");
    if (!button) return;
    const removedTrade = trades.find((trade) => trade.id === button.dataset.delete);
    trades = trades.filter((trade) => trade.id !== button.dataset.delete);
    saveUserState();
    render();
    notifyInteraction(`${removedTrade?.symbol || "交易"} 已從投資紀錄刪除`, "ok");
  });
}

bindEvents();
if (window.TWSE_MARKET_DATA) {
  loadDailyMarketData();
} else {
  hydrateUserStateOnce();
  renderControls();
  render();
  loadDailyMarketData();
}

setInterval(() => {
  const cache = readCache();
  const today = new Date().toISOString().slice(0, 10);
  if (!cache || cache.cachedAt !== today) {
    loadDailyMarketData();
  }
}, 60 * 60 * 1000);
