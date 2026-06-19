import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const outFile = resolve(root, "data", "market-data.js");

const endpoints = {
  daily: "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data",
  dailyOpenApi: "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
  companies: "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
  tpexDaily: "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=data",
  twseHistory: "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX",
  tpexHistory: "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php",
  finmindData: "https://api.finmindtrade.com/api/v4/data",
  nasdaqScreener: "https://api.nasdaq.com/api/screener/stocks?tableonly=true&limit=5000&download=true",
};

let usSymbols = {
  AAPL: "Apple",
  MSFT: "Microsoft",
  NVDA: "NVIDIA",
  AMZN: "Amazon",
  GOOGL: "Alphabet A",
  GOOG: "Alphabet C",
  META: "Meta",
  AVGO: "Broadcom",
  TSLA: "Tesla",
  BRK_B: "Berkshire Hathaway B",
  JPM: "JPMorgan Chase",
  WMT: "Walmart",
  LLY: "Eli Lilly",
  V: "Visa",
  ORCL: "Oracle",
  MA: "Mastercard",
  XOM: "Exxon Mobil",
  UNH: "UnitedHealth",
  JNJ: "Johnson & Johnson",
  HD: "Home Depot",
  PG: "Procter & Gamble",
  BAC: "Bank of America",
  ABBV: "AbbVie",
  KO: "Coca-Cola",
  PLTR: "Palantir",
  PM: "Philip Morris",
  CVX: "Chevron",
  GE: "GE Aerospace",
  CSCO: "Cisco",
  IBM: "IBM",
  WFC: "Wells Fargo",
  CRM: "Salesforce",
  ABT: "Abbott",
  MCD: "McDonald's",
  LIN: "Linde",
  MRK: "Merck",
  DIS: "Disney",
  NOW: "ServiceNow",
  T: "AT&T",
  ACN: "Accenture",
  INTU: "Intuit",
  UBER: "Uber",
  GS: "Goldman Sachs",
  RTX: "RTX",
  ISRG: "Intuitive Surgical",
  TXN: "Texas Instruments",
  VZ: "Verizon",
  AXP: "American Express",
  AMD: "AMD",
  MS: "Morgan Stanley",
  PEP: "PepsiCo",
  NFLX: "Netflix",
  COST: "Costco",
  SPY: "SPDR S&P 500 ETF",
  QQQ: "Invesco QQQ ETF",
  VOO: "Vanguard S&P 500 ETF",
  VTI: "Vanguard Total Stock Market ETF",
  IVV: "iShares Core S&P 500 ETF",
  SCHD: "Schwab US Dividend Equity ETF",
  VGT: "Vanguard Information Technology ETF",
  XLK: "Technology Select Sector SPDR ETF",
  SMH: "VanEck Semiconductor ETF",
  SOXX: "iShares Semiconductor ETF",
  DIA: "SPDR Dow Jones ETF",
  IWM: "iShares Russell 2000 ETF",
  TLT: "iShares 20+ Year Treasury Bond ETF",
  BND: "Vanguard Total Bond Market ETF",
  AGG: "iShares Core US Aggregate Bond ETF",
  IBIT: "iShares Bitcoin Trust",
};

const requiredUsEtfs = {
  SPY: "SPDR S&P 500 ETF",
  QQQ: "Invesco QQQ ETF",
  VOO: "Vanguard S&P 500 ETF",
  VTI: "Vanguard Total Stock Market ETF",
  IVV: "iShares Core S&P 500 ETF",
  SCHD: "Schwab US Dividend Equity ETF",
  VGT: "Vanguard Information Technology ETF",
  XLK: "Technology Select Sector SPDR ETF",
  SMH: "VanEck Semiconductor ETF",
  SOXX: "iShares Semiconductor ETF",
  DIA: "SPDR Dow Jones ETF",
  IWM: "iShares Russell 2000 ETF",
  TLT: "iShares 20+ Year Treasury Bond ETF",
  BND: "Vanguard Total Bond Market ETF",
  AGG: "iShares Core US Aggregate Bond ETF",
  IBIT: "iShares Bitcoin Trust",
};

const requiredHistorySymbols = ["2330", "0050", "2454"];
const taiwanDailySeriesSymbols = [
  "0050",
  "006208",
  "00631L",
  "00735",
  "00981A",
  "2303",
  "2308",
  "2317",
  "2330",
  "2454",
  "2881",
];
let dcaSeriesSymbols = [...new Set([...taiwanDailySeriesSymbols, ...Object.keys(usSymbols)])];
const taiwanDailySeriesSet = new Set(taiwanDailySeriesSymbols);

async function getJson(url, options = {}) {
  const text = await getText(url, options);
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON from ${url}: ${text.slice(0, 80).replace(/\s+/g, " ")}`);
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getText(url, options = {}) {
  let lastError;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 20_000);
    try {
      const response = await fetch(url, { ...options, cache: "no-store", signal: controller.signal });
      if (!response.ok) throw new Error(`Failed to fetch ${url}: ${response.status}`);
      return response.text();
    } catch (error) {
      lastError = error;
      if (attempt < 3) await sleep(750 * attempt);
    } finally {
      clearTimeout(timeout);
    }
  }
  throw lastError;
}

async function readExistingPayload() {
  try {
    const text = await readFile(outFile, "utf8");
    const match = text.match(/^window\.TWSE_MARKET_DATA = (.*);\s*$/s);
    return match ? JSON.parse(match[1]) : null;
  } catch {
    return null;
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

async function getTwseDailyRows() {
  try {
    const rows = normalizeTwseDailyRows(parseCsv(await getText(endpoints.daily)));
    if (rows.length > 100) return rows;
  } catch (error) {
    console.warn(`TWSE CSV daily unavailable: ${error.message}`);
  }
  return normalizeTwseDailyRows(await getJson(endpoints.dailyOpenApi));
}

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

async function getTwseHistory(isoDate) {
  const url = `${endpoints.twseHistory}?date=${isoToTwseDate(isoDate)}&type=ALLBUT0999&response=json`;
  return parseTwseHistory(await getJson(url));
}

async function getTpexHistory(isoDate) {
  const params = new URLSearchParams({
    l: "zh-tw",
    d: isoToTpexRocDate(isoDate),
    se: "EW",
    o: "json",
  });
  return parseTpexHistory(await getJson(`${endpoints.tpexHistory}?${params}`));
}

async function getHistorySnapshot(period, targetDate) {
  for (let offset = 0; offset <= 14; offset += 1) {
    const date = addDaysIso(targetDate, -offset);
    const [twseResult, tpexResult] = await Promise.allSettled([getTwseHistory(date), getTpexHistory(date)]);
    const twse = twseResult.status === "fulfilled" ? twseResult.value : {};
    const tpex = tpexResult.status === "fulfilled" ? tpexResult.value : {};
    const closes = { ...twse, ...tpex };
    const twseCount = Object.keys(twse).length;
    const tpexCount = Object.keys(tpex).length;
    const count = Object.keys(closes).length;
    if (twseCount > 100 && count > 100) {
      return {
        period,
        targetDate,
        date,
        closes,
        twseCount,
        tpexCount,
      };
    }
  }
  console.warn(`No historical snapshot found for ${period} near ${targetDate}`);
  return { period, targetDate, date: "", closes: {}, twseCount: 0, tpexCount: 0 };
}

async function getHistorySnapshots(latestDate, existingHistory = {}) {
  const lookbacks = [
    ["1w", 7],
    ["1m", 30],
    ["1y", 365],
  ];
  const periods = {};
  for (const [period, days] of lookbacks) {
    periods[period] = await getHistorySnapshot(period, addDaysIso(latestDate, -days));
  }
  const dcaSeries = await getDcaSymbolHistories(latestDate, existingHistory.dcaSeries || {});
  fillPeriodSnapshotsFromSeries(periods, dcaSeries);
  assertPeriodSnapshots(periods);
  return { latestDate, periods, dcaSeries };
}

function addMonthsIso(isoDate, months) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  date.setUTCMonth(date.getUTCMonth() + months);
  return date.toISOString().slice(0, 10);
}

function assertHistoryPayload(history) {
  const periods = ["1w", "1m", "1y"];
  const missingPeriodSymbols = missingPeriodSnapshotSymbols(history.periods || {}, periods);
  const missingDcaSeries = dcaSeriesSymbols.filter((symbol) => (history.dcaSeries?.[symbol]?.points || []).length < 2);
  if (missingPeriodSymbols.length || missingDcaSeries.length) {
    throw new Error(
      [
        missingPeriodSymbols.length ? `missing period symbols ${missingPeriodSymbols.slice(0, 12).join(", ")}` : "",
        missingDcaSeries.length ? `missing DCA series ${missingDcaSeries.slice(0, 12).join(", ")}` : "",
      ]
        .filter(Boolean)
        .join("; "),
    );
  }
}

function missingPeriodSnapshotSymbols(periods, periodNames = Object.keys(periods)) {
  return periodNames.flatMap((period) =>
    requiredHistorySymbols
      .filter((symbol) => !periods?.[period]?.closes?.[symbol])
      .map((symbol) => `${period}:${symbol}`),
  );
}

function assertPeriodSnapshots(periods) {
  const missing = missingPeriodSnapshotSymbols(periods, ["1w", "1m", "1y"]);
  if (missing.length) throw new Error(`missing period symbols ${missing.slice(0, 12).join(", ")}`);
}

function dateFromTimestamp(timestamp, timeZone) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date(timestamp * 1000));
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function taiwanDateFromTimestamp(timestamp) {
  return dateFromTimestamp(timestamp, "Asia/Taipei");
}

function marketDateFromTimestamp(symbol, timestamp) {
  return usSymbols[symbol] ? dateFromTimestamp(timestamp, "America/New_York") : taiwanDateFromTimestamp(timestamp);
}

function pricePointOnOrBefore(points, date) {
  for (let index = points.length - 1; index >= 0; index -= 1) {
    if (points[index].date <= date) return points[index];
  }
  return null;
}

function fillPeriodSnapshotsFromSeries(periods, dcaSeries) {
  for (const snapshot of Object.values(periods)) {
    if (!snapshot?.closes) continue;
    const date = snapshot.date || snapshot.targetDate;
    for (const symbol of requiredHistorySymbols) {
      if (snapshot.closes[symbol]) continue;
      const point = pricePointOnOrBefore(dcaSeries?.[symbol]?.points || [], date);
      if (point?.value > 0) snapshot.closes[symbol] = point.value;
    }
  }
}

function yahooSymbolFor(symbol) {
  if (usSymbols[symbol]) return symbol.replace("_", "-");
  return `${symbol}.TW`;
}

async function getYahooHistory(symbol, startDate, endDate) {
  const start = Math.max(0, Math.floor(new Date(`${addDaysIso(startDate, -14)}T00:00:00Z`).getTime() / 1000));
  const end = Math.floor(new Date(`${addDaysIso(endDate, 3)}T00:00:00Z`).getTime() / 1000);
  const yahooSymbol = yahooSymbolFor(symbol);
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooSymbol)}?period1=${start}&period2=${end}&interval=1d&events=history%7Cdiv%7Csplit`;
  const result = await getJson(url);
  const chart = result?.chart?.result?.[0];
  const timestamps = chart?.timestamp || [];
  const adjusted = chart?.indicators?.adjclose?.[0]?.adjclose || [];
  const closes = chart?.indicators?.quote?.[0]?.close || [];
  return timestamps
    .map((timestamp, index) => ({
      date: marketDateFromTimestamp(symbol, timestamp),
      value: toNumber(closes[index]) || toNumber(adjusted[index]),
    }))
    .filter((point) => point.date >= addDaysIso(startDate, -14) && point.date <= addDaysIso(endDate, 3) && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
}

async function getFinMindTaiwanHistory(symbol, startDate, endDate) {
  const params = new URLSearchParams({
    dataset: "TaiwanStockPrice",
    data_id: symbol,
    start_date: startDate,
    end_date: endDate,
  });
  const result = await getJson(`${endpoints.finmindData}?${params}`);
  if (result?.status !== 200 || !Array.isArray(result.data)) {
    throw new Error(result?.msg || `FinMind returned status ${result?.status || "unknown"}`);
  }
  return result.data
    .map((row) => ({
      date: row.date,
      value: toNumber(row.close),
      volume: toNumber(row.Trading_Volume),
      source: "finmind",
    }))
    .filter((point) => point.date >= startDate && point.date <= endDate && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
}

async function getCloseHistory(symbol, startDate, endDate) {
  if (taiwanDailySeriesSet.has(symbol)) {
    try {
      const points = await getFinMindTaiwanHistory(symbol, startDate, endDate);
      if (points.length && points[0].date <= addDaysIso(startDate, 31)) {
        return { points, source: "FinMind TaiwanStockPrice" };
      }
      console.warn(
        `FinMind history for ${symbol} starts at ${points[0]?.date || "none"}, expected near ${startDate}; falling back to Yahoo.`,
      );
    } catch (error) {
      console.warn(`FinMind unavailable for ${symbol}: ${error.message}`);
    }
  }

  const points = await getYahooHistory(symbol, startDate, endDate);
  return { points, source: "Yahoo Finance chart" };
}

function historyStartDateFor(symbol, latestDate) {
  if (taiwanDailySeriesSet.has(symbol)) return "2000-01-01";
  return addMonthsIso(latestDate.slice(0, 7) + "-01", -36);
}

function usableExistingSeries(existingSeries, symbol) {
  const entry = existingSeries?.[symbol];
  return (entry?.points || []).length >= 120 ? entry : null;
}

async function getDcaSymbolHistories(latestDate, existingSeries = {}) {
  const histories = {};
  for (const symbol of dcaSeriesSymbols) {
    const startDate = historyStartDateFor(symbol, latestDate);
    try {
      const { points, source } = await getCloseHistory(symbol, startDate, latestDate);
      if (points.length >= 2) {
        histories[symbol] = { symbol, startDate, endDate: latestDate, source, points };
      } else {
        const existing = usableExistingSeries(existingSeries, symbol);
        if (existing) {
          histories[symbol] = existing;
          console.warn(`Short close series for ${symbol}; using existing ${existing.points.length}-point series.`);
        }
      }
    } catch (error) {
      const existing = usableExistingSeries(existingSeries, symbol);
      if (existing) {
        histories[symbol] = existing;
        console.warn(`No close series for ${symbol}, using existing ${existing.points.length}-point series: ${error.message}`);
      } else {
        console.warn(`No close series for ${symbol}: ${error.message}`);
      }
    }
    await sleep(160);
  }
  return histories;
}

async function getUsQuote(symbol) {
  try {
    const yahooSymbol = yahooSymbolFor(symbol);
    const result = await getJson(`https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooSymbol)}?range=10d&interval=1d`);
    const chart = result?.chart?.result?.[0];
    const timestamps = chart?.timestamp || [];
    const quote = chart?.indicators?.quote?.[0] || {};
    const rows = timestamps
      .map((timestamp, index) => ({
        date: marketDateFromTimestamp(symbol, timestamp),
        open: toNumber(quote.open?.[index]),
        high: toNumber(quote.high?.[index]),
        low: toNumber(quote.low?.[index]),
        close: toNumber(quote.close?.[index]),
        volume: toNumber(quote.volume?.[index]),
      }))
      .filter((row) => row.close > 0);
    const row = rows.at(-1);
    if (!row) return null;
    return {
      symbol,
      name: usSymbols[symbol],
      date: row.date,
      open: row.open,
      high: row.high,
      low: row.low,
      close: row.close,
      volume: row.volume,
    };
  } catch (error) {
    console.warn(`Skipping ${symbol}: ${error.message}`);
    return null;
  }
}

async function getCompanyRows(existingPayload) {
  try {
    const rows = await getJson(endpoints.companies);
    if (Array.isArray(rows) && rows.length > 100) return rows;
    throw new Error(`company endpoint returned ${Array.isArray(rows) ? rows.length : "non-array"} rows`);
  } catch (error) {
    if (existingPayload?.companies?.length > 100) {
      console.warn(`Company endpoint unavailable, using existing company snapshot: ${error.message}`);
      return existingPayload.companies;
    }
    throw error;
  }
}

function normalizeUsSymbol(symbol) {
  return String(symbol || "").trim().toUpperCase().replace(/[.-]/g, "_");
}

function cleanUsCompanyName(name) {
  return String(name || "")
    .replace(/\s+(Class [A-Z] )?(Common|Capital) Stock.*$/i, "")
    .replace(/\s+Ordinary Shares.*$/i, "")
    .trim();
}

function isUsCompanySecurity(row) {
  const name = String(row?.name || "");
  const marketCap = toNumber(row?.marketCap);
  if (!(marketCap > 0) || !row?.symbol || !row?.sector) return false;
  return !/\b(ETF|ETN|Fund|Warrant|Right|Unit|Preferred Stock)\b/i.test(name);
}

async function getTopUsSymbols(existingPayload) {
  const mandatory = {
    MU: "Micron Technology",
    SNDK: "Sandisk",
  };
  try {
    const result = await getJson(endpoints.nasdaqScreener, {
      headers: {
        Accept: "application/json, text/plain, */*",
        Origin: "https://www.nasdaq.com",
        Referer: "https://www.nasdaq.com/",
        "User-Agent": "Mozilla/5.0",
      },
    });
    const ranked = (result?.data?.rows || [])
      .filter(isUsCompanySecurity)
      .map((row) => ({
        symbol: normalizeUsSymbol(row.symbol),
        name: cleanUsCompanyName(row.name),
        marketCap: toNumber(row.marketCap),
      }))
      .filter((row) => row.symbol && row.name)
      .sort((a, b) => b.marketCap - a.marketCap);
    const unique = [...new Map(ranked.map((row) => [row.symbol, row])).values()];
    const selected = unique.slice(0, 200);
    for (const [symbol, name] of Object.entries(mandatory)) {
      if (selected.some((row) => row.symbol === symbol)) continue;
      const rankedRow = unique.find((row) => row.symbol === symbol) || { symbol, name, marketCap: 0 };
      selected.pop();
      selected.push(rankedRow);
    }
    if (selected.length !== 200) throw new Error(`Nasdaq universe returned ${selected.length} usable symbols`);
    return { ...Object.fromEntries(selected.map((row) => [row.symbol, row.name])), ...requiredUsEtfs };
  } catch (error) {
    const existingCompanies = Object.fromEntries(
      (existingPayload?.usDaily || [])
        .filter((row) => row?.symbol && !requiredUsEtfs[normalizeUsSymbol(row.symbol)])
        .slice(0, 200)
        .map((row) => [normalizeUsSymbol(row.symbol), row.name || normalizeUsSymbol(row.symbol)]),
    );
    const seedCompanies = Object.fromEntries(Object.entries(usSymbols).filter(([symbol]) => !requiredUsEtfs[symbol]));
    const fallback = Object.keys(existingCompanies).length >= 200 ? existingCompanies : seedCompanies;
    const fallbackEntries = Object.entries({ ...fallback, ...mandatory });
    while (fallbackEntries.length > 200) {
      const removableIndex = fallbackEntries.findLastIndex(([symbol]) => !mandatory[symbol]);
      if (removableIndex === -1) break;
      fallbackEntries.splice(removableIndex, 1);
    }
    console.warn(`Nasdaq top-200 universe unavailable, using existing universe: ${error.message}`);
    return { ...Object.fromEntries(fallbackEntries), ...requiredUsEtfs };
  }
}

function assertUsRows(rows) {
  const symbols = new Set(rows.map((row) => row.symbol));
  const expectedCount = 200 + Object.keys(requiredUsEtfs).length;
  const missing = ["MU", "SNDK", ...Object.keys(requiredUsEtfs)].filter((symbol) => !symbols.has(symbol));
  if (rows.length !== expectedCount || symbols.size !== expectedCount || missing.length) {
    throw new Error(`invalid US universe: rows=${rows.length}, unique=${symbols.size}, missing=${missing.join(",") || "none"}`);
  }
}

async function getUsRows(existingPayload) {
  const existingBySymbol = Object.fromEntries((existingPayload?.usDaily || []).map((row) => [row.symbol, row]));
  const rows = [];
  for (const symbol of Object.keys(usSymbols)) {
    const row = await getUsQuote(symbol);
    if (row) {
      rows.push(row);
    } else if (existingBySymbol[symbol]) {
      rows.push(existingBySymbol[symbol]);
      console.warn(`Using existing US quote for ${symbol}.`);
    }
    await sleep(160);
  }
  return rows;
}

const existingPayload = await readExistingPayload();
usSymbols = await getTopUsSymbols(existingPayload);
dcaSeriesSymbols = [...new Set([...taiwanDailySeriesSymbols, ...Object.keys(usSymbols)])];
const [daily, companies, tpexCsv, usDaily] = await Promise.all([
  getTwseDailyRows(),
  getCompanyRows(existingPayload),
  getText(endpoints.tpexDaily),
  getUsRows(existingPayload),
]);
assertUsRows(usDaily);

const tpexDaily = parseCsv(tpexCsv);
const latestDate = rocDateToIso(daily?.[0]?.Date || "");
const history = latestDate ? await getHistorySnapshots(latestDate, existingPayload?.history || {}) : { latestDate: "", periods: {} };
if (latestDate) assertHistoryPayload(history);
const payload = {
  daily,
  companies,
  tpexDaily,
  usDaily,
  history,
  fetchedAt: new Date().toISOString(),
  date: daily?.[0]?.Date || "",
  source: endpoints,
};

await mkdir(dirname(outFile), { recursive: true });
await writeFile(
  outFile,
  `window.TWSE_MARKET_DATA = ${JSON.stringify(payload)};\n`,
  "utf8",
);

console.log(
  `Wrote ${outFile} with ${daily.length} TWSE rows, ${tpexDaily.length} TPEx rows, ${companies.length} company rows, ${payload.usDaily.length} US rows, ${Object.keys(history.periods).length} history snapshots, and ${Object.keys(history.dcaSeries || {}).length} DCA daily series.`,
);
