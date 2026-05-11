import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const outFile = resolve(root, "data", "market-data.js");

const endpoints = {
  daily: "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
  companies: "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
  tpexDaily: "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=data",
  twseHistory: "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX",
  tpexHistory: "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php",
};

const usSymbols = {
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

const requiredHistorySymbols = ["2330", "0050", "2454"];
const dcaSeriesSymbols = [
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
  "AAPL",
  "MSFT",
  "NVDA",
  "SPY",
  "QQQ",
];

async function getJson(url) {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status}`);
  }
  return response.json();
}

async function getText(url) {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status}`);
  }
  return response.text();
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

async function getHistorySnapshots(latestDate) {
  const lookbacks = [
    ["1w", 7],
    ["1m", 30],
    ["1y", 365],
  ];
  const periods = {};
  for (const [period, days] of lookbacks) {
    periods[period] = await getHistorySnapshot(period, addDaysIso(latestDate, -days));
  }
  assertPeriodSnapshots(periods);
  const dcaSeries = await getDcaSymbolHistories(latestDate);
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
  const missingDcaSeries = dcaSeriesSymbols.filter((symbol) => (history.dcaSeries?.[symbol]?.points || []).length < 120);
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
  if (usSymbols[symbol]) return symbol.replace("_", "-");
  return `${symbol}.TW`;
}

async function getYahooHistory(symbol, startDate, endDate) {
  const start = Math.floor(new Date(`${addDaysIso(startDate, -14)}T00:00:00Z`).getTime() / 1000);
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
      date: taiwanDateFromTimestamp(timestamp),
      value: toNumber(adjusted[index]) || toNumber(closes[index]),
    }))
    .filter((point) => point.date >= addDaysIso(startDate, -14) && point.date <= addDaysIso(endDate, 3) && point.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
}

async function getDcaSymbolHistories(latestDate) {
  const startDate = addMonthsIso(latestDate.slice(0, 7) + "-01", -36);
  const histories = {};
  for (const symbol of dcaSeriesSymbols) {
    try {
      const points = await getYahooHistory(symbol, startDate, latestDate);
      if (points.length) histories[symbol] = { symbol, startDate, endDate: latestDate, points };
    } catch (error) {
      console.warn(`No DCA series for ${symbol}: ${error.message}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 160));
  }
  return histories;
}

async function getUsQuote(symbol) {
  try {
    const stooqSymbol = symbol.replace("_", "-").toLowerCase();
    const text = await getText(`https://stooq.com/q/l/?s=${stooqSymbol}.us&f=sd2t2ohlcv&h&e=csv`);
    const [row] = parseCsv(text);
    if (!row || row.Close === "N/D") return null;
    return {
      symbol,
      name: usSymbols[symbol],
      date: row.Date,
      open: row.Open,
      high: row.High,
      low: row.Low,
      close: row.Close,
      volume: row.Volume,
    };
  } catch (error) {
    console.warn(`Skipping ${symbol}: ${error.message}`);
    return null;
  }
}

const [daily, companies, tpexCsv, usDaily] = await Promise.all([
  getJson(endpoints.daily),
  getJson(endpoints.companies),
  getText(endpoints.tpexDaily),
  (async () => {
    const rows = [];
    for (const symbol of Object.keys(usSymbols)) {
      const row = await getUsQuote(symbol);
      if (row) rows.push(row);
      await new Promise((resolve) => setTimeout(resolve, 160));
    }
    return rows;
  })(),
]);

const tpexDaily = parseCsv(tpexCsv);
const latestDate = rocDateToIso(daily?.[0]?.Date || "");
const history = latestDate ? await getHistorySnapshots(latestDate) : { latestDate: "", periods: {} };
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
