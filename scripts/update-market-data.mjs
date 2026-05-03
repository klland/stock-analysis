import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const outFile = resolve(root, "data", "market-data.js");

const endpoints = {
  daily: "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
  companies: "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
  tpexDaily: "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=data",
};

const usSymbols = {
  AAPL: "Apple",
  MSFT: "Microsoft",
  NVDA: "NVIDIA",
  GOOGL: "Alphabet A",
  GOOG: "Alphabet C",
  AMZN: "Amazon",
  META: "Meta",
  TSLA: "Tesla",
  AVGO: "Broadcom",
  AMD: "AMD",
  NFLX: "Netflix",
  COST: "Costco",
  JPM: "JPMorgan Chase",
  BRK_B: "Berkshire Hathaway B",
  SPY: "SPDR S&P 500 ETF",
  QQQ: "Invesco QQQ ETF",
  VOO: "Vanguard S&P 500 ETF",
  VTI: "Vanguard Total Stock Market ETF",
  DIA: "SPDR Dow Jones ETF",
  IWM: "iShares Russell 2000 ETF",
  TLT: "iShares 20+ Year Treasury Bond ETF",
  IBIT: "iShares Bitcoin Trust",
};

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
const payload = {
  daily,
  companies,
  tpexDaily,
  usDaily,
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
  `Wrote ${outFile} with ${daily.length} TWSE rows, ${tpexDaily.length} TPEx rows, ${companies.length} company rows, and ${payload.usDaily.length} US rows.`,
);
