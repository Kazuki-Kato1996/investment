#!/usr/bin/env node
/**
 * 朝の市場レポート生成・メール送信スクリプト
 *
 * 取得する情報:
 * - 日本株市場: 日経平均、TOPIX、マザーズ
 * - 米国株市場: S&P500、ダウ、NASDAQ
 * - 日経先物
 * - 主要ニュース・注目銘柄
 *
 * 使い方:
 *   node morning-market-report.mjs               # メール送信
 *   node morning-market-report.mjs --dry-run      # コンソール出力のみ
 */

import { createTransport } from "nodemailer";
import * as cheerio from "cheerio";
import { readFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const CONFIG_PATH = join(__dirname, "market-report-config.json");

// --- 設定読み込み ---
function loadConfig() {
  if (!existsSync(CONFIG_PATH)) {
    console.error(`設定ファイルが見つかりません: ${CONFIG_PATH}`);
    console.error("market-report-config.json を作成してください。");
    process.exit(1);
  }
  return JSON.parse(readFileSync(CONFIG_PATH, "utf-8"));
}

// --- HTTP取得ヘルパー ---
async function fetchHTML(url) {
  const res = await fetch(url, {
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
      "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
  return await res.text();
}

// --- Yahoo!ファイナンスからbodyテキストで価格・変動を抽出 ---
function parseYahooJPQuote(html) {
  const bodyText = cheerio.load(html)("body").text();
  const priceMatch = bodyText.match(/([0-9,]+\.?[0-9]*)前日比/);
  const changeMatch = bodyText.match(/前日比([+\-][0-9,]+\.?[0-9]*\([+\-][0-9.]+%\))/);
  return {
    price: priceMatch ? priceMatch[1] : "取得失敗",
    change: changeMatch ? changeMatch[1] : "",
  };
}

// --- 日本株指数 (Yahoo!ファイナンス) ---
async function fetchJapanIndices() {
  const indices = [];
  const targets = [
    { url: "https://finance.yahoo.co.jp/quote/998407.O", name: "日経平均" },
    { url: "https://finance.yahoo.co.jp/quote/998405.T", name: "TOPIX" },
    { url: "https://finance.yahoo.co.jp/quote/2516.T", name: "東証グロース250 ETF" },
  ];

  for (const target of targets) {
    try {
      const html = await fetchHTML(target.url);
      const parsed = parseYahooJPQuote(html);
      indices.push({ name: target.name, ...parsed });
    } catch (e) {
      indices.push({ name: target.name, price: "取得失敗", change: e.message });
    }
  }
  return indices;
}

// --- 米国株指数 (Yahoo Finance US) ---
async function fetchUSIndices() {
  const indices = [];
  const targets = [
    { symbol: "^DJI", name: "ダウ平均" },
    { symbol: "^GSPC", name: "S&P 500" },
    { symbol: "^IXIC", name: "NASDAQ" },
  ];

  for (const target of targets) {
    try {
      const url = `https://finance.yahoo.com/quote/${encodeURIComponent(target.symbol)}/`;
      const html = await fetchHTML(url);
      const $ = cheerio.load(html);

      const priceText =
        $(`[data-symbol="${target.symbol}"][data-field="regularMarketPrice"]`)
          .first()
          .text()
          .trim() ||
        $('fin-streamer[data-field="regularMarketPrice"]')
          .first()
          .text()
          .trim();
      const changeText =
        $(`[data-symbol="${target.symbol}"][data-field="regularMarketChange"]`)
          .first()
          .text()
          .trim() ||
        $('fin-streamer[data-field="regularMarketChange"]')
          .first()
          .text()
          .trim();
      const changePctText =
        $(
          `[data-symbol="${target.symbol}"][data-field="regularMarketChangePercent"]`
        )
          .first()
          .text()
          .trim() ||
        $('fin-streamer[data-field="regularMarketChangePercent"]')
          .first()
          .text()
          .trim();

      indices.push({
        name: target.name,
        price: priceText || "取得失敗",
        change: changeText ? `${changeText} (${changePctText})` : "",
      });
    } catch (e) {
      indices.push({ name: target.name, price: "取得失敗", change: e.message });
    }
  }
  return indices;
}

// --- 日経平均先物（nikkei225jp.comからCME・大証のデータを取得） ---
async function fetchNikkeiFutures() {
  const futures = [];
  try {
    const html = await fetchHTML("https://nikkei225jp.com/cme/");
    const $ = cheerio.load(html);

    // CME円建て先物テーブル（直近限月のみ）
    $("table").each((i, tbl) => {
      const text = $(tbl).text();
      if (text.includes("CME￥") && text.includes("限月") && futures.length === 0) {
        $(tbl).find("tr").each((_, tr) => {
          const cells = $(tr).find("td");
          if (cells.length >= 4 && futures.length === 0) {
            const label = cells.eq(0).text().trim();
            if (label.includes("CME￥")) {
              const priceStr = cells.eq(1).text().trim();
              const changeStr = cells.eq(2).text().trim();
              const priceNum = Number(priceStr.replace(/,/g, ""));
              const changeNum = Number(changeStr.replace(/,/g, ""));
              const prevClose = priceNum - changeNum;
              const pct = prevClose ? ((changeNum / prevClose) * 100).toFixed(2) : "0.00";
              futures.push({
                name: "日経先物 CME円建",
                price: priceStr,
                change: `${changeStr} (${changeNum >= 0 ? "+" : ""}${pct}%)`,
              });
            }
          }
        });
      }
    });

    // 大証ミニ（直近限月のみ）
    $("table").each((i, tbl) => {
      const text = $(tbl).text();
      if (text.includes("大証ミニ") && text.includes("限月") && text.includes("出来高") && futures.length < 2) {
        $(tbl).find("tr").each((_, tr) => {
          const cells = $(tr).find("td");
          if (cells.length >= 4 && futures.length < 2) {
            const label = cells.eq(0).text().trim();
            if (label.includes("大証ミニ")) {
              const priceStr = cells.eq(1).text().trim();
              const changeStr = cells.eq(2).text().trim();
              const priceNum = Number(priceStr.replace(/,/g, ""));
              const changeNum = Number(changeStr.replace(/,/g, ""));
              const prevClose = priceNum - changeNum;
              const pct = prevClose ? ((changeNum / prevClose) * 100).toFixed(2) : "0.00";
              futures.push({
                name: "日経先物 大証",
                price: priceStr,
                change: `${changeStr} (${changeNum >= 0 ? "+" : ""}${pct}%)`,
              });
            }
          }
        });
      }
    });

    if (futures.length > 0) return futures;
    return [{ name: "日経平均先物", price: "取得失敗", change: "" }];
  } catch (e) {
    return [{ name: "日経平均先物", price: "取得失敗", change: e.message }];
  }
}

// --- 為替 ---
async function fetchForex() {
  try {
    const url = "https://finance.yahoo.co.jp/quote/USDJPY=FX";
    const html = await fetchHTML(url);
    const bodyText = cheerio.load(html)("body").text();
    // FXはBid/Ask形式
    const bidMatch = bodyText.match(/Bid[^\d]*([0-9]+\.[0-9]+)/);
    const changeMatch = bodyText.match(/Change[^\d+\-]*([+\-]?[0-9]+\.[0-9]+)/);
    const bid = bidMatch ? parseFloat(bidMatch[1]) : 0;
    const chg = changeMatch ? parseFloat(changeMatch[1]) : 0;
    const prevClose = bid - chg;
    const pct = prevClose ? ((chg / prevClose) * 100).toFixed(2) : "0.00";
    return {
      name: "USD/JPY",
      price: bidMatch ? bidMatch[1] : "取得失敗",
      change: changeMatch ? `${chg >= 0 ? "+" : ""}${changeMatch[1]} (${chg >= 0 ? "+" : ""}${pct}%)` : "",
    };
  } catch (e) {
    return { name: "USD/JPY", price: "取得失敗", change: "" };
  }
}

// --- 米国セクターパフォーマンス (Yahoo Finance US) ---
// セクターETFシンボルとセクター名のマッピング
const SECTOR_ETFS = [
  { symbol: "XLK", name: "Technology" },
  { symbol: "XLF", name: "Financial Services" },
  { symbol: "XLY", name: "Consumer Cyclical" },
  { symbol: "XLC", name: "Communication Services" },
  { symbol: "XLV", name: "Healthcare" },
  { symbol: "XLI", name: "Industrials" },
  { symbol: "XLP", name: "Consumer Defensive" },
  { symbol: "XLE", name: "Energy" },
  { symbol: "XLB", name: "Basic Materials" },
  { symbol: "XLRE", name: "Real Estate" },
  { symbol: "XLU", name: "Utilities" },
];

async function fetchUSSectors() {
  const sectors = [];

  // 1) YTDリターンをsectorsページから取得
  const ytdMap = {};
  try {
    const html = await fetchHTML("https://finance.yahoo.com/sectors/");
    const $ = cheerio.load(html);
    $("table").first().find("tr").each((_, tr) => {
      const cells = $(tr).find("td");
      if (cells.length >= 3) {
        const name = cells.eq(0).text().trim();
        const ytdReturn = cells.eq(2).text().trim();
        if (name && name !== "All Sectors") {
          ytdMap[name] = ytdReturn;
        }
      }
    });
  } catch (e) {
    // YTD取得失敗は致命的ではない
  }

  // 2) 各セクターETFの当日変動率を取得
  const etfPromises = SECTOR_ETFS.map(async (etf) => {
    try {
      const html = await fetchHTML(`https://finance.yahoo.com/quote/${etf.symbol}/`);
      const $ = cheerio.load(html);
      const body = $("body").text();
      const prevClose = parseFloat($('fin-streamer[data-field="regularMarketPreviousClose"]').first().text().trim());
      // bodyテキストからETFシンボル直後の終値を取得
      const priceMatch = body.match(new RegExp(etf.symbol + "\\s+.*?\\s+([0-9]+\\.?[0-9]*)\\s"));
      const price = priceMatch ? parseFloat(priceMatch[1]) : 0;
      let dailyChange = "";
      if (prevClose && price) {
        const chg = price - prevClose;
        const pct = ((chg / prevClose) * 100).toFixed(2);
        dailyChange = `${chg >= 0 ? "+" : ""}${chg.toFixed(2)} (${chg >= 0 ? "+" : ""}${pct}%)`;
      }
      return {
        name: etf.name,
        dailyChange,
        ytdReturn: ytdMap[etf.name] || "",
      };
    } catch (e) {
      return { name: etf.name, dailyChange: "", ytdReturn: ytdMap[etf.name] || "" };
    }
  });

  const results = await Promise.all(etfPromises);
  return results;
}

// --- ニュースヘッドライン ---
async function fetchMarketNews() {
  const news = [];

  // Yahoo!ファイナンス マーケットニュース
  try {
    const html = await fetchHTML("https://finance.yahoo.co.jp/news/market");
    const $ = cheerio.load(html);

    $("a").each((_, el) => {
      const href = $(el).attr("href") || "";
      const title = $(el).text().trim();
      if (
        title.length > 10 &&
        title.length < 120 &&
        (href.includes("/news/") || href.includes("article")) &&
        !title.includes("ログイン") &&
        !title.includes("ポートフォリオ") &&
        !news.some((n) => n.title === title) &&
        news.filter((n) => n.source === "Yahoo!ファイナンス").length < 10
      ) {
        const fullUrl = href.startsWith("http")
          ? href
          : `https://finance.yahoo.co.jp${href}`;
        news.push({ title, url: fullUrl, source: "Yahoo!ファイナンス" });
      }
    });
  } catch (e) {
    news.push({
      title: `Yahoo!ファイナンスニュース取得失敗: ${e.message}`,
      url: "",
      source: "error",
    });
  }

  // 日経新聞（無料部分のヘッドライン）
  try {
    const html = await fetchHTML("https://www.nikkei.com/markets/");
    const $ = cheerio.load(html);

    $("a[href*='/article/']")
      .slice(0, 10)
      .each((_, el) => {
        const title = $(el).text().trim();
        const href = $(el).attr("href");
        if (title && title.length > 5 && !news.some((n) => n.title === title)) {
          const fullUrl = href?.startsWith("http")
            ? href
            : `https://www.nikkei.com${href}`;
          news.push({ title, url: fullUrl, source: "日経新聞" });
        }
      });
  } catch (e) {
    news.push({
      title: `日経新聞ニュース取得失敗: ${e.message}`,
      url: "",
      source: "error",
    });
  }

  return news;
}


// --- HTMLメール本文の生成 ---
function buildEmailHTML(data, reportTitle = "朝の市場レポート") {
  const today = new Date().toLocaleDateString("ja-JP", {
    year: "numeric",
    month: "long",
    day: "numeric",
    weekday: "long",
  });

  const indexRow = (idx) => `
    <tr>
      <td style="padding:6px 12px;border-bottom:1px solid #eee;font-weight:bold;">${idx.name}</td>
      <td style="padding:6px 12px;border-bottom:1px solid #eee;text-align:right;">${idx.price}</td>
      <td style="padding:6px 12px;border-bottom:1px solid #eee;text-align:right;color:${idx.change?.includes("-") ? "#d32f2f" : "#2e7d32"}">${idx.change || ""}</td>
    </tr>`;

  const newsItem = (n) =>
    `<li style="margin-bottom:6px;">
      <a href="${n.url}" style="color:#1a73e8;text-decoration:none;">${n.title}</a>
      <span style="color:#999;font-size:12px;"> [${n.source}]</span>
    </li>`;

  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:'Hiragino Sans','Yu Gothic',sans-serif;max-width:680px;margin:0 auto;padding:20px;color:#333;">

<h1 style="color:#1a237e;border-bottom:3px solid #1a237e;padding-bottom:10px;font-size:22px;">
  ${reportTitle}
</h1>
<p style="color:#666;margin-top:-5px;">${today}</p>

<!-- 日本株 -->
<h2 style="color:#1565c0;font-size:18px;margin-top:25px;">日本株市場</h2>
<table style="width:100%;border-collapse:collapse;font-size:14px;">
  <tr style="background:#f5f5f5;">
    <th style="padding:8px 12px;text-align:left;">指数</th>
    <th style="padding:8px 12px;text-align:right;">終値</th>
    <th style="padding:8px 12px;text-align:right;">変動</th>
  </tr>
  ${data.japanIndices.map(indexRow).join("")}
</table>

<!-- 米国株 -->
<h2 style="color:#c62828;font-size:18px;margin-top:25px;">米国株市場（前日終値）</h2>
<table style="width:100%;border-collapse:collapse;font-size:14px;">
  <tr style="background:#f5f5f5;">
    <th style="padding:8px 12px;text-align:left;">指数</th>
    <th style="padding:8px 12px;text-align:right;">終値</th>
    <th style="padding:8px 12px;text-align:right;">変動</th>
  </tr>
  ${data.usIndices.map(indexRow).join("")}
</table>

<!-- 先物・為替 -->
<h2 style="color:#4a148c;font-size:18px;margin-top:25px;">先物・為替</h2>
<table style="width:100%;border-collapse:collapse;font-size:14px;">
  <tr style="background:#f5f5f5;">
    <th style="padding:8px 12px;text-align:left;">指標</th>
    <th style="padding:8px 12px;text-align:right;">価格</th>
    <th style="padding:8px 12px;text-align:right;">変動</th>
  </tr>
  ${data.nikkeiFutures.map(indexRow).join("")}
  ${indexRow({ ...data.forex })}
</table>

<!-- 米国セクターパフォーマンス -->
${
  data.usSectors.length > 0
    ? `
<h2 style="color:#0d47a1;font-size:18px;margin-top:25px;">米国セクターパフォーマンス</h2>
<table style="width:100%;border-collapse:collapse;font-size:13px;">
  <tr style="background:#e3f2fd;"><th style="padding:6px 8px;text-align:left;">セクター</th><th style="padding:6px 8px;text-align:right;">当日変動</th><th style="padding:6px 8px;text-align:right;">YTD</th></tr>
  ${data.usSectors.map((s) => `
    <tr>
      <td style="padding:4px 8px;border-bottom:1px solid #f0f0f0;">${s.name}</td>
      <td style="padding:4px 8px;border-bottom:1px solid #f0f0f0;text-align:right;color:${s.dailyChange?.includes("-") ? "#d32f2f" : "#2e7d32"}">${s.dailyChange}</td>
      <td style="padding:4px 8px;border-bottom:1px solid #f0f0f0;text-align:right;color:${s.ytdReturn?.includes("-") ? "#d32f2f" : "#2e7d32"}">${s.ytdReturn}</td>
    </tr>`).join("")}
</table>`
    : ""
}

<!-- ニュース -->
<h2 style="color:#e65100;font-size:18px;margin-top:25px;">マーケットニュース</h2>
<ul style="padding-left:20px;font-size:14px;line-height:1.8;">
  ${data.news.map(newsItem).join("")}
</ul>

<hr style="margin-top:30px;border:none;border-top:1px solid #ddd;">
<p style="color:#999;font-size:11px;">
  このメールは自動生成されたマーケットレポートです。<br>
  情報元: Yahoo!ファイナンス, 日経新聞(無料部分)<br>
  ※ 投資判断は自己責任でお願いします。
</p>
</body>
</html>`;
}

// --- テキスト版（コンソール/dry-run用） ---
function buildPlainText(data, reportTitle = "朝の市場レポート") {
  let text = "";
  text += "========================================\n";
  text += `  ${reportTitle}\n`;
  text += `  ${new Date().toLocaleDateString("ja-JP")}\n`;
  text += "========================================\n\n";

  text += "【日本株市場】\n";
  for (const idx of data.japanIndices) {
    text += `  ${idx.name.padEnd(20)} ${idx.price}  ${idx.change}\n`;
  }

  text += "\n【米国株市場】\n";
  for (const idx of data.usIndices) {
    text += `  ${idx.name.padEnd(20)} ${idx.price}  ${idx.change}\n`;
  }

  text += "\n【先物・為替】\n";
  for (const f of data.nikkeiFutures) {
    text += `  ${f.name.padEnd(20)} ${f.price}  ${f.change}\n`;
  }
  text += `  ${data.forex.name.padEnd(20)} ${data.forex.price}  ${data.forex.change}\n`;

  if (data.usSectors.length > 0) {
    text += "\n【米国セクターパフォーマンス】\n";
    text += `  ${"セクター".padEnd(25)} ${"当日変動".padStart(16)}  ${"YTD".padStart(8)}\n`;
    for (const s of data.usSectors) {
      text += `  ${s.name.padEnd(25)} ${(s.dailyChange || "-").padStart(16)}  ${(s.ytdReturn || "-").padStart(8)}\n`;
    }
  }

  text += "\n【マーケットニュース】\n";
  for (const n of data.news) {
    text += `  - ${n.title} [${n.source}]\n    ${n.url}\n`;
  }

  return text;
}

// --- メール送信 ---
async function sendEmail(config, html, plainText, reportTitle = "朝の市場レポート") {
  const transporter = createTransport({
    service: "gmail",
    auth: {
      user: config.gmail_user,
      pass: config.gmail_app_password,
    },
  });

  const info = await transporter.sendMail({
    from: `"Market Report" <${config.gmail_user}>`,
    to: config.to_email,
    subject: `${reportTitle} - ${new Date().toLocaleDateString("ja-JP")}`,
    text: plainText,
    html: html,
  });

  console.log(`メール送信完了: ${info.messageId}`);
}

// --- メイン ---
async function main() {
  const isDryRun = process.argv.includes("--dry-run");
  const titleIdx = process.argv.indexOf("--title");
  const reportTitle = titleIdx !== -1 && process.argv[titleIdx + 1] ? process.argv[titleIdx + 1] : "朝の市場レポート";

  console.log("市場データを取得中...");

  // 並列でデータ取得
  const [japanIndices, usIndices, nikkeiFutures, forex, news, usSectors] =
    await Promise.all([
      fetchJapanIndices(),
      fetchUSIndices(),
      fetchNikkeiFutures(),
      fetchForex(),
      fetchMarketNews(),
      fetchUSSectors(),
    ]);

  const data = { japanIndices, usIndices, nikkeiFutures, forex, news, usSectors };

  const html = buildEmailHTML(data, reportTitle);
  const plainText = buildPlainText(data, reportTitle);

  if (isDryRun) {
    console.log(plainText);
    console.log("\n--- (dry-run: メール送信はスキップ) ---");
    return;
  }

  const config = loadConfig();
  await sendEmail(config, html, plainText, reportTitle);
}

main().catch((e) => {
  console.error("エラー:", e.message);
  process.exit(1);
});
