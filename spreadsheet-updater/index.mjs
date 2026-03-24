/**
 * スプレッドシート自動更新スクリプト
 *
 * 対象データ:
 * - PER推移（日次）: 日経平均株価, EPS, BPS, 売買代金（全市場, プライム, スタンダード, グロース）
 * - 空売り比率（日次）: 空売り比率合計, 出来高
 * - 信用倍率（週次 火・木）: 売り残, 買い残, 信用評価率
 * - 投資主体別（週次 火・木）: 海外, 個人, 投資信託, 信託銀行, 証券自己, 事業法人, 個人（現金）, 個人（信用）
 *
 * 使い方:
 *   node index.mjs daily    - PER推移・空売り比率を更新
 *   node index.mjs weekly   - 信用倍率・投資主体別を更新
 *   node index.mjs all      - 全て更新
 */

import { chromium } from 'playwright';
import { google } from 'googleapis';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const CREDENTIALS_PATH = join(__dirname, '..', 'investment-claude-e261a94c1e98.json');
const SPREADSHEET_ID = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

// ===== Google Sheets認証 =====
function getSheets() {
  const credentials = JSON.parse(readFileSync(CREDENTIALS_PATH, 'utf8'));
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// ===== ユーティリティ =====

// "2026-03-10" or "2026/03/10" → Excelシリアル値
function dateToSerial(dateStr) {
  const normalized = dateStr.replace(/-/g, '/');
  const parts = normalized.split('/');
  const date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const epoch = new Date(1899, 11, 30);
  return Math.round((date - epoch) / (24 * 60 * 60 * 1000));
}

// カンマ付き数値をパース（+/-記号対応）
function parseNumber(str) {
  if (!str || str === '-' || str === '' || str === '－') return null;
  return Number(str.replace(/,/g, '').replace(/[+▲▼]/g, '').trim());
}

// 符号付きでパース（+は正、▲/▼/-は負）
function parseSignedNumber(str) {
  if (!str || str === '-' || str === '' || str === '－') return null;
  const cleaned = str.replace(/,/g, '').trim();
  if (cleaned.startsWith('-') || cleaned.startsWith('▼')) {
    return -Math.abs(Number(cleaned.replace(/[-▲▼+]/g, '')));
  }
  return Number(cleaned.replace(/[+▲▼]/g, ''));
}

// ===== スクレイピング =====

async function scrapePER(page) {
  console.log('PER推移データを取得中...');
  await page.goto('https://nikkei225jp.com/data/per.php', { waitUntil: 'networkidle' });
  await page.waitForSelector('#datatbl tr td', { timeout: 10000 });

  const data = await page.evaluate(() => {
    const table = document.getElementById('datatbl');
    const rows = table.querySelectorAll('tr');
    // 行1がヘッダー、行2が最新データ
    const cells = rows[1].querySelectorAll('td');
    return {
      date: cells[0]?.textContent?.trim(),       // 日付 "2026-03-10"
      nikkei: cells[1]?.textContent?.trim(),      // 日経平均株価
      volume: cells[3]?.textContent?.trim(),      // 出来高
      per: cells[4]?.textContent?.trim(),         // PER
      pbr: cells[5]?.textContent?.trim(),         // PBR
      eps: cells[6]?.textContent?.trim(),         // EPS
      bps: cells[7]?.textContent?.trim(),         // BPS
    };
  });

  if (!data) throw new Error('PERデータの取得に失敗しました');
  console.log(`  日付: ${data.date}, 日経平均: ${data.nikkei}, EPS: ${data.eps}, BPS: ${data.bps}`);
  return data;
}

async function scrapeKarauri(page) {
  console.log('空売り比率データを取得中...');
  await page.goto('https://nikkei225jp.com/data/karauri.php', { waitUntil: 'networkidle' });
  await page.waitForSelector('#datatbl tr td', { timeout: 10000 });

  const data = await page.evaluate(() => {
    const table = document.getElementById('datatbl');
    const rows = table.querySelectorAll('tr');
    // 行0がヘッダー、行1が最新データ
    const cells = rows[1].querySelectorAll('td');
    return {
      date: cells[0]?.textContent?.trim(),       // 日付 "2026-03-11"
      nikkei: cells[1]?.textContent?.trim(),      // 日経平均株価
      volume: cells[3]?.textContent?.trim(),      // プライム出来高（百万株）
      karauriTotal: cells[4]?.textContent?.trim(),// 空売り比率合計
    };
  });

  if (!data) throw new Error('空売り比率データの取得に失敗しました');
  console.log(`  日付: ${data.date}, 日経平均: ${data.nikkei}, 出来高: ${data.volume}, 空売り比率: ${data.karauriTotal}`);
  return data;
}

async function scrapeBaibaiDaikin(page) {
  console.log('売買代金データを取得中...');
  await page.goto('https://nikkei225jp.com/chart/', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForSelector('span.txD2', { timeout: 15000 });

  const data = await page.evaluate(() => {
    const spans = document.querySelectorAll('span.txD2');
    const results = [];
    for (const span of spans) {
      if (span.textContent.includes('売買代金')) {
        let el = span;
        for (let i = 0; i < 10; i++) {
          el = el.parentElement;
          if (!el) break;
          const text = el.textContent.trim();
          const match = text.match(/(プライム|スタンダード|グロース).*売買代金[：:]([0-9,]+)/);
          if (match) {
            results.push({ market: match[1], value: match[2] });
            break;
          }
        }
      }
    }
    return results;
  });

  const prime = data.find(d => d.market === 'プライム')?.value;
  const standard = data.find(d => d.market === 'スタンダード')?.value;
  const growth = data.find(d => d.market === 'グロース')?.value;

  if (!prime) throw new Error('売買代金データの取得に失敗しました');
  console.log(`  プライム: ${prime}, スタンダード: ${standard}, グロース: ${growth}`);
  return { prime, standard, growth };
}

async function scrapeSinyou(page) {
  console.log('信用倍率データを取得中...');
  await page.goto('https://nikkei225jp.com/data/sinyou.php', { waitUntil: 'networkidle' });
  await page.waitForSelector('table');

  const data = await page.evaluate(() => {
    const tables = document.querySelectorAll('table');
    for (const t of tables) {
      const header = t.querySelector('tr');
      if (header && header.textContent.includes('売り残') && header.textContent.includes('買い残')) {
        const rows = t.querySelectorAll('tr');
        const cells = rows[1].querySelectorAll('td');
        return {
          date: cells[0]?.textContent?.trim(),         // 日付
          sellAmount: cells[2]?.textContent?.trim(),   // 売り残金額（百万円）
          buyAmount: cells[5]?.textContent?.trim(),    // 買い残金額（百万円）
          evalRate: cells[8]?.textContent?.trim(),      // 信用評価率
        };
      }
    }
    return null;
  });

  if (!data) throw new Error('信用倍率データの取得に失敗しました');
  console.log(`  日付: ${data.date}, 売り残: ${data.sellAmount}, 買い残: ${data.buyAmount}, 信用評価率: ${data.evalRate}`);
  return data;
}

async function scrapeSaitei(page) {
  console.log('裁定買い残データを取得中...');
  await page.goto('https://nikkei225jp.com/data/saitei.php', { waitUntil: 'networkidle' });
  await page.waitForSelector('#datatbl tr td', { timeout: 10000 });

  const data = await page.evaluate(() => {
    const table = document.getElementById('datatbl');
    const rows = table.querySelectorAll('tr');
    const cells = rows[1].querySelectorAll('td');
    return {
      date: cells[0]?.textContent?.trim(),         // 日付
      buyZan: cells[4]?.textContent?.trim(),        // 買い残（百万円）
      sellZan: cells[5]?.textContent?.trim(),       // 売り残（百万円）
      sabiki: cells[6]?.textContent?.trim(),        // 差引（百万円）
      sabikiZenhi: cells[7]?.textContent?.trim(),   // 差引前比（百万円）
    };
  });

  if (!data) throw new Error('裁定買い残データの取得に失敗しました');
  console.log(`  日付: ${data.date}, 買残: ${data.buyZan}, 売残: ${data.sellZan}, 差引: ${data.sabiki}, 前比: ${data.sabikiZenhi}`);
  return data;
}

async function scrapeShutai(page) {
  console.log('投資主体別データを取得中...');
  await page.goto('https://nikkei225jp.com/data/shutai.php', { waitUntil: 'networkidle' });
  await page.waitForSelector('#datatbl');

  const data = await page.evaluate(() => {
    const table = document.getElementById('datatbl');
    const rows = table.querySelectorAll('tr');
    // 行0: メインヘッダー（日付, 日経平均, 変化, 海外, 証券自己, 個人, 法人, 法人金融）
    // 行1: サブヘッダー（個人計, 現金, 信用, 投資信託, 事業法人, その他, 信託銀行, 生保損保, 都銀地銀）
    // 行2: 最新データ
    const cells = rows[2].querySelectorAll('td');
    return {
      date: cells[0]?.textContent?.trim(),             // 日付
      kaigai: cells[3]?.textContent?.trim(),            // 海外
      shokenJiko: cells[4]?.textContent?.trim(),        // 証券自己
      kojinKei: cells[5]?.textContent?.trim(),          // 個人計
      kojinGenkin: cells[6]?.textContent?.trim(),       // 個人（現金）
      kojinShinyo: cells[7]?.textContent?.trim(),       // 個人（信用）
      toshiShintaku: cells[8]?.textContent?.trim(),     // 投資信託
      jigyoHojin: cells[9]?.textContent?.trim(),        // 事業法人
      shintakuGinko: cells[11]?.textContent?.trim(),    // 信託銀行
    };
  });

  if (!data) throw new Error('投資主体別データの取得に失敗しました');
  console.log(`  日付: ${data.date}, 海外: ${data.kaigai}, 個人: ${data.kojinKei}, 個人(現金): ${data.kojinGenkin}, 個人(信用): ${data.kojinShinyo}, 事業法人: ${data.jigyoHojin}`);
  return data;
}

// ===== Google Sheets書き込み =====

async function getExistingDates(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!B3:B12`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  return res.data.values?.map(row => row[0])?.filter(Boolean) || [];
}

async function insertRowAtTop(sheets, sheetId) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        insertDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
          inheritFromBefore: false,
        },
      }],
    },
  });
}

async function writePER(sheets, data, baibaiData) {
  console.log('PER推移シートに書き込み中...');
  const sheetId = 1978398551;

  // 重複チェック
  const existingDates = await getExistingDates(sheets, 'PER推移');
  const dateFormatted = data.date.replace(/-/g, '/');
  if (existingDates.some(d => d === dateFormatted)) {
    console.log(`  ${dateFormatted} は既に入力済みです。スキップします。`);
    return false;
  }

  await insertRowAtTop(sheets, sheetId);

  const nikkei = parseNumber(data.nikkei);
  const eps = parseNumber(data.eps);
  const bps = parseNumber(data.bps);

  const prime = baibaiData ? parseNumber(baibaiData.prime) : null;
  const standard = baibaiData ? parseNumber(baibaiData.standard) : null;
  const growth = baibaiData ? parseNumber(baibaiData.growth) : null;
  const total = (prime || 0) + (standard || 0) + (growth || 0);

  const values = [[
    '',                  // A
    dateToSerial(data.date), // B: 日付
    nikkei,              // C: 日経平均
    '=(C3-C4)/C4',       // D: 増減率
    eps,                 // E: EPS
    '=(E3-E4)/E4',       // F: EPS増減率
    '=C3/E3',            // G: PER
    '=$E3*H$2',          // H: PER 12.0
    '=$E3*I$2',          // I: PER 13.0
    '=$E3*J$2',          // J: PER 14.0
    '=$E3*$K$2',         // K: PER 14.5
    '=$E3*L$2',          // L: PER 15.0
    '=$E3*M$2',          // M: PER 15.5
    '=$E3*N$2',          // N: PER 16.0
    '=$E3*O$2',          // O: PER 16.5
    '=$E3*P$2',          // P: PER 17.0
    '=$E3*Q$2',          // Q: PER 17.5
    '=$E3*R$2',          // R: PER 18.0
    '=$E3*S$2',          // S: PER 19.0
    '=$E3*T$2',          // T: PER 20.0
    '=$E3*U$2',          // U: PER 21.0
    bps,                 // V: BPS
    '=C3/V3',            // W: PBR
    '',                  // X: （空き）
    total || '',         // Y: 全市場売買代金（億円）
    prime || '',         // Z: プライム売買代金（億円）
    standard || '',      // AA: スタンダード売買代金（億円）
    growth || '',        // AB: グロース売買代金（億円）
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: "'PER推移'!A3:AB3",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });

  console.log(`  ${dateFormatted} のデータを書き込みました。`);
  return true;
}

async function writeSinyou(sheets, data, nikkeiPrice, saiteiData) {
  console.log('信用倍率シートに書き込み中...');
  const sheetId = 798317965;

  const existingDates = await getExistingDates(sheets, '信用倍率');
  const dateFormatted = data.date.replace(/-/g, '/');

  if (existingDates.some(d => d === dateFormatted)) {
    // 既存行の信用評価率・裁定データが空の場合は更新する
    const updated = await updateMissingSinyouFields(sheets, dateFormatted, data.evalRate, saiteiData);
    if (!updated) {
      console.log(`  ${dateFormatted} は既に入力済みです。スキップします。`);
    }
    return updated;
  }

  await insertRowAtTop(sheets, sheetId);

  const values = [[
    '',                          // A
    dateToSerial(data.date),     // B: 日付
    parseNumber(data.sellAmount),// C: 売り残
    parseNumber(data.buyAmount), // D: 買い残
    '=D3/C3',                   // E: 信用倍率
    parseSignedNumber(data.evalRate), // F: 信用評価率
    nikkeiPrice || '',           // G: 日経株価
    '',                          // H: （空き）
    saiteiData ? parseNumber(saiteiData.buyZan) : '',   // I: 裁定買残
    saiteiData ? parseNumber(saiteiData.sellZan) : '',   // J: 裁定売残
    saiteiData ? parseNumber(saiteiData.sabiki) : '',    // K: 差引
    saiteiData ? parseSignedNumber(saiteiData.sabikiZenhi) : '', // L: 差引全比
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: "'信用倍率'!A3:L3",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });

  console.log(`  ${dateFormatted} のデータを書き込みました。`);
  return true;
}

// 既存行の信用評価率・裁定データが空の場合に更新する
async function updateMissingSinyouFields(sheets, dateFormatted, evalRateStr, saiteiData) {
  // B列～L列を取得して、該当日付の行を探す
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'信用倍率'!B3:L52",
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = res.data.values || [];
  for (let i = 0; i < rows.length; i++) {
    const rowDate = rows[i][0]; // B列: 日付
    if (rowDate !== dateFormatted) continue;

    const rowNum = i + 3;
    const batchData = [];

    // F列: 信用評価率が空なら更新
    const rowEvalRate = rows[i][4];
    const evalRate = parseSignedNumber(evalRateStr);
    if ((!rowEvalRate || rowEvalRate === '') && evalRate !== null) {
      batchData.push({
        range: `'信用倍率'!F${rowNum}`,
        values: [[evalRate]],
      });
    }

    // I~L列: 裁定データが空なら更新
    const rowSaiteiI = rows[i][7]; // I列（B列からの相対位置7）
    if ((!rowSaiteiI || rowSaiteiI === '') && saiteiData) {
      batchData.push({
        range: `'信用倍率'!I${rowNum}:L${rowNum}`,
        values: [[
          parseNumber(saiteiData.buyZan),
          parseNumber(saiteiData.sellZan),
          parseNumber(saiteiData.sabiki),
          parseSignedNumber(saiteiData.sabikiZenhi),
        ]],
      });
    }

    if (batchData.length === 0) return false;

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
    });

    const fields = batchData.map(d => d.range.match(/!([A-Z])/)[1]).join(', ');
    console.log(`  ${dateFormatted} の空フィールド(${fields}列)を更新しました。`);
    return true;
  }
  return false;
}

async function writeKarauri(sheets, data) {
  console.log('空売り比率シートに書き込み中...');
  const sheetId = 1671609006;

  const existingDates = await getExistingDates(sheets, '空売り比率');
  const dateFormatted = data.date.replace(/-/g, '/');
  if (existingDates.some(d => d === dateFormatted)) {
    console.log(`  ${dateFormatted} は既に入力済みです。スキップします。`);
    return false;
  }

  await insertRowAtTop(sheets, sheetId);

  const values = [[
    '',                          // A
    dateToSerial(data.date),     // B: 日付
    parseNumber(data.nikkei),    // C: 日経株価
    parseNumber(data.volume),    // D: プライム出来高（百万株）
    parseNumber(data.karauriTotal), // E: 空売り比率合計
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: "'空売り比率'!A3:E3",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });

  console.log(`  ${dateFormatted} のデータを書き込みました。`);
  return true;
}

async function writeShutai(sheets, data) {
  console.log('投資主体別シートに書き込み中...');
  const sheetId = 2061132662;

  const existingDates = await getExistingDates(sheets, '投資主体別');
  const dateFormatted = data.date.replace(/-/g, '/');
  if (existingDates.some(d => d === dateFormatted)) {
    console.log(`  ${dateFormatted} は既に入力済みです。スキップします。`);
    return false;
  }

  await insertRowAtTop(sheets, sheetId);

  const values = [[
    '',                                    // A
    dateToSerial(data.date),               // B: 日付
    parseSignedNumber(data.kaigai),        // C: 海外
    parseSignedNumber(data.kojinKei),      // D: 個人
    parseSignedNumber(data.toshiShintaku), // E: 投資信託
    parseSignedNumber(data.shintakuGinko), // F: 信託銀行
    parseSignedNumber(data.shokenJiko),    // G: 証券自己
    parseSignedNumber(data.jigyoHojin),    // H: 事業法人
    '',                                    // I: （空き）
    parseSignedNumber(data.kojinGenkin),   // J: 個人（現金）
    parseSignedNumber(data.kojinShinyo),   // K: 個人（信用）
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: "'投資主体別'!A3:K3",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });

  console.log(`  ${dateFormatted} のデータを書き込みました。`);
  return true;
}

// ===== メイン処理 =====

async function main() {
  const mode = process.argv[2] || 'all';
  console.log(`\n=== スプレッドシート更新 (${mode}) ===`);
  console.log(`実行日時: ${new Date().toLocaleString('ja-JP')}\n`);

  const sheets = getSheets();
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  try {
    let nikkeiPrice = null;

    if (mode === 'daily' || mode === 'all') {
      const perData = await scrapePER(page);
      const baibaiData = await scrapeBaibaiDaikin(page);
      await writePER(sheets, perData, baibaiData);
      nikkeiPrice = parseNumber(perData.nikkei);

      const karauriData = await scrapeKarauri(page);
      await writeKarauri(sheets, karauriData);
    }

    if (mode === 'weekly' || mode === 'all') {
      // 日経株価がまだ取得されていない場合、PERページから取得
      if (!nikkeiPrice) {
        const perData = await scrapePER(page);
        nikkeiPrice = parseNumber(perData.nikkei);
      }

      const sinyouData = await scrapeSinyou(page);
      const saiteiData = await scrapeSaitei(page);
      await writeSinyou(sheets, sinyouData, nikkeiPrice, saiteiData);

      const shutaiData = await scrapeShutai(page);
      await writeShutai(sheets, shutaiData);
    }

    console.log('\n更新完了！');
  } catch (error) {
    console.error('\nエラーが発生しました:', error.message);
    console.error(error.stack);
    process.exit(1);
  } finally {
    await browser.close();
  }
}

main();
