/**
 * 投資主体別シートに個人（現金）・個人（信用）を追加するスクリプト
 *
 * 1. J2, K2 にヘッダーを書き込み
 * 2. ウェブから全データを取得し、既存行の日付とマッチングしてJ,K列をバックフィル
 *
 * 使い方:
 *   node backfill-kojin.mjs
 */

import { chromium } from 'playwright';
import { google } from 'googleapis';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const CREDENTIALS_PATH = join(__dirname, '..', 'investment-claude-e261a94c1e98.json');
const SPREADSHEET_ID = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

const credentials = JSON.parse(readFileSync(CREDENTIALS_PATH, 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// 符号付きでパース
function parseSignedNumber(str) {
  if (!str || str === '-' || str === '' || str === '－') return null;
  const cleaned = str.replace(/,/g, '').trim();
  if (cleaned.startsWith('-') || cleaned.startsWith('▼')) {
    return -Math.abs(Number(cleaned.replace(/[-▲▼+]/g, '')));
  }
  return Number(cleaned.replace(/[+▲▼]/g, ''));
}

// --- Step 1: ヘッダー追加 ---
console.log('J2, K2 にヘッダーを書き込み中...');
await sheets.spreadsheets.values.update({
  spreadsheetId: SPREADSHEET_ID,
  range: "'投資主体別'!J2:K2",
  valueInputOption: 'USER_ENTERED',
  requestBody: { values: [['個人（現金）', '個人（信用）']] },
});
console.log('  ヘッダー書き込み完了');

// --- Step 2: 既存データの日付を取得 ---
const res = await sheets.spreadsheets.values.get({
  spreadsheetId: SPREADSHEET_ID,
  range: "'投資主体別'!B3:K100",
  valueRenderOption: 'FORMATTED_VALUE',
});
const sheetRows = res.data.values || [];
console.log(`\nスプレッドシート: ${sheetRows.length}行`);

// --- Step 3: ウェブから全データを取得 ---
console.log('ウェブからデータを取得中...');
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage();
await page.goto('https://nikkei225jp.com/data/shutai.php', { waitUntil: 'networkidle' });
await page.waitForSelector('#datatbl');

const webData = await page.evaluate(() => {
  const table = document.getElementById('datatbl');
  const rows = table.querySelectorAll('tr');
  const result = [];
  for (let i = 2; i < rows.length; i++) {
    const cells = rows[i].querySelectorAll('td');
    if (cells.length >= 8) {
      result.push({
        date: cells[0]?.textContent?.trim(),
        kojinGenkin: cells[6]?.textContent?.trim(),  // 個人（現金）
        kojinShinyo: cells[7]?.textContent?.trim(),  // 個人（信用）
      });
    }
  }
  return result;
});
await browser.close();

console.log(`ウェブデータ: ${webData.length}行`);

// --- Step 4: 日付をキーにしたマップを作成 ---
const webMap = {};
for (const d of webData) {
  webMap[d.date] = d;
}

// --- Step 5: マッチングして更新対象を特定 ---
const updates = [];
for (let i = 0; i < sheetRows.length; i++) {
  const sheetDate = sheetRows[i][0]; // B列の日付
  // J列 = B列からの相対位置8（B,C,D,E,F,G,H,I,J）
  const existingJ = sheetRows[i][8];
  const existingK = sheetRows[i][9];
  if (existingJ && existingK) continue; // 既に両方値がある

  const webEntry = webMap[sheetDate];
  if (webEntry) {
    const genkin = parseSignedNumber(webEntry.kojinGenkin);
    const shinyo = parseSignedNumber(webEntry.kojinShinyo);
    if (genkin !== null || shinyo !== null) {
      updates.push({ row: i + 3, date: sheetDate, genkin, shinyo });
    }
  }
}

console.log(`\n更新対象: ${updates.length}行`);
for (const u of updates) {
  console.log(`  行${u.row}: ${u.date} → 個人(現金): ${u.genkin}, 個人(信用): ${u.shinyo}`);
}

// --- Step 6: バッチ更新 ---
if (updates.length > 0) {
  const data = updates.map(u => ({
    range: `'投資主体別'!J${u.row}:K${u.row}`,
    values: [[u.genkin, u.shinyo]],
  }));

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'RAW',
      data,
    },
  });
  console.log('\nバックフィル完了！');
} else {
  console.log('\n更新対象がありません。');
}
