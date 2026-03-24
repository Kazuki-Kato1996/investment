import { chromium } from 'playwright';
import { google } from 'googleapis';
import { readFileSync } from 'fs';

const credentials = JSON.parse(readFileSync('../investment-claude-e261a94c1e98.json', 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

// スプレッドシートの既存日付を取得
const res = await sheets.spreadsheets.values.get({
  spreadsheetId,
  range: "'投資主体別'!B3:H100",
  valueRenderOption: 'FORMATTED_VALUE',
});
const sheetRows = res.data.values || [];
console.log(`スプレッドシート: ${sheetRows.length}行`);

// ウェブから全データを取得
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
    if (cells.length >= 10) {
      result.push({
        date: cells[0]?.textContent?.trim(),
        jigyoHojin: cells[9]?.textContent?.trim(),
      });
    }
  }
  return result;
});
await browser.close();

console.log(`ウェブデータ: ${webData.length}行`);

// 日付をキーにしたマップを作成
const webMap = {};
for (const d of webData) {
  webMap[d.date] = d.jigyoHojin;
}

// マッチングして更新
const updates = [];
for (let i = 0; i < sheetRows.length; i++) {
  const sheetDate = sheetRows[i][0]; // B列の日付
  const existingH = sheetRows[i][6]; // H列（事業法人）
  if (existingH) continue; // 既に値がある

  // 日付フォーマットを合わせる（スプレッドシート: "2026/02/27", ウェブ: "2026/02/27"）
  const webValue = webMap[sheetDate];
  if (webValue) {
    const num = Number(webValue.replace(/,/g, '').replace(/[+▲▼]/g, '').trim());
    const signed = webValue.startsWith('-') || webValue.startsWith('▼') ? -num : num;
    updates.push({ row: i + 3, date: sheetDate, value: signed });
  }
}

console.log(`\n更新対象: ${updates.length}行`);
for (const u of updates) {
  console.log(`  行${u.row}: ${u.date} → 事業法人: ${u.value}`);
}

// バッチ更新
if (updates.length > 0) {
  const data = updates.map(u => ({
    range: `'投資主体別'!H${u.row}`,
    values: [[u.value]],
  }));

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: 'RAW',
      data,
    },
  });
  console.log('\n更新完了！');
} else {
  console.log('\n更新対象がありません。');
}
