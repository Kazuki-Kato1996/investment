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
const sheetId = 1978398551;

// スプレッドシートの既存日付を取得
const res = await sheets.spreadsheets.values.get({
  spreadsheetId,
  range: "'PER推移'!B3:B100",
  valueRenderOption: 'FORMATTED_VALUE',
});
const existingDates = new Set((res.data.values || []).map(r => r[0]).filter(Boolean));
console.log('=== スプレッドシート既存日付 ===');
console.log([...existingDates].slice(0, 20).join(', '));

// ウェブからPERデータを取得
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage();
await page.goto('https://nikkei225jp.com/data/per.php', { waitUntil: 'networkidle' });
await page.waitForSelector('#datatbl tr td', { timeout: 10000 });

const webData = await page.evaluate(() => {
  const table = document.getElementById('datatbl');
  const rows = table.querySelectorAll('tr');
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const cells = rows[i].querySelectorAll('td');
    if (cells.length >= 8) {
      result.push({
        date: cells[0]?.textContent?.trim(),
        nikkei: cells[1]?.textContent?.trim(),
        eps: cells[6]?.textContent?.trim(),
        bps: cells[7]?.textContent?.trim(),
      });
    }
  }
  return result;
});
await browser.close();

console.log(`\nウェブデータ: ${webData.length}行`);

// 2/2〜3/9 の範囲で抜けている日付を特定
const missingData = [];
for (const d of webData) {
  const normalized = d.date.replace(/-/g, '/');
  // 2026/02/02 〜 2026/03/09 の範囲
  if (normalized >= '2026/02/03' && normalized <= '2026/03/09') {
    if (!existingDates.has(normalized)) {
      missingData.push(d);
    }
  }
}

console.log(`\n=== 抜けている日付: ${missingData.length}件 ===`);
for (const d of missingData) {
  console.log(`  ${d.date}: 日経=${d.nikkei}, EPS=${d.eps}, BPS=${d.bps}`);
}

if (missingData.length === 0) {
  console.log('抜けはありません。');
  process.exit(0);
}

// 日付をExcelシリアル値に変換
function dateToSerial(dateStr) {
  const normalized = dateStr.replace(/-/g, '/');
  const parts = normalized.split('/');
  const date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const epoch = new Date(1899, 11, 30);
  return Math.round((date - epoch) / (24 * 60 * 60 * 1000));
}

function parseNumber(str) {
  if (!str || str === '-' || str === '') return null;
  return Number(str.replace(/,/g, '').replace(/[+▲▼]/g, '').trim());
}

// 古い日付から順に挿入（行3に挿入するので、新しい日付が上になる）
// → 新しい日付から順に挿入する
missingData.sort((a, b) => a.date.localeCompare(b.date)); // 古い順
// 逆順にして新しい日付から挿入
missingData.reverse();

for (const d of missingData) {
  console.log(`\n${d.date} を挿入中...`);

  // 行3に空行を挿入
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        insertDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
          inheritFromBefore: false,
        },
      }],
    },
  });

  const nikkei = parseNumber(d.nikkei);
  const eps = parseNumber(d.eps);
  const bps = parseNumber(d.bps);

  const values = [[
    '',
    dateToSerial(d.date),
    nikkei,
    '=(C3-C4)/C4',
    eps,
    '=(E3-E4)/E4',
    '=C3/E3',
    '=$E3*H$2',
    '=$E3*I$2',
    '=$E3*J$2',
    '=$E3*$K$2',
    '=$E3*L$2',
    '=$E3*M$2',
    '=$E3*N$2',
    '=$E3*O$2',
    '=$E3*P$2',
    '=$E3*Q$2',
    '=$E3*R$2',
    '=$E3*S$2',
    '=$E3*T$2',
    bps,
    '=C3/U3',
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: "'PER推移'!A3:V3",
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });

  console.log(`  完了: 日経=${nikkei}, EPS=${eps}, BPS=${bps}`);
}

console.log(`\n全${missingData.length}件の挿入が完了しました。`);
