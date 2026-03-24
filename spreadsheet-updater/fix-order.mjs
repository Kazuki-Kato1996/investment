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

// まず全データ行数を確認
const res = await sheets.spreadsheets.values.get({
  spreadsheetId,
  range: "'PER推移'!B3:B200",
  valueRenderOption: 'FORMATTED_VALUE',
});
const rowCount = res.data.values?.filter(r => r[0]).length || 0;
console.log(`データ行数: ${rowCount}`);

// B列（日付）で降順ソート（行3〜最終行）
await sheets.spreadsheets.batchUpdate({
  spreadsheetId,
  requestBody: {
    requests: [{
      sortRange: {
        range: {
          sheetId,
          startRowIndex: 2,           // 行3（0-indexed）
          endRowIndex: 2 + rowCount,  // データの最終行まで
          startColumnIndex: 0,        // A列
          endColumnIndex: 22,         // V列まで
        },
        sortSpecs: [{
          dimensionIndex: 1,  // B列（日付）でソート
          sortOrder: 'DESCENDING',
        }],
      },
    }],
  },
});

console.log('降順に並び替えました。');

// 確認
const check = await sheets.spreadsheets.values.get({
  spreadsheetId,
  range: "'PER推移'!B3:C10",
  valueRenderOption: 'FORMATTED_VALUE',
});
console.log('\n=== 並び替え後（先頭8行） ===');
check.data.values?.forEach((row, i) => console.log(`行${i+3}: ${row[0]} | ${row[1]}`));
