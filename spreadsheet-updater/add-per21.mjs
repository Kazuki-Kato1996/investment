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

// 1. T列(index 19)とU列(index 20)の間に列を挿入 → U列(index 20)に挿入
await sheets.spreadsheets.batchUpdate({
  spreadsheetId,
  requestBody: {
    requests: [{
      insertDimension: {
        range: {
          sheetId,
          dimension: 'COLUMNS',
          startIndex: 20, // U列の位置に挿入
          endIndex: 21,
        },
        inheritFromBefore: true,
      },
    }],
  },
});
console.log('U列を挿入しました。');

// 2. U2に「21.0」ヘッダーを設定
await sheets.spreadsheets.values.update({
  spreadsheetId,
  range: "'PER推移'!U2",
  valueInputOption: 'RAW',
  requestBody: { values: [[21.0]] },
});
console.log('U2に21.0を設定しました。');

// 3. 全データ行にPER21倍の数式を入力
const res = await sheets.spreadsheets.values.get({
  spreadsheetId,
  range: "'PER推移'!B3:B200",
  valueRenderOption: 'FORMATTED_VALUE',
});
const rowCount = res.data.values?.filter(r => r[0]).length || 0;
console.log(`データ行数: ${rowCount}`);

// U3〜U(rowCount+2) に数式 =$E3*U$2 を一括入力
const formulas = [];
for (let i = 3; i <= rowCount + 2; i++) {
  formulas.push([`=$E${i}*U$2`]);
}

await sheets.spreadsheets.values.update({
  spreadsheetId,
  range: `'PER推移'!U3:U${rowCount + 2}`,
  valueInputOption: 'USER_ENTERED',
  requestBody: { values: formulas },
});
console.log(`U3:U${rowCount + 2} に数式を入力しました。`);

// 確認
const check = await sheets.spreadsheets.values.get({
  spreadsheetId,
  range: "'PER推移'!T2:V5",
  valueRenderOption: 'FORMATTED_VALUE',
});
console.log('\n=== 確認（T-V列, 行2-5） ===');
check.data.values?.forEach((row, i) => console.log(`行${i+2}: ${JSON.stringify(row)}`));
