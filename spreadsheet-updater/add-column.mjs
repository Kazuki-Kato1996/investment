import { google } from 'googleapis';
import { readFileSync } from 'fs';

const credentials = JSON.parse(readFileSync('../investment-claude-e261a94c1e98.json', 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';
const sheetId = 2061132662;

// H列（index 7）に事業法人の列を挿入
await sheets.spreadsheets.batchUpdate({
  spreadsheetId,
  requestBody: {
    requests: [
      {
        insertDimension: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 7, // H列
            endIndex: 8,
          },
          inheritFromBefore: true,
        },
      },
    ],
  },
});

// H2に「事業法人」ヘッダーを追加
await sheets.spreadsheets.values.update({
  spreadsheetId,
  range: "'投資主体別'!H2",
  valueInputOption: 'RAW',
  requestBody: {
    values: [['事業法人']],
  },
});

console.log('事業法人の列（H列）を追加しました');
