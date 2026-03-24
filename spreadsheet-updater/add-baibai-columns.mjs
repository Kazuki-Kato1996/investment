/**
 * PER推移シートにY,Z,AA,AB列の売買代金ヘッダーを追加するスクリプト
 *
 * 使い方:
 *   node add-baibai-columns.mjs
 */

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

// Y2, Z2, AA2, AB2 にヘッダーを書き込み
console.log('PER推移シートに売買代金ヘッダーを書き込み中...');
await sheets.spreadsheets.values.update({
  spreadsheetId: SPREADSHEET_ID,
  range: "'PER推移'!Y2:AB2",
  valueInputOption: 'USER_ENTERED',
  requestBody: {
    values: [['全市場売買代金(億円)', 'プライム売買代金(億円)', 'スタンダード売買代金(億円)', 'グロース売買代金(億円)']],
  },
});

console.log('ヘッダー書き込み完了:');
console.log('  Y2: 全市場売買代金(億円)');
console.log('  Z2: プライム売買代金(億円)');
console.log('  AA2: スタンダード売買代金(億円)');
console.log('  AB2: グロース売買代金(億円)');
