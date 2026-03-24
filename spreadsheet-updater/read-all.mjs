import { google } from 'googleapis';
import { readFileSync, writeFileSync } from 'fs';

const credentials = JSON.parse(readFileSync('../investment-claude-e261a94c1e98.json', 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

const sheetNames = ['年間集計', '2026.3', '2026.2', '2026.1', 'PER推移', '信用倍率', '投資主体別', '購入価額'];

const allData = {};

for (const name of sheetNames) {
  console.log(`Reading: ${name}`);
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${name}'!A1:AZ200`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    allData[name] = res.data.values || [];
    console.log(`  ${allData[name].length} rows`);
  } catch (e) {
    console.log(`  Error: ${e.message}`);
  }
}

writeFileSync('sheet-data.json', JSON.stringify(allData, null, 2));
console.log('\nSaved to sheet-data.json');
