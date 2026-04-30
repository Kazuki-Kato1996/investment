import { google } from 'googleapis';
import { readFileSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const outDir = __dirname;

const credentials = JSON.parse(
  readFileSync('/Users/k.kato/Claude/investment/investment-claude-e261a94c1e98.json', 'utf8')
);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
});
const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

// 対象シート: 年間集計 + 月次シート（2026.x）すべて
const targetGids = [1737242053, 1181632804];
const targetTitlePattern = /^2026\.\d+$/;

const label = process.argv[2] || 'snapshot';

// シート一覧を取得してgid→名前のマップを作る
const meta = await sheets.spreadsheets.get({ spreadsheetId });
const allSheets = meta.data.sheets.map((s) => ({
  gid: s.properties.sheetId,
  title: s.properties.title,
  rowCount: s.properties.gridProperties.rowCount,
  colCount: s.properties.gridProperties.columnCount,
}));

console.log('All sheets:');
allSheets.forEach((s) => console.log(`  gid=${s.gid}\t${s.title}`));

const targets = allSheets.filter(
  (s) => targetGids.includes(s.gid) || targetTitlePattern.test(s.title)
);
console.log(`\nTarget sheets:`);
targets.forEach((s) => console.log(`  gid=${s.gid}\t${s.title}`));

const snapshot = {
  takenAt: new Date().toISOString(),
  label,
  allSheetTitles: allSheets.map((s) => s.title),
  sheets: {},
};

for (const t of targets) {
  console.log(`\nReading: ${t.title} (gid=${t.gid})`);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${t.title}'!A1:AZ500`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const formulas = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${t.title}'!A1:AZ500`,
    valueRenderOption: 'FORMULA',
  });
  snapshot.sheets[t.title] = {
    gid: t.gid,
    values: res.data.values || [],
    formulas: formulas.data.values || [],
  };
  console.log(`  ${(res.data.values || []).length} rows`);
}

const outPath = join(outDir, `snapshot-${label}.json`);
writeFileSync(outPath, JSON.stringify(snapshot, null, 2));
console.log(`\nSaved: ${outPath}`);
