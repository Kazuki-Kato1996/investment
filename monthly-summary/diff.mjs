import { readFileSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

const beforeLabel = process.argv[2] || 'before';
const afterLabel = process.argv[3] || 'after';

const before = JSON.parse(
  readFileSync(join(__dirname, `snapshot-${beforeLabel}.json`), 'utf8')
);
const after = JSON.parse(
  readFileSync(join(__dirname, `snapshot-${afterLabel}.json`), 'utf8')
);

// 列番号→A1表記
function colToLetter(col) {
  let s = '';
  let n = col;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

const result = {
  beforeAt: before.takenAt,
  afterAt: after.takenAt,
  newSheets: [],
  removedSheets: [],
  sheetDiffs: {},
};

const beforeTitles = new Set(before.allSheetTitles);
const afterTitles = new Set(after.allSheetTitles);
result.newSheets = [...afterTitles].filter((t) => !beforeTitles.has(t));
result.removedSheets = [...beforeTitles].filter((t) => !afterTitles.has(t));

const allTitles = new Set([
  ...Object.keys(before.sheets),
  ...Object.keys(after.sheets),
]);

for (const title of allTitles) {
  const b = before.sheets[title];
  const a = after.sheets[title];
  const diffs = [];

  if (!b && a) {
    diffs.push({ type: 'sheet_added', title });
  } else if (b && !a) {
    diffs.push({ type: 'sheet_removed', title });
  } else if (b && a) {
    const bv = b.values;
    const av = a.values;
    const bf = b.formulas;
    const af = a.formulas;
    const maxRow = Math.max(bv.length, av.length);
    for (let r = 0; r < maxRow; r++) {
      const br = bv[r] || [];
      const ar = av[r] || [];
      const bfr = bf[r] || [];
      const afr = af[r] || [];
      const maxCol = Math.max(br.length, ar.length, bfr.length, afr.length);
      for (let c = 0; c < maxCol; c++) {
        const bVal = br[c] ?? '';
        const aVal = ar[c] ?? '';
        const bForm = bfr[c] ?? '';
        const aForm = afr[c] ?? '';
        if (bVal !== aVal || bForm !== aForm) {
          diffs.push({
            cell: `${colToLetter(c)}${r + 1}`,
            row: r + 1,
            col: c,
            before: { value: bVal, formula: bForm },
            after: { value: aVal, formula: aForm },
          });
        }
      }
    }
  }

  if (diffs.length > 0) {
    result.sheetDiffs[title] = diffs;
  }
}

const outPath = join(__dirname, `diff-${beforeLabel}-vs-${afterLabel}.json`);
writeFileSync(outPath, JSON.stringify(result, null, 2));

// サマリ表示
console.log(`Before: ${result.beforeAt}`);
console.log(`After:  ${result.afterAt}`);
console.log(`\nNew sheets: ${result.newSheets.join(', ') || '(none)'}`);
console.log(`Removed sheets: ${result.removedSheets.join(', ') || '(none)'}`);
console.log('\nSheet diffs:');
for (const [title, diffs] of Object.entries(result.sheetDiffs)) {
  console.log(`  ${title}: ${diffs.length} changes`);
}
console.log(`\nSaved: ${outPath}`);
