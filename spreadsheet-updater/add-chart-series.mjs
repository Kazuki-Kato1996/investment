import { google } from 'googleapis';
import { readFileSync } from 'fs';

const credentials = JSON.parse(readFileSync('../investment-claude-e261a94c1e98.json', 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

const chartId = 693971544;

// まず現在のチャート情報を取得して既存seriesを確認
const res = await sheets.spreadsheets.get({
  spreadsheetId,
  fields: 'sheets.charts',
});

let chart = null;
for (const sheet of res.data.sheets) {
  if (sheet.charts) {
    for (const c of sheet.charts) {
      if (c.chartId === chartId) {
        chart = c;
        break;
      }
    }
  }
}

if (!chart) {
  console.error('チャートが見つかりません');
  process.exit(1);
}

const currentSeriesCount = chart.spec.basicChart.series.length;
console.log(`現在のシリーズ数: ${currentSeriesCount}`);

// 既存のシリーズの列インデックスを確認
for (let i = 0; i < chart.spec.basicChart.series.length; i++) {
  const s = chart.spec.basicChart.series[i];
  const src = s.series.sourceRange.sources[0];
  console.log(`  series[${i}]: column ${src.startColumnIndex}-${src.endColumnIndex}`);
}

// PER 21.0 (U列 = columnIndex 20) のシリーズを最後の位置に追加
// PER 20.0の次（最後のシリーズの後）に追加
await sheets.spreadsheets.batchUpdate({
  spreadsheetId,
  requestBody: {
    requests: [{
      updateChartSpec: {
        chartId,
        spec: {
          ...chart.spec,
          basicChart: {
            ...chart.spec.basicChart,
            series: [
              ...chart.spec.basicChart.series,
              {
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId: 1978398551,
                      startRowIndex: 0,
                      endRowIndex: 234,
                      startColumnIndex: 20,  // U列 (PER 21.0)
                      endColumnIndex: 21,
                    }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                lineStyle: { width: 1 },
                dataLabel: {
                  type: 'NONE',
                  textFormat: { fontFamily: 'Roboto' },
                },
              },
            ],
          },
        },
      },
    }],
  },
});

console.log('PER 21.0のラインをグラフに追加しました。');
