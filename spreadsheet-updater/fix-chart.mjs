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
const sheetId = 1978398551;
const endRow = 1234;

// チャートを再構築（headerCount: 1、startRowIndex: 1で系列名を正しく取得）
const spec = {
  basicChart: {
    chartType: 'LINE',
    axis: [
      { position: 'BOTTOM_AXIS', viewWindowOptions: {} },
      { position: 'LEFT_AXIS', viewWindowOptions: { viewWindowMin: 25000 } },
    ],
    domains: [{
      domain: {
        sourceRange: {
          sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 1, endColumnIndex: 2 }],
        },
      },
    }],
    series: [
      // C列: 日経平均株価
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 2, endColumnIndex: 3 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 2 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // I列: PER 13.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 8, endColumnIndex: 9 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // J列: PER 14.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 9, endColumnIndex: 10 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // L列: PER 15.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 11, endColumnIndex: 12 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // N列: PER 16.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 13, endColumnIndex: 14 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // P列: PER 17.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 15, endColumnIndex: 16 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // R列: PER 18.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 17, endColumnIndex: 18 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // S列: PER 19.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 18, endColumnIndex: 19 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // T列: PER 20.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 19, endColumnIndex: 20 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
      // U列: PER 21.0
      {
        series: { sourceRange: { sources: [{ sheetId, startRowIndex: 1, endRowIndex: endRow, startColumnIndex: 20, endColumnIndex: 21 }] } },
        targetAxis: 'LEFT_AXIS',
        lineStyle: { width: 1 },
        dataLabel: { type: 'NONE', textFormat: { fontFamily: 'Roboto' } },
      },
    ],
    headerCount: 1,
    lineSmoothing: true,
  },
  hiddenDimensionStrategy: 'SKIP_HIDDEN_ROWS_AND_COLUMNS',
  fontName: 'Roboto',
};

await sheets.spreadsheets.batchUpdate({
  spreadsheetId,
  requestBody: {
    requests: [{
      updateChartSpec: { chartId, spec },
    }],
  },
});

console.log('グラフを修正しました（凡例ラベル復元 + PER 21.0追加）');
