import { google } from 'googleapis';
import { readFileSync } from 'fs';

const credentials = JSON.parse(readFileSync('../investment-claude-e261a94c1e98.json', 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

const SHUTAI_SHEET_ID = 2061132662;
const SINYOU_SHEET_ID = 798317965;

// 既存の条件付き書式を確認・削除してから新規追加
const res = await sheets.spreadsheets.get({
  spreadsheetId,
  fields: 'sheets.conditionalFormats,sheets.properties.sheetId',
});

// 既存の条件付き書式のインデックスを収集（対象シートのみ）
const deleteRequests = [];
for (const sheet of res.data.sheets) {
  const sheetId = sheet.properties.sheetId;
  if (sheetId === SHUTAI_SHEET_ID || sheetId === SINYOU_SHEET_ID) {
    if (sheet.conditionalFormats) {
      for (let i = sheet.conditionalFormats.length - 1; i >= 0; i--) {
        deleteRequests.push({
          deleteConditionalFormatRule: { sheetId, index: i },
        });
      }
    }
  }
}

// 条件付き書式ルールを作成
const formatRequests = [];

// 投資主体別: C3:H1000 (海外, 個人, 投資信託, 信託銀行, 証券自己, 事業法人)
// プラス → 緑背景 + 黒文字
formatRequests.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [{
        sheetId: SHUTAI_SHEET_ID,
        startRowIndex: 2,
        endRowIndex: 1000,
        startColumnIndex: 2, // C列
        endColumnIndex: 11,  // K列まで
      }],
      booleanRule: {
        condition: {
          type: 'NUMBER_GREATER_THAN_EQ',
          values: [{ userEnteredValue: '0' }],
        },
        format: {
          backgroundColor: { red: 0, green: 1, blue: 0 },
          textFormat: { foregroundColor: { red: 0, green: 0, blue: 0 } },
        },
      },
    },
    index: 0,
  },
});

// マイナス → 赤背景 + 白文字
formatRequests.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [{
        sheetId: SHUTAI_SHEET_ID,
        startRowIndex: 2,
        endRowIndex: 1000,
        startColumnIndex: 2,
        endColumnIndex: 11,  // K列まで
      }],
      booleanRule: {
        condition: {
          type: 'NUMBER_LESS',
          values: [{ userEnteredValue: '0' }],
        },
        format: {
          backgroundColor: { red: 1, green: 0, blue: 0 },
          textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 } },
        },
      },
    },
    index: 1,
  },
});

// 信用倍率: F3:F1000 (信用評価率)
// プラス → 緑背景 + 黒文字
formatRequests.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [{
        sheetId: SINYOU_SHEET_ID,
        startRowIndex: 2,
        endRowIndex: 1000,
        startColumnIndex: 5, // F列（信用評価率）
        endColumnIndex: 6,
      }],
      booleanRule: {
        condition: {
          type: 'NUMBER_GREATER_THAN_EQ',
          values: [{ userEnteredValue: '0' }],
        },
        format: {
          backgroundColor: { red: 0, green: 1, blue: 0 },
          textFormat: { foregroundColor: { red: 0, green: 0, blue: 0 } },
        },
      },
    },
    index: 0,
  },
});

// マイナス → 赤背景 + 白文字
formatRequests.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [{
        sheetId: SINYOU_SHEET_ID,
        startRowIndex: 2,
        endRowIndex: 1000,
        startColumnIndex: 5,
        endColumnIndex: 6,
      }],
      booleanRule: {
        condition: {
          type: 'NUMBER_LESS',
          values: [{ userEnteredValue: '0' }],
        },
        format: {
          backgroundColor: { red: 1, green: 0, blue: 0 },
          textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 } },
        },
      },
    },
    index: 1,
  },
});

// 実行
const allRequests = [...deleteRequests, ...formatRequests];
await sheets.spreadsheets.batchUpdate({
  spreadsheetId,
  requestBody: { requests: allRequests },
});

console.log('条件付き書式を適用しました:');
console.log('  投資主体別 C3:K1000 - プラス:緑+黒文字 / マイナス:赤+白文字');
console.log('  信用倍率   F3:F1000 - プラス:緑+黒文字 / マイナス:赤+白文字');
