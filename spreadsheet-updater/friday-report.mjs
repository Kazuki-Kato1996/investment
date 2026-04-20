#!/usr/bin/env node
/**
 * 金曜投資評価レポート生成・メール送信スクリプト
 *
 * Googleスプレッドシートからポートフォリオデータ・市場指標を読み取り、
 * 投資評価レポート形式のHTMLメールを生成・送信する。
 *
 * 使い方:
 *   node friday-report.mjs               # メール送信
 *   node friday-report.mjs --dry-run     # コンソール出力のみ
 */

import { google } from 'googleapis';
import { createTransport } from 'nodemailer';
import { readFileSync, existsSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const CREDENTIALS_PATH = join(__dirname, '..', 'investment-claude-e261a94c1e98.json');
const EMAIL_CONFIG_PATH = join(__dirname, '..', 'morning-report', 'market-report-config.json');
const SPREADSHEET_ID = '12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk';

// --- オーナーコメント読み取り（月次シートのE2:E6） ---
async function readOwnerComment(sheets, sheetName) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!E2:E6`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const rows = res.data.values || [];
    const lines = rows.map(r => (r[0] || '').replace(/^\/\/\s?/, '').trim()).filter(l => l.length > 0);
    return lines.join('\n');
  } catch (e) {
    console.error(`  オーナーコメント読み取りエラー: ${e.message}`);
    return '';
  }
}

// --- Google Sheets認証 ---
function getSheets() {
  const credentials = JSON.parse(readFileSync(CREDENTIALS_PATH, 'utf8'));
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  return google.sheets({ version: 'v4', auth });
}

// --- メール設定読み込み ---
function loadEmailConfig() {
  if (!existsSync(EMAIL_CONFIG_PATH)) {
    console.error(`メール設定ファイルが見つかりません: ${EMAIL_CONFIG_PATH}`);
    process.exit(1);
  }
  return JSON.parse(readFileSync(EMAIL_CONFIG_PATH, 'utf-8'));
}

// --- ユーティリティ ---
function parseNum(str) {
  if (!str || str === '-' || str === '' || str === '－') return 0;
  return Number(String(str).replace(/,/g, '').replace(/[▲▼ ]/g, '').trim()) || 0;
}

function parseSignedNum(str) {
  if (!str || str === '-' || str === '' || str === '－') return 0;
  const cleaned = String(str).replace(/,/g, '').trim();
  if (cleaned.startsWith('▲') || cleaned.startsWith('▼') || cleaned.startsWith('-')) {
    return -Math.abs(Number(cleaned.replace(/[▲▼\-+, ]/g, '')));
  }
  return Number(cleaned.replace(/[+, ]/g, '')) || 0;
}

function fmt(n) {
  if (n === 0 || n === null || n === undefined) return '-';
  return Number(n).toLocaleString('ja-JP');
}

function fmtPct(str) {
  if (!str || str === '-' || str === '') return '-';
  return String(str);
}

function pctClass(str) {
  if (!str) return '';
  const s = String(str);
  if (s.includes('-') || s.includes('▲') || s.includes('▼')) return 'minus';
  if (s.match(/[0-9]/)) return 'plus';
  return '';
}

function signedFmt(n) {
  if (n === 0 || n === null) return '-';
  const prefix = n > 0 ? '+' : '';
  return prefix + Number(n).toLocaleString('ja-JP');
}

// --- スプレッドシートからデータ読み込み ---
async function readSheet(sheets, sheetName, range = 'A1:AZ200') {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!${range}`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    return res.data.values || [];
  } catch (e) {
    console.error(`  シート「${sheetName}」の読み込みエラー: ${e.message}`);
    return [];
  }
}

// 現在の月シート名を取得（例: "2026.3"）
function getCurrentMonthSheetName() {
  const now = new Date();
  return `${now.getFullYear()}.${now.getMonth() + 1}`;
}

// 前月のシート名を取得
function getPrevMonthSheetName() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${prev.getFullYear()}.${prev.getMonth() + 1}`;
}

// --- 年間集計パース ---
function parseAnnualSummary(rows) {
  // row[0]=header, row[1]=column names, row[2]=12月, row[3]=1月, ...
  const months = [];
  for (let i = 2; i < rows.length && i < 20; i++) {
    const r = rows[i];
    if (!r || !r[1] || !r[1].match(/[0-9]月|12月/)) continue;
    // データがない月はスキップ
    if (!r[9] || r[9].trim() === '') continue;
    months.push({
      month: r[1]?.trim(),
      costBasis: r[2]?.trim(),
      unrealizedPL: r[3]?.trim(),
      marketValue: r[4]?.trim(),
      okachCash: r[5]?.trim(),
      sbiCash: r[6]?.trim(),
      sbiUS: r[7]?.trim(),
      realizedPL: r[8]?.trim(),
      totalAssets: r[9]?.trim(),
      evalRate: r[10]?.trim(),
      monthlyChange: r[11]?.trim(),
      annualPL: r[12]?.trim(),
      nikkeiBase: r[14]?.trim(),
      nikkei: r[15]?.trim(),
      nikkeiChange: r[16]?.trim(),
      nikkeiAnnual: r[17]?.trim(),
      topixBase: r[18]?.trim(),
      topix: r[19]?.trim(),
      topixChange: r[20]?.trim(),
      topixAnnual: r[21]?.trim(),
      sp500Base: r[22]?.trim(),
      sp500: r[23]?.trim(),
      sp500Change: r[24]?.trim(),
      sp500Annual: r[25]?.trim(),
    });
  }
  return months;
}

// --- 月別シートから保有銘柄をパース ---
function parseMonthlySheet(rows) {
  const sections = { rakuten: [], sbiCash: [], sbiMargin: [], okachi: [], sbiUS: [] };
  let currentSection = null;
  let summaries = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const joined = (r || []).join('');

    // セクション検出
    if (joined.includes('楽天証券') && !joined.includes('売却日')) {
      currentSection = 'rakuten';
      continue;
    }
    if (joined.includes('SBI証券　信用取引') || joined.includes('SBI証券　信用')) {
      currentSection = 'sbiMargin';
      continue;
    }
    if (joined.includes('SBI証券') && !joined.includes('信用') && !joined.includes('売却日') && !joined.includes('ドル')) {
      // SBI証券のヘッダー行を区別
      if (i > 100) {
        currentSection = 'sbiUS';
      } else if (currentSection !== 'sbiMargin') {
        currentSection = 'sbiCash';
      }
      continue;
    }
    if (joined.includes('岡地証券')) {
      currentSection = 'okachi';
      continue;
    }
    if (joined.includes('SBIドル資産')) {
      currentSection = 'sbiUS';
      continue;
    }

    // ヘッダー行スキップ
    if (joined.includes('売却日') && joined.includes('購入日')) continue;

    // 集計行の検出（保有証券原価/有価証券原価/取得価額 + 保有時価/月末時価 + 総評価損益）
    if ((joined.includes('保有証券原価') || joined.includes('有価証券原価') || joined.includes('取得価額'))
        && (joined.includes('保有時価') || joined.includes('月末時価'))
        && joined.includes('総評価損益')) {
      // 総計行は別扱い
      if (joined.includes('総計')) {
        const nextRow = rows[i + 1];
        if (nextRow) {
          summaries['total'] = {
            costBasis: nextRow[9]?.trim(),
            marketValue: nextRow[10]?.trim(),
            unrealizedPL: nextRow[11]?.trim(),
            evalRate: nextRow[12]?.trim(),
            realizedPL: nextRow[14]?.trim(),
          };
        }
      } else {
        const nextRow = rows[i + 1];
        if (nextRow && currentSection) {
          summaries[currentSection] = {
            costBasis: nextRow[9]?.trim(),
            marketValue: nextRow[10]?.trim(),
            unrealizedPL: nextRow[11]?.trim(),
            evalRate: nextRow[12]?.trim(),
            realizedPL: nextRow[14]?.trim(),
          };
        }
      }
      continue;
    }

    // 保有銘柄の行をパース（コード列が数値またはアルファベット）
    if (!currentSection || !r || r.length < 10) continue;
    const code = r[3]?.trim();
    const name = r[4]?.trim();
    if (!code || !name || name === '' || joined.includes('現金') || joined.includes('ドル変換')) continue;
    if (!code.match(/^[0-9A-Z]/)) continue;

    // 売却済み（売却日あり＆保有証券原価なし）はスキップ
    const sellDate = r[1]?.trim();
    const holdingCost = r[9]?.trim();
    if (sellDate && sellDate !== 'L' && !holdingCost) continue;

    const holding = {
      code,
      name,
      type: r[5]?.trim() || '',
      shares: r[6]?.trim() || '',
      price: r[7]?.trim() || '',
      costBasis: r[8]?.trim() || '',
      holdingCost: holdingCost || '',
      currentPrice: r[10]?.trim() || '',
      unrealizedPL: r[11]?.trim() || '',
      realizedPL: r[14]?.trim() || '',
      plRate: r[15]?.trim() || '',
    };

    if (currentSection && holding.holdingCost) {
      sections[currentSection].push(holding);
    }
  }

  return { sections, summaries };
}

// --- PER推移パース ---
function parsePERData(rows) {
  const data = [];
  for (let i = 2; i < Math.min(rows.length, 12); i++) {
    const r = rows[i];
    if (!r || !r[1]) continue;
    data.push({
      date: r[1]?.trim(),
      nikkei: r[2]?.trim(),
      change: r[3]?.trim(),
      eps: r[4]?.trim(),
      per: r[6]?.trim(),
      bps: r[21]?.trim(),
      pbr: r[22]?.trim(),
    });
  }
  return data;
}

// --- 信用倍率パース ---
function parseCreditData(rows) {
  const data = [];
  for (let i = 2; i < Math.min(rows.length, 8); i++) {
    const r = rows[i];
    if (!r || !r[1]) continue;
    data.push({
      date: r[1]?.trim(),
      sellBalance: r[2]?.trim(),
      buyBalance: r[3]?.trim(),
      creditRatio: r[4]?.trim(),
      evalRate: r[5]?.trim() || '-',
      nikkei: r[6]?.trim(),
    });
  }
  return data;
}

// --- 投資主体別パース ---
function parseInvestorData(rows) {
  const data = [];
  for (let i = 2; i < Math.min(rows.length, 7); i++) {
    const r = rows[i];
    if (!r || !r[1]) continue;
    data.push({
      date: r[1]?.trim(),
      foreign: r[2]?.trim(),
      individual: r[3]?.trim(),
      investTrust: r[4]?.trim(),
      trustBank: r[5]?.trim(),
      securities: r[6]?.trim(),
      corporate: r[7]?.trim(),
      individualCash: r[9]?.trim(),    // J列: 個人（現金）
      individualCredit: r[10]?.trim(), // K列: 個人（信用）
    });
  }
  return data;
}

// --- オーナーコメントから銘柄別方針を抽出 ---
function parseOwnerPolicies(ownerComment, allHoldings) {
  if (!ownerComment) return {};
  const policies = {};
  // 銘柄名の出現位置でブロックを分割（銘柄名が登場する箇所〜次の銘柄名の手前まで）
  const holdingNames = allHoldings.map(h => h.name).filter(n => n);
  const text = ownerComment.replace(/\n/g, '');
  for (const name of holdingNames) {
    if (!text.includes(name)) continue;
    // 銘柄名の出現位置を特定
    const startIdx = text.indexOf(name);
    // 次の銘柄名の出現位置を探す（現在の銘柄以降で最も近いもの）
    let endIdx = text.length;
    for (const otherName of holdingNames) {
      if (otherName === name) continue;
      const otherIdx = text.indexOf(otherName, startIdx + name.length);
      if (otherIdx > startIdx && otherIdx < endIdx) {
        endIdx = otherIdx;
      }
    }
    // ブロックを抽出して末尾の句点を整理
    let block = text.slice(startIdx, endIdx).trim();
    if (!block.endsWith('。')) block += '。';
    policies[name] = block;
  }
  return policies;
}

// --- データに基づく分析コメント・戦略・アクションプランの自動生成 ---
function generateAnalysis(data) {
  const { annual, monthly, per, credit, investors } = data;
  const baseMonth = annual[0];
  const latestMonth = annual[annual.length - 1];
  const prevMonth = annual.length > 2 ? annual[annual.length - 2] : null;

  // --- 数値準備 ---
  const monthlyChg = parseFloat(String(latestMonth.monthlyChange).replace(/[%,]/g, '')) || 0;
  const annualPLPct = parseFloat(String(latestMonth.annualPL).replace(/[%,]/g, '')) || 0;
  const portfolioAnnual = annualPLPct;
  const nikkeiAnnual = parseFloat(String(latestMonth.nikkeiAnnual).replace(/[%,]/g, '')) || 0;
  const topixAnnual = parseFloat(String(latestMonth.topixAnnual).replace(/[%,]/g, '')) || 0;
  const sp500Annual = parseFloat(String(latestMonth.sp500Annual).replace(/[%,]/g, '')) || 0;
  const latestPER = per.length > 0 ? parseFloat(String(per[0].per).replace(/,/g, '')) || 0 : 0;
  const latestEPS = per.length > 0 ? parseFloat(String(per[0].eps).replace(/,/g, '')) || 0 : 0;
  const latestNikkei = per.length > 0 ? parseFloat(String(per[0].nikkei).replace(/,/g, '')) || 0 : 0;
  const latestBPS = per.length > 0 ? parseFloat(String(per[0].bps).replace(/,/g, '')) || 0 : 0;
  const latestPBR = per.length > 0 ? parseFloat(String(per[0].pbr).replace(/,/g, '')) || 0 : 0;
  const creditRatio = credit.length > 0 ? parseFloat(String(credit[0].creditRatio).replace(/,/g, '')) || 0 : 0;
  const creditEvalRate = credit.length > 0 ? parseFloat(String(credit[0].evalRate).replace(/[%,]/g, '')) || 0 : 0;
  const buyBalance = credit.length > 0 ? parseNum(credit[0].buyBalance) : 0;
  const totalAssets = parseNum(latestMonth.totalAssets);
  const okachCash = parseNum(latestMonth.okachCash);
  const sbiCash = parseNum(latestMonth.sbiCash);
  const cashTotal = okachCash + sbiCash;
  const cashRatio = totalAssets > 0 ? (cashTotal / totalAssets * 100) : 0;
  const realizedPL = parseNum(latestMonth.realizedPL);

  // 保有銘柄分析
  const allHoldings = [
    ...monthly.sections.okachi,
    ...monthly.sections.sbiCash,
    ...monthly.sections.sbiMargin,
    ...monthly.sections.rakuten,
  ].filter(h => h.holdingCost && h.plRate);

  const getRate = (h) => parseFloat(String(h.plRate).replace(/[%,]/g, '')) || 0;

  // オーナーコメントから銘柄別方針を抽出
  const ownerPolicies = parseOwnerPolicies(data.ownerComment, allHoldings);

  // 利確候補（+50%以上）
  const profitTakers = allHoldings
    .filter(h => getRate(h) >= 50)
    .sort((a, b) => getRate(b) - getRate(a));

  // 損切り候補（-15%以下）
  const stopLoss = allHoldings
    .filter(h => getRate(h) <= -15)
    .sort((a, b) => getRate(a) - getRate(b));

  // 含み損注意（-5%〜-15%）
  const watchList = allHoldings
    .filter(h => getRate(h) > -15 && getRate(h) <= -5)
    .sort((a, b) => getRate(a) - getRate(b));

  // --- 1. ポートフォリオ全体の評価コメント ---
  let portfolioComment = '';
  if (monthlyChg > 3) {
    portfolioComment = `前月比${latestMonth.monthlyChange}と好調。`;
  } else if (monthlyChg >= 0) {
    portfolioComment = `前月比${latestMonth.monthlyChange}と横ばい〜微増を維持。`;
  } else if (monthlyChg > -5) {
    portfolioComment = `前月比${latestMonth.monthlyChange}と小幅調整。`;
  } else {
    portfolioComment = `前月比${latestMonth.monthlyChange}と大きく調整。`;
  }
  portfolioComment += `年間実現損益は+${(realizedPL / 10000).toFixed(1)}万円。`;

  // ベンチマーク比較
  let benchComment = '';
  const outperformCount = [nikkeiAnnual, topixAnnual, sp500Annual].filter(b => portfolioAnnual > b).length;
  if (outperformCount === 3) {
    benchComment = `ポートフォリオは年初来で全ベンチマーク（日経${nikkeiAnnual.toFixed(2)}%、TOPIX${topixAnnual.toFixed(2)}%、S&P500${sp500Annual.toFixed(2)}%）をアウトパフォーム。`;
  } else if (outperformCount >= 1) {
    const beaten = [];
    if (portfolioAnnual > nikkeiAnnual) beaten.push('日経平均');
    if (portfolioAnnual > topixAnnual) beaten.push('TOPIX');
    if (portfolioAnnual > sp500Annual) beaten.push('S&P500');
    benchComment = `${beaten.join('・')}を上回るパフォーマンス。`;
  } else {
    benchComment = `年初来では主要指数をアンダーパフォーム。ポジション調整の検討余地あり。`;
  }

  // --- 2. 岡地コメント ---
  let okachiComment = '';
  if (monthly.summaries.okachi) {
    const okachiRate = parseFloat(String(monthly.summaries.okachi.evalRate).replace(/[%,]/g, '')) || 0;
    const okachiTop = monthly.sections.okachi.filter(h => h.holdingCost).sort((a, b) => getRate(b) - getRate(a));
    const topNames = okachiTop.slice(0, 3).map(h => `${h.name}(${h.plRate})`).join('、');
    okachiComment = `評価損益率${monthly.summaries.okachi.evalRate}。${topNames}が貢献。`;
    const okachiLoss = okachiTop.filter(h => getRate(h) < -10);
    if (okachiLoss.length > 0) {
      okachiComment += `含み損注意: ${okachiLoss.map(h => `${h.name}(${h.plRate})`).join('、')}。`;
    }
  }

  // --- 3. SBI信用コメント ---
  let marginComment = '';
  if (monthly.summaries.sbiMargin) {
    const marginPL = monthly.summaries.sbiMargin.realizedPL;
    const marginPLNum = parseSignedNum(marginPL);
    if (marginPLNum < 0) {
      marginComment = `信用取引の確定損益は${marginPL}。損切りルールの徹底が必要。`;
    } else {
      marginComment = `信用取引の確定損益は${marginPL}。`;
    }
  }

  // --- 4. 市場環境コメント ---
  let perComment = '';
  if (latestPER > 0) {
    if (latestPER > 20) perComment += `PER ${latestPER.toFixed(2)}倍は過去平均（14-16倍）対比で割高圏。`;
    else if (latestPER > 17) perComment += `PER ${latestPER.toFixed(2)}倍はやや高めだが、EPS成長を考慮すると許容範囲。`;
    else if (latestPER > 14) perComment += `PER ${latestPER.toFixed(2)}倍は適正水準。`;
    else perComment += `PER ${latestPER.toFixed(2)}倍は割安圏。投資機会の可能性。`;
    if (latestEPS > 0) {
      const per18Target = Math.round(latestEPS * 18);
      const per15Target = Math.round(latestEPS * 15);
      perComment += ` EPS ${latestEPS.toFixed(0)}円基準でPER18倍=${per18Target.toLocaleString()}円、PER15倍=${per15Target.toLocaleString()}円。`;
    }
    if (latestBPS > 0) {
      perComment += `PBR1.0倍=${latestBPS.toLocaleString()}円が理論的下値メド。`;
    }
  }

  let creditComment = '';
  if (creditRatio > 0) {
    if (creditRatio > 5) creditComment += `信用倍率${creditRatio.toFixed(2)}倍は高水準。買い残の整理に時間がかかる可能性。`;
    else if (creditRatio > 3) creditComment += `信用倍率${creditRatio.toFixed(2)}倍。需給面ではやや重い。`;
    else creditComment += `信用倍率${creditRatio.toFixed(2)}倍。需給面は改善傾向。`;
    if (creditEvalRate < -10) creditComment += ` 信用評価率${creditEvalRate.toFixed(2)}%は追証ラインに接近。投げ売りリスクに注意。`;
    else if (creditEvalRate < -5) creditComment += ` 信用評価率${creditEvalRate.toFixed(2)}%。更なる悪化なら追証リスクに注意。`;
  }

  let investorComment = '';
  if (investors.length > 0) {
    const latest = investors[0];
    const foreignVal = parseSignedNum(latest.foreign);
    const individualVal = parseSignedNum(latest.individual);
    const corporateVal = parseSignedNum(latest.corporate);
    investorComment += `海外投資家: ${foreignVal > 0 ? '買い越し' : '売り越し'}(${latest.foreign})。`;
    investorComment += `個人: ${individualVal > 0 ? '買い越し' : '売り越し'}。`;
    investorComment += `事業法人: ${corporateVal > 0 ? '買い越し（自社株買い等）' : '売り越し'}。`;
  }

  // --- 5. 投資戦略 ---
  // 市場見通し
  let marketOutlook = '';
  if (latestEPS > 0 && latestNikkei > 0) {
    const per18 = Math.round(latestEPS * 18);
    const per15 = Math.round(latestEPS * 15);
    const per20 = Math.round(latestEPS * 20);
    marketOutlook += `<p><strong>メインシナリオ:</strong> 日経平均 ${per15.toLocaleString()}〜${per18.toLocaleString()}円のレンジ（PER15-18倍）</p>`;
    marketOutlook += `<ul>`;
    if (latestPER > 18) {
      marketOutlook += `<li>現在PER${latestPER.toFixed(1)}倍はレンジ上限に接近。調整リスクに注意</li>`;
    } else if (latestPER < 15) {
      marketOutlook += `<li>現在PER${latestPER.toFixed(1)}倍はレンジ下限付近。押し目買いの好機の可能性</li>`;
    } else {
      marketOutlook += `<li>現在PER${latestPER.toFixed(1)}倍は適正レンジ内</li>`;
    }
    marketOutlook += `<li>EPS ${latestEPS.toFixed(0)}円が維持されればPER15倍=${per15.toLocaleString()}円が下値メド</li>`;
    if (creditRatio > 5) {
      marketOutlook += `<li>信用倍率${creditRatio.toFixed(1)}倍と高水準。買い残の整理に時間がかかる見込み</li>`;
    }
    marketOutlook += `</ul>`;
    marketOutlook += `<p><strong>リスクシナリオ:</strong> ${per15.toLocaleString()}円以下への急落</p>`;
    marketOutlook += `<ul><li>信用評価率悪化→追証→投げ売りの連鎖</li><li>米国景気後退懸念、円高進行</li></ul>`;
    marketOutlook += `<p><strong>上振れシナリオ:</strong> ${per20.toLocaleString()}円超</p>`;
    marketOutlook += `<ul><li>EPSの上方修正、海外マネーの流入加速</li></ul>`;
  }

  // ポートフォリオ戦略
  let portfolioStrategy = '';
  // 現金比率
  portfolioStrategy += `<div class="strategy-item"><p><strong>1. 現金比率: 現在約${cashRatio.toFixed(0)}%（${(cashTotal / 10000).toFixed(0)}万円）</strong></p>`;
  if (cashRatio > 30) {
    portfolioStrategy += `<p>現金比率は高め。急落時の買い増し余力として維持しつつ、押し目では段階的に投下を検討。</p>`;
  } else if (cashRatio > 20) {
    portfolioStrategy += `<p>現金比率は適正水準。急落時の買い増し余力を維持。</p>`;
  } else {
    portfolioStrategy += `<p>現金比率が低い。リスク管理のため一部利確で現金比率を高めることを検討。</p>`;
  }
  portfolioStrategy += `</div>`;

  // 単元株（100株）かどうか判定
  const isMinUnit = (h) => parseNum(h.shares) <= 100;

  // 利確候補
  if (profitTakers.length > 0) {
    portfolioStrategy += `<div class="strategy-item"><p><strong>2. 利益確定検討銘柄（+50%超）</strong></p><ul>`;
    for (const h of profitTakers.slice(0, 5)) {
      const policy = ownerPolicies[h.name];
      if (policy) {
        portfolioStrategy += `<li><strong>${h.name}(${h.code})</strong>: ${h.plRate} → <span style="color:#1a3a6b;">【オーナー方針】${policy}</span></li>`;
      } else if (isMinUnit(h)) {
        portfolioStrategy += `<li><strong>${h.name}(${h.code})</strong>: ${h.plRate}（${h.shares}株） → 単元株のため全売却or継続保有の判断が必要</li>`;
      } else {
        portfolioStrategy += `<li><strong>${h.name}(${h.code})</strong>: ${h.plRate}（${h.shares}株） → 一部利確を推奨</li>`;
      }
    }
    portfolioStrategy += `</ul></div>`;
  }

  // 損切り候補
  if (stopLoss.length > 0) {
    portfolioStrategy += `<div class="risk-item"><p><strong>3. 損切り検討銘柄（-15%超）</strong></p><ul>`;
    for (const h of stopLoss.slice(0, 5)) {
      const policy = ownerPolicies[h.name];
      if (policy) {
        portfolioStrategy += `<li><strong>${h.name}(${h.code})</strong>: ${h.plRate} → <span style="color:#1a3a6b;">【オーナー方針】${policy}</span></li>`;
      } else if (isMinUnit(h)) {
        portfolioStrategy += `<li><strong>${h.name}(${h.code})</strong>: ${h.plRate}（${h.shares}株） → 単元株のため全損切りor業績見極めの判断が必要</li>`;
      } else {
        portfolioStrategy += `<li><strong>${h.name}(${h.code})</strong>: ${h.plRate}（${h.shares}株） → 損切りまたは見極めが必要</li>`;
      }
    }
    portfolioStrategy += `</ul></div>`;
  }

  // 信用取引ルール
  if (monthly.summaries.sbiMargin) {
    const marginPLNum = parseSignedNum(monthly.summaries.sbiMargin.realizedPL);
    if (marginPLNum < -100000) {
      portfolioStrategy += `<div class="risk-item"><p><strong>4. 信用取引ルール改善</strong></p><ul>`;
      portfolioStrategy += `<li>損切りルール: エントリー価格から-5%で機械的に損切り</li>`;
      portfolioStrategy += `<li>ポジションサイズ: 1銘柄あたり信用建玉100万円以下に制限</li>`;
      portfolioStrategy += `<li>政策テーマ銘柄の空売りは原則禁止</li>`;
      portfolioStrategy += `</ul></div>`;
    }
  }

  // --- 6. アクションプラン ---
  const actions = [];
  let priority = 1;

  // 信用損失が大きければルール策定を最優先
  if (monthly.summaries.sbiMargin) {
    const marginPLNum = parseSignedNum(monthly.summaries.sbiMargin.realizedPL);
    if (marginPLNum < -100000) {
      actions.push({ priority: priority++, action: '信用取引の損切りルール策定・適用', timing: '即時', reason: `信用取引損失${monthly.summaries.sbiMargin.realizedPL}の防止` });
    }
  }

  // 利確候補（+100%超は即実行、+50%超は今週中）
  for (const h of profitTakers.slice(0, 3)) {
    const policy = ownerPolicies[h.name];
    if (policy) {
      actions.push({ priority: priority++, action: `${h.name}【オーナー方針】`, timing: '-', reason: policy });
    } else if (isMinUnit(h)) {
      const rate = getRate(h);
      if (rate >= 100) {
        actions.push({ priority: priority++, action: `${h.name} 全売却判断`, timing: '今週中', reason: `${h.plRate}（${h.shares}株・単元株のため一部利確不可）` });
      } else {
        actions.push({ priority: priority++, action: `${h.name} 売却/保有判断`, timing: '今月中', reason: `${h.plRate}（${h.shares}株・単元株のため一部利確不可）` });
      }
    } else {
      const rate = getRate(h);
      if (rate >= 100) {
        actions.push({ priority: priority++, action: `${h.name} 一部利確`, timing: '今週中', reason: `${h.plRate}の利益確保` });
      } else {
        actions.push({ priority: priority++, action: `${h.name} 利確検討`, timing: '今月中', reason: `${h.plRate}、上昇余地を見極め` });
      }
    }
  }

  // 損切り候補
  for (const h of stopLoss.slice(0, 3)) {
    const policy = ownerPolicies[h.name];
    if (policy) {
      actions.push({ priority: priority++, action: `${h.name}【オーナー方針】`, timing: '-', reason: policy });
    } else if (isMinUnit(h)) {
      actions.push({ priority: priority++, action: `${h.name} 全損切り/継続判断`, timing: '決算後', reason: `含み損${h.plRate}（${h.shares}株・単元株のため一部損切り不可）` });
    } else {
      const rate = getRate(h);
      if (rate <= -30) {
        actions.push({ priority: priority++, action: `${h.name} 損切り`, timing: '今月中', reason: `含み損${h.plRate}の拡大防止` });
      } else {
        actions.push({ priority: priority++, action: `${h.name} 見極め`, timing: '決算後', reason: `含み損${h.plRate}、業績確認後に判断` });
      }
    }
  }

  return {
    portfolioComment,
    benchComment,
    okachiComment,
    marginComment,
    perComment,
    creditComment,
    investorComment,
    marketOutlook,
    portfolioStrategy,
    actions,
  };
}

// --- HTMLレポート生成 ---
function buildReportHTML(data) {
  const today = new Date().toLocaleDateString('ja-JP', {
    year: 'numeric', month: 'long', day: 'numeric', weekday: 'long',
  });

  const { annual, monthly, per, credit, investors } = data;

  // オーナーコメント読み取り
  const ownerComment = data.ownerComment || '';

  // 分析コメント・戦略・アクションプラン生成
  const analysis = generateAnalysis(data);

  // 最新月と基準月(12月)を取得
  const baseMonth = annual[0]; // 12月
  const latestMonth = annual[annual.length - 1]; // 最新月
  const latestMonthLabel = latestMonth.month;

  // KPI
  const totalAssets = latestMonth.totalAssets;
  const totalAssetsMan = (parseNum(totalAssets) / 10000).toFixed(0);
  const monthlyChange = latestMonth.monthlyChange || '-';
  const annualPL = latestMonth.realizedPL;
  const annualPLMan = (parseNum(annualPL) / 10000).toFixed(1);
  const evalRate = latestMonth.evalRate || '-';

  // 岡地の注目銘柄（損益率が高い順にソート）
  const okachi = monthly.sections.okachi
    .filter(h => h.holdingCost && h.plRate)
    .sort((a, b) => {
      const rateA = parseFloat(String(a.plRate).replace(/[%,]/g, '')) || 0;
      const rateB = parseFloat(String(b.plRate).replace(/[%,]/g, '')) || 0;
      return rateB - rateA;
    });

  // 注目銘柄（上位5 + 含み損銘柄）
  const topOkachi = okachi.slice(0, 5);
  const lossOkachi = okachi.filter(h => {
    const rate = parseFloat(String(h.plRate).replace(/[%,]/g, '')) || 0;
    return rate < -5;
  });
  const featuredOkachi = [...topOkachi, ...lossOkachi.filter(h => !topOkachi.includes(h))];

  // SBI現物
  const sbiCash = monthly.sections.sbiCash
    .filter(h => h.holdingCost && h.plRate)
    .sort((a, b) => {
      const rateA = parseFloat(String(a.plRate).replace(/[%,]/g, '')) || 0;
      const rateB = parseFloat(String(b.plRate).replace(/[%,]/g, '')) || 0;
      return rateB - rateA;
    });

  // SBI信用
  const sbiMargin = monthly.sections.sbiMargin
    .filter(h => h.holdingCost && h.plRate);

  // 楽天
  const rakuten = monthly.sections.rakuten
    .filter(h => h.holdingCost && h.plRate);

  const holdingRow = (h) => {
    const rate = String(h.plRate);
    const cls = rate.includes('-') || rate.includes('▲') ? 'minus' : 'plus';
    return `<tr>
      <td style="text-align:left;">${h.name}</td>
      <td>${h.shares}</td>
      <td>${h.price}</td>
      <td>${h.currentPrice}</td>
      <td class="${cls}">${rate}</td>
    </tr>`;
  };

  const signedCell = (val) => {
    const s = String(val || '-');
    const cls = s.includes('-') || s.includes('▲') || s.includes('▼') ? 'minus' : 'plus';
    return `<td class="${cls}">${s}</td>`;
  };

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  body { font-family: 'Hiragino Sans','Yu Gothic',sans-serif; max-width: 700px; margin: 0 auto; padding: 20px; color: #1a1a1a; font-size: 13px; line-height: 1.5; }
  h1 { font-size: 20px; text-align: center; border-bottom: 3px solid #1a3a6b; padding-bottom: 8px; color: #1a3a6b; }
  .date { text-align: center; color: #666; margin-bottom: 20px; font-size: 12px; }
  h2 { font-size: 15px; color: #fff; background: #1a3a6b; padding: 5px 10px; margin: 20px 0 8px; border-radius: 3px; }
  h3 { font-size: 13px; color: #1a3a6b; border-left: 4px solid #1a3a6b; padding-left: 8px; margin: 14px 0 6px; }
  table { border-collapse: collapse; width: 100%; margin: 6px 0 14px; font-size: 12px; }
  th, td { border: 1px solid #ccc; padding: 4px 8px; text-align: right; }
  th { background: #e8eef6; color: #1a3a6b; font-weight: bold; text-align: center; }
  td:first-child { text-align: left; }
  .plus { color: #006600; font-weight: bold; }
  .minus { color: #cc0000; font-weight: bold; }
  .highlight { background: #fffde0; }
  .summary-box { background: #f0f4f8; border: 1px solid #b0c4de; border-radius: 5px; padding: 10px 14px; margin: 10px 0; font-size: 12px; }
  .summary-box p { margin: 4px 0; }
  .kpi-grid { display: flex; gap: 10px; margin: 12px 0; }
  .kpi { flex: 1; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 10px; text-align: center; }
  .kpi .label { font-size: 10px; color: #666; }
  .kpi .value { font-size: 18px; font-weight: bold; color: #1a3a6b; }
  .kpi .sub { font-size: 11px; }
  ul { margin: 4px 0; padding-left: 18px; }
  li { margin: 2px 0; }
  .strategy-item { background: #f8faf8; border-left: 3px solid #2d8f2d; padding: 8px 12px; margin: 8px 0; font-size: 12px; }
  .risk-item { background: #fdf5f5; border-left: 3px solid #cc3333; padding: 8px 12px; margin: 8px 0; font-size: 12px; }
  .strategy-item p, .risk-item p { margin: 3px 0; }
  .footer { color: #999; font-size: 10px; margin-top: 30px; border-top: 1px solid #ddd; padding-top: 10px; }
  .owner-comment { background: #fff8e1; border: 2px solid #f9a825; border-radius: 5px; padding: 12px 16px; margin: 12px 0 16px; font-size: 13px; line-height: 1.6; }
  .owner-comment .label { font-size: 11px; color: #f57f17; font-weight: bold; margin-bottom: 4px; }
</style></head>
<body>

<h1>投資評価レポート・今後の投資戦略</h1>
<p class="date">${today}作成</p>

${ownerComment ? `<div class="owner-comment">
  <div class="label">オーナーの見解・方針</div>
  <p>${ownerComment.replace(/\n/g, '<br>')}</p>
</div>` : ''}

<!-- KPI -->
<div class="kpi-grid">
  <div class="kpi">
    <div class="label">総資産</div>
    <div class="value">${totalAssetsMan}万円</div>
    <div class="sub">前月比 <span class="${pctClass(monthlyChange)}">${fmtPct(monthlyChange)}</span></div>
  </div>
  <div class="kpi">
    <div class="label">年間損益（実現）</div>
    <div class="value" style="color:#006600">+${annualPLMan}万円</div>
    <div class="sub">売却益累計</div>
  </div>
  <div class="kpi">
    <div class="label">評価損益率</div>
    <div class="value">${fmtPct(evalRate)}</div>
    <div class="sub">取得原価比</div>
  </div>
</div>

<h2>1. ポートフォリオ全体の評価</h2>

<h3>1-1. 月次パフォーマンス推移</h3>
<table>
  <tr><th>月</th><th>資産合計</th><th>取得原価</th><th>評価益</th><th>評価率</th><th>月間増減率</th><th>年間損益</th><th>売却損益</th></tr>
  ${annual.map((m, i) => `<tr${i > 0 && i < annual.length - 1 ? ' class="highlight"' : ''}>
    <td>${m.month}${i === 0 ? '(基準)' : ''}</td>
    <td>${m.totalAssets || '-'}</td>
    <td>${m.costBasis || '-'}</td>
    <td>${m.unrealizedPL || '-'}</td>
    <td>${fmtPct(m.evalRate)}</td>
    ${signedCell(m.monthlyChange)}
    <td>${fmtPct(m.annualPL)}</td>
    <td>${m.realizedPL || '-'}</td>
  </tr>`).join('')}
</table>

<div class="summary-box">
  <p><strong>評価:</strong> ${analysis.portfolioComment}</p>
</div>

<h3>1-2. ベンチマーク比較</h3>
<table>
  <tr><th>指標</th><th>12月末</th><th>${latestMonthLabel}時点</th><th>増減率</th></tr>
  <tr><td>ポートフォリオ</td><td>${baseMonth.totalAssets}</td><td>${latestMonth.totalAssets}</td>
    ${signedCell(latestMonth.annualPL)}</tr>
  <tr><td>日経平均</td><td>${baseMonth.nikkei}</td><td>${latestMonth.nikkei}</td>
    ${signedCell(latestMonth.nikkeiAnnual)}</tr>
  <tr><td>TOPIX</td><td>${baseMonth.topix}</td><td>${latestMonth.topix}</td>
    ${signedCell(latestMonth.topixAnnual)}</tr>
  <tr><td>S&P500</td><td>${baseMonth.sp500}</td><td>${latestMonth.sp500}</td>
    ${signedCell(latestMonth.sp500Annual)}</tr>
</table>

<div class="summary-box">
  <p><strong>評価:</strong> ${analysis.benchComment}</p>
</div>

<h2>2. 証券口座別分析</h2>

<h3>2-1. 岡地証券（中長期コア）</h3>
${monthly.summaries.okachi ? `<table>
  <tr><th>分類</th><th>原価</th><th>時価</th><th>評価損益</th><th>損益率</th></tr>
  <tr><td>合計（${latestMonthLabel}）</td>
    <td>${monthly.summaries.okachi.costBasis}</td>
    <td>${monthly.summaries.okachi.marketValue}</td>
    ${signedCell(monthly.summaries.okachi.unrealizedPL)}
    ${signedCell(monthly.summaries.okachi.evalRate)}
  </tr>
</table>` : ''}
<p><strong>注目保有銘柄:</strong></p>
<table>
  <tr><th>銘柄</th><th>株数</th><th>取得価格</th><th>現在値</th><th>損益率</th></tr>
  ${featuredOkachi.map(holdingRow).join('')}
</table>

${analysis.okachiComment ? `<div class="summary-box">
  <p><strong>評価:</strong> ${analysis.okachiComment}</p>
</div>` : ''}

<h3>2-2. SBI証券（現物）</h3>
${monthly.summaries.sbiCash ? `<table>
  <tr><th>区分</th><th>原価</th><th>時価</th><th>評価損益</th><th>売買損益</th></tr>
  <tr><td>現物</td>
    <td>${monthly.summaries.sbiCash.costBasis}</td>
    <td>${monthly.summaries.sbiCash.marketValue}</td>
    ${signedCell(monthly.summaries.sbiCash.unrealizedPL + ' (' + monthly.summaries.sbiCash.evalRate + ')')}
    ${signedCell(monthly.summaries.sbiCash.realizedPL)}
  </tr>
</table>` : ''}
${sbiCash.length > 0 ? `<p><strong>主要保有銘柄:</strong></p>
<table>
  <tr><th>銘柄</th><th>株数</th><th>取得価格</th><th>現在値</th><th>損益率</th></tr>
  ${sbiCash.slice(0, 5).map(holdingRow).join('')}
  ${sbiCash.filter(h => { const r = parseFloat(String(h.plRate).replace(/[%,]/g,''))||0; return r < -10; })
    .filter(h => !sbiCash.slice(0,5).includes(h)).map(holdingRow).join('')}
</table>` : ''}

<h3>2-3. SBI証券（信用取引）</h3>
${monthly.summaries.sbiMargin ? `<table>
  <tr><th>区分</th><th>原価</th><th>時価</th><th>評価損益</th><th>売買損益</th></tr>
  <tr><td>信用</td>
    <td>${monthly.summaries.sbiMargin.costBasis}</td>
    <td>${monthly.summaries.sbiMargin.marketValue}</td>
    ${signedCell(monthly.summaries.sbiMargin.unrealizedPL + ' (' + monthly.summaries.sbiMargin.evalRate + ')')}
    ${signedCell(monthly.summaries.sbiMargin.realizedPL)}
  </tr>
</table>` : ''}

${analysis.marginComment ? `<div class="summary-box">
  <p><strong>評価:</strong> ${analysis.marginComment}</p>
</div>` : ''}

<h3>2-4. 楽天証券</h3>
${monthly.summaries.rakuten ? `<table>
  <tr><th>区分</th><th>原価</th><th>時価</th><th>評価損益率</th></tr>
  <tr><td>${latestMonthLabel}</td>
    <td>${monthly.summaries.rakuten.costBasis}</td>
    <td>${monthly.summaries.rakuten.marketValue}</td>
    ${signedCell(monthly.summaries.rakuten.evalRate)}
  </tr>
</table>` : ''}

<h2>3. 市場環境分析</h2>

<h3>3-1. 日経平均PER・EPS推移</h3>
<table>
  <tr><th>日付</th><th>日経平均</th><th>PER</th><th>EPS</th><th>BPS</th><th>PBR</th></tr>
  ${per.slice(0, 5).map(p => `<tr>
    <td>${p.date}</td><td>${p.nikkei}</td><td>${p.per}</td>
    <td>${p.eps}</td><td>${p.bps}</td><td>${p.pbr}</td>
  </tr>`).join('')}
</table>

${analysis.perComment ? `<div class="summary-box">
  <p><strong>分析:</strong> ${analysis.perComment}</p>
</div>` : ''}

<h3>3-2. 信用取引指標</h3>
<table>
  <tr><th>日付</th><th>売り残</th><th>買い残</th><th>信用倍率</th><th>信用評価率</th></tr>
  ${credit.slice(0, 5).map(c => `<tr>
    <td>${c.date}</td><td>${c.sellBalance}</td><td>${c.buyBalance}</td>
    <td>${c.creditRatio}</td><td>${c.evalRate}</td>
  </tr>`).join('')}
</table>

${analysis.creditComment ? `<div class="summary-box">
  <p><strong>分析:</strong> ${analysis.creditComment}</p>
</div>` : ''}

<h3>3-3. 投資主体別動向</h3>
<table>
  <tr><th>日付</th><th>海外</th><th>個人</th><th>個人(現金)</th><th>個人(信用)</th><th>投資信託</th><th>信託銀行</th><th>証券自己</th><th>事業法人</th></tr>
  ${investors.slice(0, 4).map(inv => `<tr>
    <td>${inv.date}</td>
    ${signedCell(inv.foreign)}${signedCell(inv.individual)}
    ${signedCell(inv.individualCash)}${signedCell(inv.individualCredit)}
    ${signedCell(inv.investTrust)}${signedCell(inv.trustBank)}
    ${signedCell(inv.securities)}${signedCell(inv.corporate)}
  </tr>`).join('')}
</table>

${analysis.investorComment ? `<div class="summary-box">
  <p><strong>分析:</strong> ${analysis.investorComment}</p>
</div>` : ''}

<h2>4. 今後の投資戦略</h2>

<h3>4-1. 市場見通し</h3>
<div class="summary-box">
  ${analysis.marketOutlook}
</div>

<h3>4-2. ポートフォリオ戦略</h3>
${analysis.portfolioStrategy}

<h2>5. アクションプラン（優先度順）</h2>

${analysis.actions.length > 0 ? `<table>
  <tr><th>#</th><th>アクション</th><th>時期</th><th>理由</th></tr>
  ${analysis.actions.map(a => `<tr>
    <td>${a.priority}</td>
    <td>${a.action}</td>
    <td>${a.timing}</td>
    <td>${a.reason}</td>
  </tr>`).join('')}
</table>` : '<p>現時点で緊急のアクション項目はありません。</p>'}

<h2>6. 総合サマリー</h2>

${monthly.summaries.total ? `<table>
  <tr><th>項目</th><th>金額</th></tr>
  <tr><td>有価証券原価（合計）</td><td>${monthly.summaries.total.costBasis}</td></tr>
  <tr><td>時価評価額（合計）</td><td>${monthly.summaries.total.marketValue}</td></tr>
  <tr><td>評価損益</td>${signedCell(monthly.summaries.total.unrealizedPL + ' (' + monthly.summaries.total.evalRate + ')')}</tr>
  <tr><td>年間実現損益</td>${signedCell(monthly.summaries.total.realizedPL)}</tr>
  <tr><td>資産合計</td><td><strong>${latestMonth.totalAssets}</strong></td></tr>
</table>` : ''}

<div class="footer">
  <p>このメールは自動生成された週次投資評価レポートです。</p>
  <p>データソース: Googleスプレッドシート（ポートフォリオ）、nikkei225jp.com（市場指標）</p>
  <p>※ 投資判断は自己責任でお願いします。</p>
</div>

</body></html>`;
}

// --- テキスト版（dry-run用） ---
function buildPlainText(data) {
  const { annual, monthly } = data;
  const latest = annual[annual.length - 1];
  let text = '';
  text += '========================================\n';
  text += '  投資評価レポート\n';
  text += `  ${new Date().toLocaleDateString('ja-JP')}\n`;
  text += '========================================\n\n';

  text += '【ポートフォリオ概要】\n';
  text += `  資産合計:     ${latest.totalAssets}\n`;
  text += `  取得原価:     ${latest.costBasis}\n`;
  text += `  評価損益率:   ${latest.evalRate}\n`;
  text += `  前月比:       ${latest.monthlyChange}\n`;
  text += `  年間損益:     ${latest.annualPL}\n`;
  text += `  売却損益:     ${latest.realizedPL}\n\n`;

  text += '【ベンチマーク】\n';
  text += `  日経平均:     ${latest.nikkei}  (${latest.nikkeiAnnual})\n`;
  text += `  TOPIX:        ${latest.topix}  (${latest.topixAnnual})\n`;
  text += `  S&P500:       ${latest.sp500}  (${latest.sp500Annual})\n\n`;

  if (monthly.summaries.okachi) {
    text += '【岡地証券】\n';
    text += `  原価: ${monthly.summaries.okachi.costBasis}  時価: ${monthly.summaries.okachi.marketValue}  損益率: ${monthly.summaries.okachi.evalRate}\n\n`;
  }
  if (monthly.summaries.sbiCash) {
    text += '【SBI証券（現物）】\n';
    text += `  原価: ${monthly.summaries.sbiCash.costBasis}  時価: ${monthly.summaries.sbiCash.marketValue}  損益率: ${monthly.summaries.sbiCash.evalRate}\n\n`;
  }
  if (monthly.summaries.sbiMargin) {
    text += '【SBI証券（信用）】\n';
    text += `  原価: ${monthly.summaries.sbiMargin.costBasis}  時価: ${monthly.summaries.sbiMargin.marketValue}  損益率: ${monthly.summaries.sbiMargin.evalRate}\n\n`;
  }

  text += '【PER推移】\n';
  for (const p of data.per.slice(0, 3)) {
    text += `  ${p.date}  日経: ${p.nikkei}  PER: ${p.per}  EPS: ${p.eps}\n`;
  }

  return text;
}

// --- メール送信 ---
async function sendEmail(config, html, plainText) {
  const transporter = createTransport({
    service: 'gmail',
    auth: {
      user: config.gmail_user,
      pass: config.gmail_app_password,
    },
  });

  const info = await transporter.sendMail({
    from: `"Investment Report" <${config.gmail_user}>`,
    to: config.to_email,
    subject: `投資評価レポート - ${new Date().toLocaleDateString('ja-JP')}`,
    text: plainText,
    html: html,
  });

  console.log(`メール送信完了: ${info.messageId}`);
}

// --- メイン ---
async function main() {
  const isDryRun = process.argv.includes('--dry-run');

  console.log('スプレッドシートからデータを取得中...');
  const sheets = getSheets();

  const currentMonthSheet = getCurrentMonthSheetName();
  console.log(`  対象月: ${currentMonthSheet}`);

  // 並列でデータ取得
  const [annualRows, monthlyRows, perRows, creditRows, investorRows, ownerComment] = await Promise.all([
    readSheet(sheets, '年間集計'),
    readSheet(sheets, currentMonthSheet),
    readSheet(sheets, 'PER推移'),
    readSheet(sheets, '信用倍率'),
    readSheet(sheets, '投資主体別'),
    readOwnerComment(sheets, currentMonthSheet),
  ]);

  const data = {
    annual: parseAnnualSummary(annualRows),
    monthly: parseMonthlySheet(monthlyRows),
    per: parsePERData(perRows),
    credit: parseCreditData(creditRows),
    investors: parseInvestorData(investorRows),
    ownerComment,
  };

  console.log(`  年間集計: ${data.annual.length}ヶ月分`);
  console.log(`  保有銘柄: 岡地${data.monthly.sections.okachi.length}件, SBI現物${data.monthly.sections.sbiCash.length}件, SBI信用${data.monthly.sections.sbiMargin.length}件, 楽天${data.monthly.sections.rakuten.length}件`);
  console.log(`  PER推移: ${data.per.length}件, 信用倍率: ${data.credit.length}件, 投資主体別: ${data.investors.length}件`);

  const html = buildReportHTML(data);
  const plainText = buildPlainText(data);

  if (isDryRun) {
    console.log(plainText);
    console.log('\n--- (dry-run: メール送信はスキップ) ---');
    return;
  }

  const emailConfig = loadEmailConfig();
  await sendEmail(emailConfig, html, plainText);
}

main().catch((e) => {
  console.error('エラー:', e.message);
  console.error(e.stack);
  process.exit(1);
});
