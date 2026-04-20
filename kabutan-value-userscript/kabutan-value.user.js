// ==UserScript==
// @name         Kabutan バリュー投資テーブル
// @namespace    local.kkato.kabutan-value
// @version      0.1.0
// @description  株探の決算ページに、バリュー投資シートを再現した自作テーブルを追加表示する
// @match        https://kabutan.jp/stock/finance*
// @run-at       document-end
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const CONTAINER_ID = 'kk-value-table-box';
  const PER_STORAGE_KEY = 'kk-value-target-per';
  const DEFAULT_TARGET_PER = 18;

  const getTargetPer = () => {
    const v = parseFloat(localStorage.getItem(PER_STORAGE_KEY));
    return Number.isFinite(v) && v > 0 ? v : DEFAULT_TARGET_PER;
  };
  const setTargetPer = (v) => {
    localStorage.setItem(PER_STORAGE_KEY, String(v));
  };

  // ---------- ユーティリティ ----------
  const toNum = (s) => {
    if (s == null) return NaN;
    const n = parseFloat(String(s).replace(/[,%円倍\s]/g, ''));
    return Number.isFinite(n) ? n : NaN;
  };

  const fmt = (n, digits = 2) =>
    Number.isFinite(n) ? n.toLocaleString('ja-JP', { maximumFractionDigits: digits, minimumFractionDigits: digits }) : '—';

  const fmtInt = (n) =>
    Number.isFinite(n) ? Math.round(n).toLocaleString('ja-JP') : '—';

  const fmtPct = (r, digits = 2) =>
    Number.isFinite(r) ? (r * 100).toFixed(digits) + '%' : '—';

  // ---------- DOMからスクレイピング ----------
  function scrape() {
    // 現在株価
    const priceText = document.querySelector('.kabuka')?.textContent ?? '';
    const price = toNum(priceText);

    // PER, PBR, 利回り
    const indicatorTds = document.querySelectorAll('#stockinfo_i3 table tbody tr:first-child td');
    const per = toNum(indicatorTds[0]?.textContent);
    const yieldPct = toNum(indicatorTds[2]?.textContent); // %表示の数値

    // 通期業績（result）からEPSと配当の履歴
    const rows = document.querySelectorAll('.fin_year_result_d table tbody tr');
    const yearly = []; // { term: '2025.03', isForecast: bool, eps, dps }
    rows.forEach((tr) => {
      const th = tr.querySelector('th[scope="row"]');
      if (!th) return;
      const thText = th.textContent || '';
      const m = thText.match(/(\d{4}\.\d{2})/);
      if (!m) return; // 「前期比」などは除外
      const tds = tr.querySelectorAll('td');
      if (tds.length < 6) return;
      const eps = toNum(tds[4]?.textContent);
      const dps = toNum(tds[5]?.textContent);
      yearly.push({
        term: m[1],
        isForecast: /予/.test(thText),
        eps,
        dps,
      });
    });

    return { price, per, yieldPct, yearly };
  }

  // ---------- 計算 ----------
  function calc(data) {
    const { price, yearly } = data;
    if (!yearly.length) return null;

    // 今期予想行（予マーク優先。予マークがない、またはEPSが取得できない場合は1年前=直近実績を使用）
    let currentIdx = yearly.findIndex((y) => y.isForecast && Number.isFinite(y.eps));
    if (currentIdx === -1) {
      // 予マーク行がない/EPSなし → 直近実績（yearly末尾から遡ってEPSがある行）
      for (let i = yearly.length - 1; i >= 0; i--) {
        if (!yearly[i].isForecast && Number.isFinite(yearly[i].eps)) {
          currentIdx = i;
          break;
        }
      }
      if (currentIdx === -1) currentIdx = yearly.length - 1;
    }
    const current = yearly[currentIdx];

    // 履歴は今期予想を先頭とし、古い方向へ辿る
    // kabutan表は古い→新しい順なので、current以前を逆順で
    // 今期予想 + 過去10年分までに制限
    const history = yearly.slice(0, currentIdx + 1).reverse().slice(0, 11); // [今期予想, 1年前, ..., 10年前]
    const epsSeries = history.map((h) => h.eps).filter((v) => Number.isFinite(v));

    // 成長率を指定年数分で計算するヘルパー
    // nYears = "N年分" のラベル上の年数。
    // データ点は N+1 個（今期予想 + 1年前 〜 N年前）、前年比は N 個。
    // 年平均成長率は「前年比の合計 ÷ データ点数（N+1）」で計算する（ユーザー指定）。
    // CAGR は (今期予想 / N年前)^(1/N) - 1。
    const calcGrowth = (nYears) => {
      // 実際に使えるデータ点数（N+1 個）
      const desiredPoints = nYears + 1;
      const actualPoints = Math.min(desiredPoints, epsSeries.length);
      const actualN = actualPoints - 1; // 実際の年数（=前年比の個数）
      let cg = NaN;
      let ag = NaN;
      if (actualN >= 1) {
        const epsNow = epsSeries[0];
        const epsOld = epsSeries[actualN];
        if (epsNow > 0 && epsOld > 0) {
          cg = Math.pow(epsNow / epsOld, 1 / actualN) - 1;
        }
        let sum = 0;
        let validCount = 0;
        for (let i = 0; i < actualN; i++) {
          const a = epsSeries[i];
          const b = epsSeries[i + 1];
          if (Number.isFinite(a) && Number.isFinite(b) && b !== 0) {
            sum += a / b - 1;
            validCount++;
          }
        }
        if (validCount > 0) {
          // 分母はデータ点数（= actualN + 1）
          ag = sum / actualPoints;
        }
      }
      // 採用成長率: CAGRと年平均成長率の平均。年平均成長率がマイナスならCAGRを採用
      let gr = NaN;
      if (Number.isFinite(cg) && Number.isFinite(ag)) {
        gr = ag < 0 ? cg : (cg + ag) / 2;
      } else if (Number.isFinite(cg)) {
        gr = cg;
      } else if (Number.isFinite(ag)) {
        gr = ag;
      }
      return { lookback: actualN, cagr: cg, avgGrowth: ag, growth: gr };
    };

    // 10年分（最大）・5年分・3年分
    const growth10 = calcGrowth(10);
    const growth5 = calcGrowth(5);
    const growth3 = calcGrowth(3);
    const { lookback: maxLookback, cagr, avgGrowth, growth } = growth10;

    // 配当性向 = 今期配当 / 今期EPS
    const payoutRatio = Number.isFinite(current.dps) && Number.isFinite(current.eps) && current.eps !== 0
      ? current.dps / current.eps
      : NaN;

    // 配当利回り（現在） = 今期配当 / 現在株価
    const currentDivYield = Number.isFinite(current.dps) && Number.isFinite(price) && price > 0
      ? current.dps / price
      : NaN;

    const TARGET_PER = getTargetPer();
    const MUL_15 = Math.pow(1.15, 5); // 2.011357
    const MUL_12 = Math.pow(1.12, 5); // 1.762342

    // 成長率ごとの将来予測を計算するヘルパー
    const projectFuture = (gr) => {
      const fEps = [current.eps];
      const fDps = [current.dps];
      for (let i = 1; i <= 5; i++) {
        const prev = fEps[i - 1];
        const next = Number.isFinite(prev) && Number.isFinite(gr) ? prev * (1 + gr) : NaN;
        fEps.push(next);
        fDps.push(Number.isFinite(next) && Number.isFinite(payoutRatio) ? next * payoutRatio : NaN);
      }
      const pAt5 = Number.isFinite(fEps[5]) ? fEps[5] * TARGET_PER : NaN;
      const yAt5 = Number.isFinite(fDps[5]) && Number.isFinite(pAt5) && pAt5 > 0 ? fDps[5] / pAt5 : NaN;
      const pGain = Number.isFinite(pAt5) && Number.isFinite(price) ? pAt5 - price : NaN;
      const dSum = fDps.slice(1, 6).reduce((s, v) => s + (Number.isFinite(v) ? v : 0), 0);
      const tGain = Number.isFinite(pGain) ? pGain + dSum : NaN;
      const bp15 = Number.isFinite(tGain) ? tGain / MUL_15 : NaN;
      const bp12 = Number.isFinite(tGain) ? tGain / MUL_12 : NaN;
      return {
        futureEps: fEps,
        futureDps: fDps,
        priceAt5: pAt5,
        yieldAt5: yAt5,
        priceGain: pGain,
        divSum: dSum,
        totalGain: tGain,
        buyPrice15: bp15,
        buyPrice12: bp12,
      };
    };

    // 10年分・5年分・3年分の成長率それぞれで将来を予測
    const proj10 = projectFuture(growth);
    const proj5 = projectFuture(growth5.growth);
    const proj3 = projectFuture(growth3.growth);

    // 投資判断用（従来通り10年ベース）
    const {
      futureEps, futureDps, priceAt5, yieldAt5,
      priceGain, divSum, totalGain, buyPrice15, buyPrice12,
    } = proj10;

    return {
      current,
      history,
      epsSeries,
      lookback: maxLookback,
      cagr,
      avgGrowth,
      growth,
      growth3,
      growth5,
      proj10,
      proj5,
      proj3,
      payoutRatio,
      currentDivYield,
      futureEps,
      futureDps,
      priceAt5,
      yieldAt5,
      priceGain,
      divSum,
      totalGain,
      buyPrice15,
      buyPrice12,
      targetPer: TARGET_PER,
    };
  }

  // ---------- 描画 ----------
  function render(data, result) {
    const { price, per, yieldPct } = data;
    const existing = document.getElementById(CONTAINER_ID);
    if (existing) existing.remove();

    const box = document.createElement('div');
    box.id = CONTAINER_ID;
    box.style.cssText = `
      margin: 12px 0;
      padding: 12px 16px;
      border: 2px solid #2a6496;
      border-radius: 6px;
      background: #f7fbff;
      font-size: 13px;
      line-height: 1.55;
      color: #222;
    `;

    if (!result) {
      box.innerHTML = `<div style="color:#c00;">バリュー投資テーブル: 通期業績データを取得できませんでした。</div>`;
      insertBox(box);
      return;
    }

    const {
      current, history, lookback, cagr, avgGrowth, growth, growth3, growth5,
      proj10, proj5, proj3,
      payoutRatio, currentDivYield,
      futureEps, futureDps,
      priceAt5, yieldAt5, priceGain, divSum, totalGain,
      buyPrice15, buyPrice12, targetPer,
    } = result;

    const tblStyle = `border-collapse:collapse; margin:6px 0;`;
    const thStyle = `background:#e8f0fa; border:1px solid #bcd; padding:3px 8px; font-weight:600;`;
    const tdStyle = `border:1px solid #bcd; padding:3px 8px;`;
    const tdR = `${tdStyle} text-align:right;`;

    // 将来行テーブル（10年分/5年分/3年分の成長率でEPS・株価を並列表示）
    const labels = ['現在', '1年後', '2年後', '3年後', '4年後', '5年後'];
    const futureRows = [];
    for (let i = 0; i < 6; i++) {
      const isLast = i === 5;
      const perCell = i === 0 ? fmt(per, 2) : isLast ? String(targetPer) : '';
      const yieldCell = i === 0 ? fmtPct(currentDivYield) : isLast ? fmtPct(proj10.yieldAt5) : '';
      const price10 = i === 0 ? fmtInt(price) : isLast ? fmtInt(proj10.priceAt5) : '';
      const price5 = i === 0 ? fmtInt(price) : isLast ? fmtInt(proj5.priceAt5) : '';
      const price3 = i === 0 ? fmtInt(price) : isLast ? fmtInt(proj3.priceAt5) : '';
      futureRows.push(`
        <tr>
          <td style="${tdStyle}">${labels[i]}</td>
          <td style="${tdR}">${fmtInt(proj10.futureEps[i])}</td>
          <td style="${tdR}">${fmtInt(proj5.futureEps[i])}</td>
          <td style="${tdR}">${fmtInt(proj3.futureEps[i])}</td>
          <td style="${tdR}">${price10}</td>
          <td style="${tdR}">${price5}</td>
          <td style="${tdR}">${price3}</td>
          <td style="${tdR}">${perCell}</td>
          <td style="${tdR}">${fmtInt(proj10.futureDps[i])}</td>
          <td style="${tdR}">${yieldCell}</td>
        </tr>`);
    }

    // 履歴行（今期予想 + 過去）
    const histRows = history.map((h, i) => {
      const label = i === 0 ? '今期予想' : `${i}年前`;
      return `
        <tr>
          <td style="${tdStyle}">${label}</td>
          <td style="${tdStyle}">${h.term}</td>
          <td style="${tdR}">${fmt(h.eps, 2)}</td>
          <td style="${tdR}">${fmt(h.dps, 2)}</td>
        </tr>`;
    }).join('');

    box.innerHTML = `
      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
        <div style="font-weight:700; font-size:14px; color:#2a6496;">
          バリュー投資テーブル（自作）
        </div>
        <div style="font-size:12px;">
          5年後PER:
          <input id="kk-value-per-input" type="number" min="1" step="0.1" value="${targetPer}"
            style="width:60px; padding:2px 4px; border:1px solid #bcd; border-radius:3px; text-align:right;">
          倍
        </div>
      </div>

      <div style="font-weight:600; margin-top:4px;">投資判断</div>
      <table style="${tblStyle}">
        <thead>
          <tr>
            <th style="${thStyle}"></th>
            <th style="${thStyle}">${lookback}年成長率</th>
            <th style="${thStyle}">${growth5.lookback}年成長率</th>
            <th style="${thStyle}">${growth3.lookback}年成長率</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="${tdStyle}">値上がり益（5年）</td>
            <td style="${tdR}">${fmtInt(proj10.priceGain)} 円</td>
            <td style="${tdR}">${fmtInt(proj5.priceGain)} 円</td>
            <td style="${tdR}">${fmtInt(proj3.priceGain)} 円</td>
          </tr>
          <tr>
            <td style="${tdStyle}">配当合計（5年）</td>
            <td style="${tdR}">${fmtInt(proj10.divSum)} 円</td>
            <td style="${tdR}">${fmtInt(proj5.divSum)} 円</td>
            <td style="${tdR}">${fmtInt(proj3.divSum)} 円</td>
          </tr>
          <tr>
            <td style="${tdStyle}">値上がり + 配当</td>
            <td style="${tdR}">${fmtInt(proj10.totalGain)} 円</td>
            <td style="${tdR}">${fmtInt(proj5.totalGain)} 円</td>
            <td style="${tdR}">${fmtInt(proj3.totalGain)} 円</td>
          </tr>
          <tr>
            <td style="${tdStyle}">年15%を狙う購入価額</td>
            <td style="${tdR} background:#ffe8e8;">${fmtInt(proj10.buyPrice15)} 円</td>
            <td style="${tdR} background:#ffe8e8;">${fmtInt(proj5.buyPrice15)} 円</td>
            <td style="${tdR} background:#ffe8e8;">${fmtInt(proj3.buyPrice15)} 円</td>
          </tr>
          <tr>
            <td style="${tdStyle}">年12%を狙う購入価額</td>
            <td style="${tdR} background:#fff0d8;">${fmtInt(proj10.buyPrice12)} 円</td>
            <td style="${tdR} background:#fff0d8;">${fmtInt(proj5.buyPrice12)} 円</td>
            <td style="${tdR} background:#fff0d8;">${fmtInt(proj3.buyPrice12)} 円</td>
          </tr>
        </tbody>
      </table>

      <div style="font-weight:600; margin-top:10px;">将来予測（PER${targetPer}倍で5年後評価）</div>
      <table style="${tblStyle}">
        <thead>
          <tr>
            <th style="${thStyle}" rowspan="2">年度</th>
            <th style="${thStyle}" colspan="3">EPS</th>
            <th style="${thStyle}" colspan="3">株価</th>
            <th style="${thStyle}" rowspan="2">PER</th>
            <th style="${thStyle}" rowspan="2">配当金</th>
            <th style="${thStyle}" rowspan="2">配当利回り</th>
          </tr>
          <tr>
            <th style="${thStyle}">${lookback}年</th>
            <th style="${thStyle}">${growth5.lookback}年</th>
            <th style="${thStyle}">${growth3.lookback}年</th>
            <th style="${thStyle}">${lookback}年</th>
            <th style="${thStyle}">${growth5.lookback}年</th>
            <th style="${thStyle}">${growth3.lookback}年</th>
          </tr>
        </thead>
        <tbody>${futureRows.join('')}</tbody>
      </table>

      <div style="display:flex; gap:24px; flex-wrap:wrap; align-items:flex-start; margin-top:10px;">
        <div>
          <div>
            <div style="font-weight:600; margin-top:4px;">基本情報（株探から取得）</div>
            <table style="${tblStyle}">
              <tr><td style="${tdStyle}">今期予想EPS</td><td style="${tdR}">${fmt(current.eps, 2)} 円</td></tr>
              <tr><td style="${tdStyle}">今期予想配当</td><td style="${tdR}">${fmt(current.dps, 2)} 円</td></tr>
              <tr><td style="${tdStyle}">配当性向</td><td style="${tdR}">${fmtPct(payoutRatio)}</td></tr>
            </table>
          </div>

          <div>
            <div style="font-weight:600; margin-top:10px;">EPS成長率</div>
            <table style="${tblStyle}">
              <tr>
                <th style="${thStyle}"></th>
                <th style="${thStyle}">${lookback}年分</th>
                <th style="${thStyle}">${growth5.lookback}年分</th>
                <th style="${thStyle}">${growth3.lookback}年分</th>
              </tr>
              <tr>
                <td style="${tdStyle}">CAGR</td>
                <td style="${tdR}">${fmtPct(cagr)}</td>
                <td style="${tdR}">${fmtPct(growth5.cagr)}</td>
                <td style="${tdR}">${fmtPct(growth3.cagr)}</td>
              </tr>
              <tr>
                <td style="${tdStyle}">年平均成長率</td>
                <td style="${tdR}">${fmtPct(avgGrowth)}</td>
                <td style="${tdR}">${fmtPct(growth5.avgGrowth)}</td>
                <td style="${tdR}">${fmtPct(growth3.avgGrowth)}</td>
              </tr>
              <tr>
                <td style="${tdStyle}">採用成長率</td>
                <td style="${tdR} background:#fff6d0;">${fmtPct(growth)}</td>
                <td style="${tdR}">${fmtPct(growth5.growth)}</td>
                <td style="${tdR}">${fmtPct(growth3.growth)}</td>
              </tr>
            </table>
          </div>
        </div>

        <div>
          <div style="font-weight:600; margin-top:4px;">EPS履歴（通期業績から取得）</div>
          <table style="${tblStyle}">
            <thead>
              <tr>
                <th style="${thStyle}">区分</th>
                <th style="${thStyle}">決算期</th>
                <th style="${thStyle}">修正EPS</th>
                <th style="${thStyle}">修正1株配</th>
              </tr>
            </thead>
            <tbody>${histRows}</tbody>
          </table>
        </div>
      </div>

      <div style="font-size:11px; color:#666; margin-top:6px;">
        ※ 配当性向は一定と仮定。5年後株価は上部の PER 入力値を使用。
        EPS成長率は CAGR と年平均の単純平均（年平均がマイナスなら CAGR のみ採用）。
        10年未満のデータしか取得できなかった場合は、取得できた年数で計算。
      </div>
    `;

    insertBox(box);

    // PER入力の変更ハンドラ
    const perInput = box.querySelector('#kk-value-per-input');
    if (perInput) {
      perInput.addEventListener('change', () => {
        const v = parseFloat(perInput.value);
        if (Number.isFinite(v) && v > 0) {
          setTargetPer(v);
          runAndRender();
        }
      });
    }
  }

  function insertBox(box) {
    const financeBox = document.getElementById('finance_box');
    const chartMenu = document.querySelector('.chart_menu');
    if (financeBox && financeBox.parentNode) {
      financeBox.parentNode.insertBefore(box, financeBox);
    } else if (chartMenu && chartMenu.parentNode) {
      chartMenu.parentNode.insertBefore(box, chartMenu.nextSibling);
    } else {
      document.body.prepend(box);
    }
  }

  // ---------- 実行フロー ----------
  function runAndRender() {
    const data = scrape();
    const result = calc(data);
    render(data, result);
  }

  function expandAndRun() {
    // プレミアム会員向けの「過去を表示」ボタンを自動クリック
    const btn = document.getElementById('oc_b1_year_result');
    const tbody = document.querySelector('.fin_year_result_d table tbody');

    if (!btn || !tbody) {
      runAndRender();
      return;
    }

    const initialRowCount = tbody.querySelectorAll('tr').length;

    // DOM変化を監視して、行が増えたら再描画
    const observer = new MutationObserver(() => {
      const now = tbody.querySelectorAll('tr').length;
      if (now > initialRowCount) {
        observer.disconnect();
        runAndRender();
      }
    });
    observer.observe(tbody, { childList: true, subtree: true });

    // 先に現在のデータで描画（展開後に再描画される）
    runAndRender();

    try {
      btn.click();
    } catch (e) {
      // クリック失敗時は現在のデータのまま
    }

    // タイムアウト: 3秒経っても増えなかったら諦める
    setTimeout(() => observer.disconnect(), 3000);
  }

  // ページ読込直後のDOM準備を少し待つ
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', expandAndRun);
  } else {
    expandAndRun();
  }
})();
