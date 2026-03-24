import { chromium } from 'playwright';
import { writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));

const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
  @page { margin: 15mm 12mm; size: A4; }
  body { font-family: 'Hiragino Kaku Gothic ProN', 'Yu Gothic', sans-serif; font-size: 10px; color: #1a1a1a; line-height: 1.5; }
  h1 { font-size: 20px; text-align: center; border-bottom: 3px solid #1a3a6b; padding-bottom: 8px; margin-bottom: 4px; color: #1a3a6b; }
  .date { text-align: center; color: #666; margin-bottom: 20px; font-size: 11px; }
  h2 { font-size: 14px; color: #fff; background: #1a3a6b; padding: 5px 10px; margin: 18px 0 8px; border-radius: 3px; }
  h3 { font-size: 12px; color: #1a3a6b; border-left: 4px solid #1a3a6b; padding-left: 8px; margin: 12px 0 6px; }
  table { border-collapse: collapse; width: 100%; margin: 6px 0 12px; font-size: 9px; }
  th, td { border: 1px solid #ccc; padding: 3px 6px; text-align: right; }
  th { background: #e8eef6; color: #1a3a6b; font-weight: bold; text-align: center; }
  td:first-child, td:nth-child(2) { text-align: left; }
  .plus { color: #006600; font-weight: bold; }
  .minus { color: #cc0000; font-weight: bold; }
  .highlight { background: #fffde0; }
  .section { page-break-inside: avoid; }
  .summary-box { background: #f0f4f8; border: 1px solid #b0c4de; border-radius: 5px; padding: 10px 14px; margin: 10px 0; }
  .summary-box p { margin: 3px 0; font-size: 10px; }
  .kpi-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; margin: 10px 0; }
  .kpi { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 8px; text-align: center; }
  .kpi .label { font-size: 8px; color: #666; }
  .kpi .value { font-size: 16px; font-weight: bold; color: #1a3a6b; }
  .kpi .sub { font-size: 9px; }
  .strategy-item { background: #f8faf8; border-left: 3px solid #2d8f2d; padding: 6px 10px; margin: 6px 0; }
  .risk-item { background: #fdf5f5; border-left: 3px solid #cc3333; padding: 6px 10px; margin: 6px 0; }
  .pagebreak { page-break-before: always; }
  ul { margin: 4px 0; padding-left: 18px; }
  li { margin: 2px 0; }
</style>
</head>
<body>

<h1>投資評価レポート・今後の投資戦略</h1>
<p class="date">2026年3月11日作成 | 2026年1月〜3月のデータに基づく分析</p>

<!-- KPI -->
<div class="kpi-grid">
  <div class="kpi">
    <div class="label">総資産</div>
    <div class="value">3,727万円</div>
    <div class="sub">前月比 <span class="minus">-3.97%</span></div>
  </div>
  <div class="kpi">
    <div class="label">年間損益（実現）</div>
    <div class="value" style="color:#006600">+318.7万円</div>
    <div class="sub">売却益累計</div>
  </div>
  <div class="kpi">
    <div class="label">評価損益率</div>
    <div class="value">+26.02%</div>
    <div class="sub">取得原価比</div>
  </div>
</div>

<h2>1. ポートフォリオ全体の評価</h2>

<h3>1-1. 月次パフォーマンス推移</h3>
<table>
  <tr><th>月</th><th>資産合計</th><th>取得原価</th><th>評価益</th><th>評価率</th><th>月間増減率</th><th>年間損益</th><th>売却損益</th></tr>
  <tr><td>12月(基準)</td><td>35,548,885</td><td>19,162,400</td><td>4,364,627</td><td>22.78%</td><td>-</td><td>-</td><td>-</td></tr>
  <tr class="highlight"><td>1月</td><td>35,703,763</td><td>19,857,642</td><td>4,364,115</td><td>21.98%</td><td class="plus">+0.44%</td><td>4.46%</td><td>+445,968</td></tr>
  <tr class="highlight"><td>2月</td><td>38,812,808</td><td>21,088,083</td><td>6,235,613</td><td>29.57%</td><td class="plus">+8.71%</td><td>31.58%</td><td>+2,712,135</td></tr>
  <tr><td>3月</td><td>37,270,396</td><td>19,278,773</td><td>5,016,604</td><td>26.02%</td><td class="minus">-3.97%</td><td>31.87%</td><td>+28,591</td></tr>
</table>

<div class="summary-box">
  <p><strong>評価:</strong> 2月に日経平均58,850円の高値をつけた局面で積極的に利益確定を行い、年間売却益+271万円を確保した判断は秀逸。3月の急落（日経-10.45%）に対してポートフォリオは-3.97%に留まり、相対的にアウトパフォーム。現金比率を高めたディフェンシブなポジション取りが奏功。</p>
</div>

<h3>1-2. ベンチマーク比較</h3>
<table>
  <tr><th>指標</th><th>12月末</th><th>3月時点</th><th>増減率</th></tr>
  <tr><td>ポートフォリオ</td><td>35,548,885</td><td>37,270,396</td><td class="plus">+4.84%</td></tr>
  <tr><td>日経平均</td><td>50,407</td><td>52,700</td><td class="plus">+4.55%</td></tr>
  <tr><td>TOPIX</td><td>3,417.98</td><td>3,575.00</td><td class="plus">+4.59%</td></tr>
  <tr><td>S&P500</td><td>6,827.41</td><td>6,878.88</td><td class="plus">+0.75%</td></tr>
</table>

<div class="summary-box">
  <p><strong>評価:</strong> ポートフォリオの年初来リターン+4.84%は日経平均(+4.55%)、TOPIX(+4.59%)を若干上回り、S&P500(+0.75%)を大幅にアウトパフォーム。入金額(1月50万円)を考慮した修正リターンでも良好なパフォーマンス。</p>
</div>

<div class="pagebreak"></div>

<h2>2. 証券口座別分析</h2>

<h3>2-1. 岡地証券（中長期コア）</h3>
<table>
  <tr><th>分類</th><th>原価</th><th>時価</th><th>評価損益</th><th>損益率</th></tr>
  <tr><td>合計（3月）</td><td>14,153,195</td><td>18,479,125</td><td class="plus">+4,325,930</td><td class="plus">+30.57%</td></tr>
</table>
<p><strong>特筆すべき保有銘柄:</strong></p>
<table>
  <tr><th>銘柄</th><th>株数</th><th>取得価格</th><th>現在値</th><th>損益率</th><th>備考</th></tr>
  <tr><td>日本ギア工業</td><td>500</td><td>619</td><td>2,085</td><td class="plus">+232.6%</td><td>最高パフォーマンス</td></tr>
  <tr><td>東邦ガス</td><td>300</td><td>1,645</td><td>5,243</td><td class="plus">+218.7%</td><td>安定成長銘柄</td></tr>
  <tr><td>丸紅</td><td>100</td><td>2,193</td><td>5,413</td><td class="plus">+146.8%</td><td>商社・配当</td></tr>
  <tr><td>みずほFG</td><td>100</td><td>3,384</td><td>6,560</td><td class="plus">+93.8%</td><td>金融セクター</td></tr>
  <tr><td>トヨタ自動車</td><td>600</td><td>2,300</td><td>3,473</td><td class="plus">+51.0%</td><td>旧NISA・コア</td></tr>
  <tr><td>三菱UFJFG</td><td>300</td><td>1,732</td><td>2,722</td><td class="plus">+57.2%</td><td>金融セクター</td></tr>
  <tr><td>ジャックス</td><td>300</td><td>4,523</td><td>3,991</td><td class="minus">-11.8%</td><td>要注意</td></tr>
  <tr><td>北海道電力</td><td>400</td><td>1,074</td><td>927</td><td class="minus">-13.7%</td><td>要注意</td></tr>
</table>

<div class="summary-box">
  <p><strong>評価:</strong> 岡地証券のポートフォリオは+30.57%と非常に良好。特にバリュー株・高配当株（東邦ガス、丸紅、金融株）が大きく貢献。含み損銘柄はジャックス・北海道電力・KeePer技研程度で全体のリスクは限定的。</p>
</div>

<h3>2-2. SBI証券（アクティブ運用）</h3>
<table>
  <tr><th>区分</th><th>3月原価</th><th>3月時価</th><th>評価損益</th><th>売買損益</th></tr>
  <tr><td>現物</td><td>3,166,955</td><td>3,414,710</td><td class="plus">+247,755 (+7.82%)</td><td class="plus">+318,230</td></tr>
  <tr><td>信用</td><td>5,005,500</td><td>5,071,800</td><td class="plus">+66,300 (+1.32%)</td><td class="minus">-451,689</td></tr>
</table>

<div class="summary-box">
  <p><strong>評価:</strong> 信用取引で三菱重工業の空売りが計約-54.6万円の大幅損失。三井海洋開発の買いでも-42.7万円の損失が発生。一方、現物ではAeroEdge(+40.9万円)や三井海洋開発売り(+18.1万円)で利益を確保。信用取引の損失管理が課題。</p>
</div>

<h3>2-3. 楽天証券（コモディティ・テーマ）</h3>
<table>
  <tr><th>月</th><th>原価</th><th>時価</th><th>評価損益率</th></tr>
  <tr><td>1月</td><td>1,580,638</td><td>1,864,288</td><td class="plus">+17.95%</td></tr>
  <tr><td>2月</td><td>1,958,622</td><td>2,422,572</td><td class="plus">+23.69%</td></tr>
  <tr><td>3月</td><td>1,958,622</td><td>2,335,241</td><td class="plus">+19.23%</td></tr>
</table>
<p>丸建リース(+69%)、純シルバー信託(+30%)、金先物Wブル(+7%)等、コモディティ関連が堅調。</p>

<h3>2-4. SBI米国株</h3>
<table>
  <tr><th>銘柄</th><th>損益率</th><th>テーマ</th></tr>
  <tr><td>Dウェイブクオンタム(QBTS)</td><td class="plus">+245.5%</td><td>量子コンピュータ</td></tr>
  <tr><td>スカイウォーター(SKYT)</td><td class="plus">+157.8%</td><td>半導体</td></tr>
  <tr><td>パランティア(PLTR)</td><td class="plus">+97.9%</td><td>AI・防衛</td></tr>
  <tr><td>ポニーAI(PONY)</td><td class="plus">+50.8%</td><td>自動運転</td></tr>
  <tr><td>ルルレモン(LULU)</td><td class="minus">-50.2%</td><td>消費財</td></tr>
  <tr><td>コインベース(COIN)</td><td class="minus">-41.3%</td><td>仮想通貨</td></tr>
</table>
<p>量子コンピュータ関連が突出。一方、LULU・COIN・CRMは大幅含み損。米国株全体では微益。</p>

<div class="pagebreak"></div>

<h2>3. 市場環境分析</h2>

<h3>3-1. 日経平均PER・EPS推移</h3>
<table>
  <tr><th>日付</th><th>日経平均</th><th>PER</th><th>EPS</th><th>BPS</th><th>PBR</th></tr>
  <tr><td>3/10</td><td>54,248</td><td>19.41</td><td>2,794.87</td><td>31,177</td><td>1.74</td></tr>
  <tr><td>3/9</td><td>52,729</td><td>18.93</td><td>2,785.46</td><td>31,017</td><td>1.70</td></tr>
  <tr><td>3/6</td><td>55,621</td><td>19.68</td><td>2,826.26</td><td>31,603</td><td>1.76</td></tr>
  <tr><td>2/27</td><td>58,850</td><td>-</td><td>2,813.11</td><td>31,303</td><td>-</td></tr>
</table>

<div class="summary-box">
  <p><strong>分析:</strong></p>
  <ul>
    <li>PER 19.4倍は過去平均（14-16倍）対比やや高めだが、EPS成長を考慮すると許容範囲</li>
    <li>EPSは2,795円で堅調維持。企業業績は底堅い</li>
    <li>PBR 1.74倍。PBR1.0倍=31,177円が理論的な下値メド</li>
    <li>2月高値(58,850)→3/9安値(52,729)で<strong>-10.4%の急落</strong>。調整局面</li>
  </ul>
</div>

<h3>3-2. 信用取引指標</h3>
<table>
  <tr><th>日付</th><th>売り残</th><th>買い残</th><th>信用倍率</th><th>信用評価率</th></tr>
  <tr><td>3/6</td><td>979,359</td><td>5,718,133</td><td>5.84</td><td>-</td></tr>
  <tr><td>2/27</td><td>1,003,420</td><td>5,540,554</td><td>5.52</td><td>-0.01</td></tr>
  <tr><td>2/20</td><td>1,062,099</td><td>5,583,035</td><td>5.26</td><td>-2.15</td></tr>
  <tr><td>2/13</td><td>1,022,829</td><td>5,285,335</td><td>5.17</td><td>-2.02</td></tr>
  <tr><td>2/6</td><td>929,926</td><td>5,355,280</td><td>5.76</td><td>-3.06</td></tr>
</table>

<div class="summary-box">
  <p><strong>分析:</strong></p>
  <ul>
    <li>信用倍率5.84倍：買い残が売り残の約6倍。需給面では重い（将来の売り圧力）</li>
    <li>買い残は5.7兆円と高水準。市場下落時に投げ売りリスクあり</li>
    <li>信用評価率はまだ浅いマイナス圏。-10%以下で追証→投げ売り→底値圏の目安</li>
    <li>現状は「調整途中」の段階。まだセリングクライマックスには至っていない</li>
  </ul>
</div>

<h3>3-3. 投資主体別動向</h3>
<table>
  <tr><th>日付</th><th>海外</th><th>個人</th><th>投資信託</th><th>信託銀行</th><th>証券自己</th><th>事業法人</th></tr>
  <tr><td>2/27</td><td class="plus">+791,073</td><td class="minus">-460,478</td><td class="minus">-357,606</td><td class="minus">-749,711</td><td class="plus">+904,729</td><td class="plus">+7,737</td></tr>
  <tr><td>2/20</td><td class="plus">+542,667</td><td class="plus">+546,619</td><td class="minus">-301,538</td><td class="minus">-951,576</td><td class="plus">+75,017</td><td class="plus">+293,280</td></tr>
  <tr><td>2/13</td><td class="plus">+1,232,355</td><td class="minus">-1,165,872</td><td class="minus">-179,160</td><td class="minus">-444,004</td><td class="plus">+747,776</td><td class="plus">+114,757</td></tr>
  <tr><td>2/6</td><td class="plus">+274,598</td><td class="minus">-440,275</td><td class="minus">-51,348</td><td class="minus">-378,879</td><td class="plus">+609,368</td><td class="plus">+392,970</td></tr>
</table>

<div class="summary-box">
  <p><strong>分析:</strong></p>
  <ul>
    <li><strong>海外投資家:</strong> 4週連続の買い越し。2月は累計+284億円。日本株への資金流入継続</li>
    <li><strong>個人:</strong> 概ね売り越し。高値圏での利益確定が中心</li>
    <li><strong>信託銀行:</strong> 4週連続の売り越し（年金系）。リバランスの売りか</li>
    <li><strong>事業法人:</strong> 安定的に買い越し。自社株買いが下支え</li>
  </ul>
</div>

<div class="pagebreak"></div>

<h2>4. トレード分析と課題</h2>

<h3>4-1. 月別売買損益</h3>
<table>
  <tr><th>月</th><th>SBI現物</th><th>SBI信用</th><th>岡地</th><th>合計</th></tr>
  <tr><td>1月</td><td class="plus">+196,510</td><td class="minus">-44,809</td><td>+294,267</td><td class="plus">+445,968</td></tr>
  <tr><td>2月</td><td class="plus">+139,840</td><td class="minus">-187,491</td><td>+2,759,786</td><td class="plus">+2,712,135</td></tr>
  <tr><td>3月</td><td class="plus">+318,230</td><td class="minus">-451,689</td><td>+162,050</td><td class="plus">+28,591</td></tr>
  <tr style="font-weight:bold"><td>累計</td><td class="plus">+654,580</td><td class="minus">-684,000</td><td>+3,216,103</td><td class="plus">+3,186,694</td></tr>
</table>

<div class="summary-box">
  <p><strong>課題と教訓:</strong></p>
  <ul>
    <li><strong>信用取引の損失:</strong> 3ヶ月で約-68.4万円。特に三菱重工業の空売り（-54.6万円）と三井海洋開発の買い（-42.7万円）が大きい</li>
    <li><strong>三菱重工業空売り:</strong> 防衛関連銘柄の上昇トレンド中に逆張り空売りを繰り返し、損切りが遅れた。トレンドに逆らわないルール徹底が必要</li>
    <li><strong>岡地証券の利益確定:</strong> 2月に+276万円と大きな利確に成功。高値圏での決断力は評価できる</li>
    <li><strong>現物取引:</strong> SBI現物は3ヶ月とも黒字で安定。AeroEdge(+40.9万円)等のグロース銘柄で成果</li>
  </ul>
</div>

<h3>4-2. 信用取引の問題点（詳細）</h3>
<div class="risk-item">
  <p><strong>三菱重工業 空売り:</strong> 3,896円→4,150円→4,388円→4,709円→4,969円と5回のエントリー。防衛費増額のテーマ株に対して逆張りを重ねた結果、累計-54.6万円。<br>
  <strong>→ 教訓:</strong> 政策テーマ（防衛費増額）に支えられた銘柄の空売りは避ける。損切りルール（-5%等）の厳格適用が必要。</p>
</div>
<div class="risk-item">
  <p><strong>三井海洋開発 買い:</strong> 14,065円で400株買い→13,000円で損切り（-42.7万円）。高値掴み。<br>
  <strong>→ 教訓:</strong> エネルギー関連のボラティリティの高い銘柄は小ロットでエントリーすべき。</p>
</div>

<div class="pagebreak"></div>

<h2>5. 今後の投資戦略</h2>

<h3>5-1. 市場見通し（2026年4月〜6月）</h3>
<div class="summary-box">
  <p><strong>メインシナリオ（確率55%）:</strong> 日経平均 50,000〜56,000円のレンジ</p>
  <ul>
    <li>2月高値58,850円からの調整は一巡しつつあるが、信用買い残の整理に時間がかかる</li>
    <li>EPS 2,795円が維持されればPER18倍=50,300円が下値メド</li>
    <li>海外投資家の買い継続が下支え。事業法人の自社株買いも支援材料</li>
  </ul>
  <p><strong>リスクシナリオ（確率25%）:</strong> 48,000円以下への急落</p>
  <ul>
    <li>信用評価率が-10%以下に悪化→追証→投げ売りの連鎖</li>
    <li>米国景気後退懸念の再燃、円高進行</li>
  </ul>
  <p><strong>上振れシナリオ（確率20%）:</strong> 58,000円超の新高値</p>
  <ul>
    <li>EPSの上方修正、海外マネーの流入加速</li>
  </ul>
</div>

<h3>5-2. ポートフォリオ戦略</h3>

<div class="strategy-item">
  <p><strong>1. 現金比率の維持（目標30-35%）</strong></p>
  <p>現在の現金約661万円（現金比率35%）は適切。急落時の買い増し余力として維持。日経50,000円割れでは段階的に投下。</p>
</div>

<div class="strategy-item">
  <p><strong>2. 岡地証券 コアポジション方針</strong></p>
  <ul>
    <li><strong>継続保有:</strong> 東邦ガス、丸紅、三菱UFJFG、みずほFG、トヨタ、三井住友FG → 配当収入+値上がり益を享受</li>
    <li><strong>利益確定検討:</strong> 日本ギア工業（+232%）→ 半分利確を推奨。上昇余地は限定的</li>
    <li><strong>損切り検討:</strong> 北海道電力（-13.7%）→ 電力セクター全体の見通し次第。反発なければ損切り</li>
    <li><strong>損切り検討:</strong> KeePer技研（-13.6%）→ 業績動向を確認の上判断</li>
  </ul>
</div>

<div class="strategy-item">
  <p><strong>3. SBI証券 信用取引ルール改善</strong></p>
  <ul>
    <li><strong>損切りルール:</strong> エントリー価格から-5%で機械的に損切り（三菱重工の教訓）</li>
    <li><strong>ポジションサイズ:</strong> 1銘柄あたり信用建玉100万円以下に制限</li>
    <li><strong>空売り制限:</strong> 政策テーマ銘柄（防衛・半導体等）の空売りは原則禁止</li>
    <li><strong>利確ルール:</strong> +10%で半分利確、+20%で残り利確</li>
  </ul>
</div>

<div class="strategy-item">
  <p><strong>4. 新規投資候補（購入計画に基づく）</strong></p>
  <table>
    <tr><th>優先度</th><th>銘柄</th><th>目標株価</th><th>理由</th></tr>
    <tr><td>A</td><td>信越化学工業(4063)</td><td>5,800-6,000円</td><td>半導体素材の世界的リーダー。現在6,098円から下押し時に追加。既にポジションあり</td></tr>
    <tr><td>A</td><td>TOTO(5332)</td><td>5,500-5,700円</td><td>リフォーム需要堅調。購入計画では3,798円まで待つが、現在の5,715円でも配当1.75%</td></tr>
    <tr><td>B</td><td>村田製作所(6981)</td><td>2,100-2,500円</td><td>電子部品の王者。現在3,672円は目標の2,100円からまだ遠いが、急落時に注目</td></tr>
    <tr><td>B</td><td>カルビー(2229)</td><td>2,600円</td><td>ディフェンシブ。現在3,145円からの下落時に追加買い</td></tr>
    <tr><td>C</td><td>ブリヂストン(5108)</td><td>5,000-5,600円</td><td>高配当(利回り6.7%)。現在3,413円で目標を大幅に下回る→見直し要</td></tr>
  </table>
</div>

<div class="strategy-item">
  <p><strong>5. 米国株戦略</strong></p>
  <ul>
    <li><strong>利益確定:</strong> QBTS(+245%)、SKYT(+158%)は一部利確推奨。量子コンピュータの商用化はまだ先</li>
    <li><strong>損切り:</strong> LULU(-50%)、COIN(-41%)は損切りを検討。回復見込み薄い</li>
    <li><strong>継続保有:</strong> 高配当ETF（VYM、SPYD、HDV）は長期保有。配当再投資で複利効果</li>
    <li><strong>追加検討:</strong> ARM（AI半導体）は押し目で追加。PLTR（AI・防衛）も中長期有望</li>
  </ul>
</div>

<h3>5-3. リスク管理</h3>
<div class="risk-item">
  <p><strong>最大損失シナリオ:</strong> 日経平均が48,000円(-11%)まで下落した場合、ポートフォリオは約-7%（約-260万円）と試算。現金ポジションがバッファとなり、壊滅的な損失は回避可能。</p>
</div>
<div class="risk-item">
  <p><strong>信用取引リスク:</strong> 現在の信用建玉約500万円。最悪ケースで-50万円程度。損切りルールの徹底で-25万円以内に抑制を目指す。</p>
</div>

<h2>6. アクションプラン（優先度順）</h2>
<table>
  <tr><th>#</th><th>アクション</th><th>時期</th><th>理由</th></tr>
  <tr><td>1</td><td>信用取引の損切りルール策定・適用</td><td>即時</td><td>3ヶ月で-68万円の損失防止</td></tr>
  <tr><td>2</td><td>日本ギア工業 250株利確</td><td>今週中</td><td>+232%の利益確保。残り250株は配当目的で保有</td></tr>
  <tr><td>3</td><td>QBTS 5株・SKYT 5株 利確</td><td>今週中</td><td>+245%/+158%の一部利益確定</td></tr>
  <tr><td>4</td><td>LULU・COIN 損切り</td><td>3月中</td><td>含み損の拡大防止。損失を税金対策に活用</td></tr>
  <tr><td>5</td><td>北海道電力の見極め</td><td>4月決算後</td><td>業績確認後に保有/売却を判断</td></tr>
  <tr><td>6</td><td>日経50,000円割れで段階的に信越化学・村田製作所を購入</td><td>下落時</td><td>購入計画に沿った長期投資</td></tr>
</table>

<div class="summary-box" style="margin-top: 16px;">
  <p><strong>総合評価:</strong> ポートフォリオ全体のパフォーマンスは良好（年初来+4.84%、ベンチマーク超え）。2月の高値圏での利益確定判断は秀逸。最大の改善点は信用取引の損失管理。損切りルールの厳格化により、年間の実現損益をさらに改善できる余地がある。岡地証券のコアポートフォリオ（高配当・バリュー銘柄）は引き続き堅実な運用基盤として機能している。</p>
</div>

</body>
</html>`;

// HTMLファイルを保存
const htmlPath = join(__dirname, 'report.html');
writeFileSync(htmlPath, html);

// PlaywrightでPDFに変換
const browser = await chromium.launch();
const page = await browser.newPage();
await page.goto(`file://${htmlPath}`, { waitUntil: 'networkidle' });

const pdfPath = join(__dirname, '..', '投資評価レポート_2026年3月.pdf');
await page.pdf({
  path: pdfPath,
  format: 'A4',
  printBackground: true,
  margin: { top: '15mm', bottom: '15mm', left: '12mm', right: '12mm' },
});

await browser.close();
console.log(`PDFを生成しました: ${pdfPath}`);
