import XLSX from "xlsx";
import { google } from "googleapis";
import { readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const CRED_PATH = "/Users/k.kato/Claude/investment/investment-claude-e261a94c1e98.json";

const wb = XLSX.readFile("/tmp/kessan.xlsx");
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

function serialToDate(serial) {
  if (!serial || typeof serial !== "number") return null;
  const d = new Date((serial - 25569) * 86400000);
  const m = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  if (isNaN(m) || isNaN(day)) return null;
  return m + "/" + day;
}

// Build code -> date map
const codeToDate = {};
for (let i = 5; i < data.length; i++) {
  const row = data[i];
  if (!row || !row[2]) continue;
  const d = serialToDate(row[0]);
  if (d) codeToDate[String(row[1])] = d;
}

// Company name -> stock code mapping
const codeMap = {
  "デンソー":"6902","ヒューリック":"3003","キムラユニティー":"9368",
  "東邦ガス":"9533","商船三井":"9104","三井物産":"8031","双日":"2768",
  "エラン":"6099","DMG森精機":"6141","日本郵船":"9101",
  "川崎重工":"7012","帝人":"3401","日清紡":"3105","カルビー":"2229",
  "ピジョン":"7956","高砂熱学":"1969","ゴールドウィン":"8111",
  "住友商事":"8053","伊藤忠":"8001","三菱商事":"8058",
  "矢作建設":"1870","中部飼料":"2053","兼松":"8020",
  "日清食品":"2897","湖池屋":"2226","トヨタ自動車":"7203",
  "日本ハム":"2282","三菱重工":"7011","シノブフーズ":"2903",
  "共和コーポ":"6570","日本ギア":"6356","石油資源開発":"1662",
  "大和ハウス":"1925","ヤクルト本社":"2267","安藤ハザマ":"1719",
  "横浜ゴム":"5101","ブリヂストン":"5108",
  "アステラス製薬":"4503","日立建機":"6305","アドバンテスト":"6857",
  "野村HD":"8604","スカパー":"9412","キーエンス":"6861","信越化学":"4063",
  "日立製作所":"6501","ゲンキー":"9267","積水化学":"4204",
  "小松製作所":"6301","TOTO":"5332","ソシオネクスト":"6526",
  "さくらインター":"3778","中部電力":"9502","豊田通商":"8015",
  "レーザーテック":"6920","OLC":"4661","東京メトロ":"9023",
  "ZOZO":"3092","伊勢化学":"4107","日本特殊陶業":"5334",
  "JR東海":"9022","牧野フライス":"6135","村田製作所":"6981",
  "JR東":"9020","東京エレク":"8035","LIXIL":"5938",
  "カゴメ":"2811","センチュリー21":"8898","東京電力":"9501",
  "MonotaRO":"3064","GMOフィナンシャル":"7177",
  "エムスリー":"2413","JR西":"9021","日本航空":"9201","丸紅":"8002",
  "メルカリ":"4385","HENNGE":"4475","川崎汽船":"9107",
  "JT":"2914","SBIレオスひふみ":"165A",
  "ライオン":"4912","地主":"3252","SUMCO":"3436","花王":"4452",
  "BASE":"4477","味の素":"2802","武田薬品":"4502",
  "バンダイ":"7832","KADOKAWA":"9468","沖縄セルラー":"9436",
  "IHI":"7013","全国保証":"7164","オムロン":"6645",
  "ダイキン":"6367","荏原実業":"6328","クラウドワークス":"3900","任天堂":"7974",
  "芙蓉総合リース":"8424","日本ヒューム":"5262","丸紅リース":"197A",
  "稲畑産業":"8098","スターゼン":"8043","イエローハット":"9882",
  "クボタ":"6326","太陽誘電":"6976","ニチコン":"6996",
  "三井不動産":"8801","東京建物":"8804","アップルインター":"2788",
  "セリア":"2782","ショーボンドHD":"1414","鴻池運輸":"9025",
  "十六FG":"7380","高速":"7504","ラウンドワン":"4680",
  "リクルート":"6098","日本製鉄":"5401","DeNA":"2432",
  "TOWA":"6315","IBJ":"6071","SCREEN":"7735",
  "浜松ホトニクス":"6965","住信SBI":"7163","Jパワー":"9513",
  "ホンダ":"7267","JX金属":"5016",
  "イーグランド":"3294","UBE":"4208","守谷輸送":"6226",
  "資生堂":"4911","芝浦機械":"6104","日本酸素":"4091","スズキ":"7269",
  "澁澤倉庫":"9304","ワークマン":"7564","オリックス":"8591",
  "理経":"8226","三菱地所":"8802","神戸製鋼":"5406","楽天銀行":"5838",
  "LaboroAI":"5586","古河電気":"5801","フジクラ":"5803",
  "ロート製薬":"4527","ニトリ":"9843","エア・ウォーター":"4088",
  "GSユアサ":"6674","ソフトバンクG":"9984","丸井G":"8252",
  "寿スピリッツ":"2222","ローム":"6963","カバー":"5253",
  "太平洋セメント":"5233","タカラトミー":"7867","ゼンショー":"7550",
  "日通":"9147","亀田製菓":"2220","TOYOTIRE":"5105",
  "名村造船":"7014","鉄建建設":"1815","KOKUSAI":"6525",
  "三井E&S":"7003","ダイフク":"6383","東ソー":"4042",
  "ゴルフ・ドゥ":"3319","タスキHD":"166A",
  "アズパートナーズ":"160A","宮地エンジ":"3431","INFORICH":"9338",
  "TOPPAN":"7911","スクエニ":"9684","東京センチュリー":"8439",
  "メタプラネット":"3350","京セラ":"6971","三井住友FG":"8316",
  "横河ブリッジ":"5911","芝浦メカ":"6590","日揮HD":"1963",
  "楽天G":"4755","MCJ":"6670","KDDI":"9433","ムゲンエステート":"3299",
  "KIYOラーニング":"7353","揚羽":"9330","NTN":"6472","サンリオ":"8136",
  "ソニーG":"6758",
  "フルッタフルッタ":"2586","荏原製作所":"6361","全保連":"5845",
  "すかいらーく":"3197","エアトリ":"6191","AeroEdge":"7409",
  "カウリス":"153A","インバウンドプラット":"5587","エアークローゼット":"9557",
  "共立メンテ":"9616","三菱HC":"8593","サイバー":"4751",
  "KeePer":"6036","フェローテック":"6890","フリー":"4478",
  "野村マイクロ":"6254","みずほFG":"8411","ジャックス":"8584",
  "Abalance":"3856","サイバーセキュリティ":"4493","三菱UFJFG":"8306",
  "キオクシア":"285A",
  "東映アニメ":"4816","カナデン":"8081",
  "東京海上HD":"8766","SOMPO":"8630","MS&AD":"8725",
  "SBIHD":"8473","AGC":"5201","ENEOS":"5020","INPEX":"1605","NTT":"9432",
};

// Resolve dates
const companyDates = {};
const notFound = [];
for (const [name, code] of Object.entries(codeMap)) {
  const d = codeToDate[code];
  if (d) companyDates[name] = d;
  else notFound.push(name + "(" + code + ")");
}
if (notFound.length) console.log("JPXに未掲載:", notFound.join(", "));

// 2026 business days
const businessDays = ["4/24","4/27","4/28","4/30","5/1","5/7","5/8","5/11","5/12","5/13","5/14","5/15","5/18","5/19","5/20"];
const byDate = {};
for (const d of businessDays) byDate[d] = { close: [], own: [] };

// Original close column companies
const origClose = [
  "アステラス製薬","日立建機","アドバンテスト","野村HD","スカパー","キーエンス","信越化学",
  "日立製作所","ゲンキー","積水化学","小松製作所","TOTO","ソシオネクスト",
  "さくらインター","中部電力","豊田通商","レーザーテック","OLC","東京メトロ",
  "ZOZO","伊勢化学","日本特殊陶業","JR東海","牧野フライス","村田製作所",
  "JR東","東京エレク","LIXIL","カゴメ","センチュリー21","東京電力",
  "MonotaRO","GMOフィナンシャル",
  "エムスリー","JR西","日本航空","丸紅",
  "メルカリ","HENNGE","川崎汽船","JT","SBIレオスひふみ",
  "ライオン","地主","SUMCO","花王","BASE","味の素","武田薬品","バンダイ","KADOKAWA",
  "沖縄セルラー","IHI","全国保証","オムロン","ダイキン","荏原実業","クラウドワークス","任天堂",
  "芙蓉総合リース","日本ヒューム","丸紅リース","稲畑産業","スターゼン","イエローハット",
  "クボタ","太陽誘電","ニチコン","三井不動産","東京建物","アップルインター","セリア",
  "ショーボンドHD","鴻池運輸","十六FG","高速","ラウンドワン","リクルート","日本製鉄",
  "DeNA","TOWA","IBJ","SCREEN","浜松ホトニクス","住信SBI","Jパワー","ホンダ","JX金属",
  "イーグランド","UBE","守谷輸送","資生堂","芝浦機械","日本酸素","スズキ","澁澤倉庫",
  "ワークマン","オリックス","理経","三菱地所","神戸製鋼","楽天銀行","LaboroAI",
  "古河電気","フジクラ","ロート製薬","ニトリ","エア・ウォーター","GSユアサ","ソフトバンクG",
  "丸井G","寿スピリッツ","ローム","カバー","太平洋セメント","タカラトミー","ゼンショー",
  "日通","亀田製菓","TOYOTIRE","名村造船","鉄建建設","KOKUSAI","三井E&S","ダイフク",
  "東ソー","ゴルフ・ドゥ","タスキHD",
  "アズパートナーズ","宮地エンジ","INFORICH","TOPPAN","スクエニ","東京センチュリー",
  "メタプラネット","京セラ","三井住友FG","横河ブリッジ","芝浦メカ","日揮HD","楽天G",
  "MCJ","KDDI","ムゲンエステート","KIYOラーニング","揚羽","NTN","サンリオ","ソニーG",
  "フルッタフルッタ","荏原製作所","全保連","すかいらーく","エアトリ","AeroEdge","カウリス",
  "インバウンドプラット","エアークローゼット","共立メンテ","三菱HC","サイバー","KeePer",
  "フェローテック","フリー","野村マイクロ","みずほFG","ジャックス","Abalance",
  "サイバーセキュリティ","三菱UFJFG","キオクシア",
  "東映アニメ","カナデン",
  "東京海上HD","SOMPO","MS&AD",
];

// Original AM/PM companies (put in close since times unknown for 2026)
const origAmPm = [
  "デンソー","ヒューリック","キムラユニティー","東邦ガス","商船三井",
  "三井物産","双日","エラン","DMG森精機","日本郵船",
  "SBIHD","川崎重工","帝人","日清紡","カルビー",
  "ピジョン","高砂熱学","ゴールドウィン",
  "住友商事","伊藤忠","三菱商事","矢作建設","中部飼料",
  "兼松","日清食品","湖池屋","トヨタ自動車",
  "NTT","日本ハム","三菱重工","シノブフーズ",
  "AGC","ENEOS","共和コーポ","日本ギア","石油資源開発","INPEX","大和ハウス","ヤクルト本社",
  "安藤ハザマ","横浜ゴム","ブリヂストン",
];

// Original own companies
const origOwn = [
  "小松製作所","TOTO","レーザーテック","中部電力",
  "東邦ガス","日本特殊陶業","村田製作所",
  "双日","エムスリー","丸紅",
  "トヨタ自動車","NTT","丸紅リース","アップルインター","Jパワー",
  "カルビー","AGC","イーグランド",
  "INPEX","ヤクルト本社","エア・ウォーター","GSユアサ","カバー",
  "楽天G","三井住友FG",
  "すかいらーく","AeroEdge","カウリス","野村マイクロ","三菱UFJFG","みずほFG","ジャックス","Abalance",
  "カナデン",
];

// Place companies
for (const name of [...origClose, ...origAmPm]) {
  const d = companyDates[name];
  if (d && byDate[d]) {
    if (!byDate[d].close.includes(name)) byDate[d].close.push(name);
  }
}
for (const name of origOwn) {
  const d = companyDates[name];
  if (d && byDate[d]) {
    if (!byDate[d].own.includes(name)) byDate[d].own.push(name);
  }
}

// Print summary
console.log("\n=== 2026年 4.5月決算シート ===\n");
for (const d of businessDays) {
  const c = byDate[d].close;
  const o = byDate[d].own;
  console.log(`${d} (${c.length}銘柄): ${c.join("、") || "(なし)"}`);
  if (o.length) console.log(`  所有: ${o.join("、")}`);
}

// Write to spreadsheet
const credentials = JSON.parse(readFileSync(CRED_PATH, "utf8"));
const auth = new google.auth.GoogleAuth({ credentials, scopes: ["https://www.googleapis.com/auth/spreadsheets"] });
const sheets = google.sheets({ version: "v4", auth });
const spreadsheetId = "12hycp-InFw3fGUkyMjdW9lvOWF8mOX6pWfJbEVV2RKk";

const rows = businessDays.map(d => {
  const closeText = byDate[d].close.join("、");
  const ownText = byDate[d].own.join("、");
  return ["", d, "", "", closeText, ownText];
});

await sheets.spreadsheets.values.clear({
  spreadsheetId,
  range: "'4.5月決算'!A2:F50",
});

await sheets.spreadsheets.values.update({
  spreadsheetId,
  range: "'4.5月決算'!A2:F" + (rows.length + 1),
  valueInputOption: "RAW",
  requestBody: { values: rows },
});

console.log("\nスプレッドシート更新完了！");
