import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const repoRoot = path.resolve(".");
const outputDir = path.join(repoRoot, "outputs", "pet-news-template");
const outputPath = path.join(outputDir, "pet-news-google-sheets-template.xlsx");
const previewDir = path.join(outputDir, "previews");

const colors = {
  ink: "#1D2B2F",
  muted: "#65757A",
  line: "#DCE4E6",
  soft: "#DFF4F6",
  brand: "#248C96",
  header: "#124E57",
  high: "#FCE4E4",
  medium: "#FFF0D9",
  low: "#E5F5EE",
  white: "#FFFFFF",
};

const workbook = Workbook.create();

const dashboard = workbook.worksheets.add("dashboard");
const alerts = workbook.worksheets.add("alerts");
const keywords = workbook.worksheets.add("keywords");
const categories = workbook.worksheets.add("categories");
const team = workbook.worksheets.add("team");
const settings = workbook.worksheets.add("settings");
const setup = workbook.worksheets.add("setup");

for (const sheet of [dashboard, alerts, keywords, categories, team, settings, setup]) {
  sheet.showGridLines = false;
}

buildSettings(settings);
buildCategories(categories);
buildTeam(team);
buildKeywords(keywords);
buildAlerts(alerts);
buildDashboard(dashboard);
buildSetup(setup);

await fs.mkdir(previewDir, { recursive: true });
for (const sheetName of ["dashboard", "alerts", "keywords", "setup"]) {
  const ranges = {
    dashboard: "A1:O24",
    alerts: "A1:O12",
    keywords: "A1:I10",
    setup: "A1:F11",
  };
  const preview = await workbook.render({
    sheetName,
    range: ranges[sheetName],
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(previewDir, `${sheetName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

await fs.mkdir(outputDir, { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
process.exit(0);

function buildDashboard(sheet) {
  sheet.freezePanes.freezeRows(3);
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [["ペット業界ニュース ダッシュボード"]];
  sheet.getRange("A1").format = {
    fill: colors.header,
    font: { bold: true, color: colors.white, size: 18 },
    horizontalAlignment: "left",
    verticalAlignment: "middle",
  };
  sheet.getRange("A1:H1").format.rowHeightPx = 42;
  sheet.getRange("A2:H2").merge();
  sheet.getRange("A2").values = [["Googleアラートの取り込み状況と、チーム内の確認ステータスを集計します。"]];
  sheet.getRange("A2").format = {
    fill: colors.header,
    font: { color: colors.white, size: 11 },
    verticalAlignment: "middle",
  };
  sheet.getRange("A2:H2").format.rowHeightPx = 28;

  sheet.getRange("A4:B6").values = [
    ["全件", ""],
    ["未確認", ""],
    ["重要度 高", ""],
  ];
  sheet.getRange("D4:E6").values = [
    ["確認中", ""],
    ["共有済み", ""],
    ["7日以内", ""],
  ];
  sheet.getRange("B4:B6").formulas = [
    ['=COUNTA(alerts!D2:D1000)'],
    ['=COUNTIF(alerts!H2:H1000,"未確認")'],
    ['=COUNTIF(alerts!I2:I1000,"高")'],
  ];
  sheet.getRange("E4:E6").formulas = [
    ['=COUNTIF(alerts!H2:H1000,"確認中")'],
    ['=COUNTIF(alerts!H2:H1000,"共有済み")'],
    ['=COUNTIFS(alerts!B2:B1000,">="&TODAY()-7,alerts!B2:B1000,"<="&TODAY()+1)'],
  ];
  styleMetricBlock(sheet.getRange("A4:B6"));
  styleMetricBlock(sheet.getRange("D4:E6"));

  sheet.getRange("A9:H9").values = [["最新ニュース", "", "", "", "", "", "", ""]];
  sheet.getRange("A9:H9").merge();
  sheet.getRange("A9").format = sectionTitleFormat();
  sheet.getRange("A10:H10").values = [["受信日", "キーワード", "タイトル", "媒体", "重要度", "ステータス", "担当", "メモ"]];
  sheet.getRange("A11:H16").formulas = [
    ["=alerts!B2", "=alerts!C2", "=alerts!D2", "=alerts!E2", "=alerts!I2", "=alerts!H2", "=alerts!J2", "=alerts!M2"],
    ["=alerts!B3", "=alerts!C3", "=alerts!D3", "=alerts!E3", "=alerts!I3", "=alerts!H3", "=alerts!J3", "=alerts!M3"],
    ["=alerts!B4", "=alerts!C4", "=alerts!D4", "=alerts!E4", "=alerts!I4", "=alerts!H4", "=alerts!J4", "=alerts!M4"],
    ["=alerts!B5", "=alerts!C5", "=alerts!D5", "=alerts!E5", "=alerts!I5", "=alerts!H5", "=alerts!J5", "=alerts!M5"],
    ["=alerts!B6", "=alerts!C6", "=alerts!D6", "=alerts!E6", "=alerts!I6", "=alerts!H6", "=alerts!J6", "=alerts!M6"],
    ["=alerts!B7", "=alerts!C7", "=alerts!D7", "=alerts!E7", "=alerts!I7", "=alerts!H7", "=alerts!J7", "=alerts!M7"],
  ];
  styleTable(sheet.getRange("A10:H16"));
  sheet.getRange("A11:A16").format.numberFormat = "yyyy-mm-dd hh:mm";

  sheet.getRange("J4:K8").values = [
    ["ステータス", "件数"],
    ["未確認", ""],
    ["確認中", ""],
    ["共有済み", ""],
    ["保留", ""],
  ];
  sheet.getRange("K5:K8").formulas = [
    ['=COUNTIF(alerts!H2:H1000,J5)'],
    ['=COUNTIF(alerts!H2:H1000,J6)'],
    ['=COUNTIF(alerts!H2:H1000,J7)'],
    ['=COUNTIF(alerts!H2:H1000,J8)'],
  ];
  styleTable(sheet.getRange("J4:K8"));

  const chart = sheet.charts.add("bar", sheet.getRange("J4:K8"));
  chart.title = "ステータス別件数";
  chart.hasLegend = false;
  chart.xAxis = { axisType: "textAxis" };
  chart.yAxis = { numberFormatCode: "0" };
  chart.setPosition("J10", "O24");

  setWidths(sheet, {
    A: 130,
    B: 160,
    C: 320,
    D: 150,
    E: 90,
    F: 110,
    G: 100,
    H: 220,
    J: 130,
    K: 80,
  });
}

function buildAlerts(sheet) {
  const headers = [
    "id",
    "receivedAt",
    "alertKeyword",
    "title",
    "source",
    "url",
    "snippet",
    "status",
    "priority",
    "owner",
    "category",
    "tags",
    "memo",
    "publishedAt",
    "importedAt",
  ];
  sheet.getRange("A1:O1").values = [headers];
  sheet.getRange("A2:O7").values = [
    [
      "sample-001",
      dateTime(2026, 5, 25, 9, 20),
      "ペット保険 市場",
      "ペット関連サービス、定期利用型の加入動向に変化",
      "業界ニュースサンプル",
      "https://example.com/pet-market-subscription",
      "価格だけでなく補償範囲や手続きの簡便さが選定要素になっている。",
      "未確認",
      "高",
      "佐藤",
      "市場動向",
      "市場,加入動向",
      "",
      dateOnly(2026, 5, 25),
      dateTime(2026, 5, 25, 10, 0),
    ],
    [
      "sample-002",
      dateTime(2026, 5, 24, 8, 45),
      "動物医療 DX",
      "動物病院向けオンライン受付、都市部で導入が進む",
      "ヘルスケアビジネスサンプル",
      "https://example.com/vet-dx-checkin",
      "オンライン受付や事前問診により来院前の情報整理と待ち時間短縮を目指す動き。",
      "確認中",
      "中",
      "田中",
      "医療DX",
      "動物病院,DX",
      "既存顧客向けサービス連携の参考にする。",
      dateOnly(2026, 5, 24),
      dateTime(2026, 5, 24, 9, 30),
    ],
    [
      "sample-003",
      dateTime(2026, 5, 23, 12, 10),
      "保険 金融庁",
      "少額短期保険分野の情報開示に関する資料が更新",
      "行政情報サンプル",
      "https://example.com/insurance-disclosure-update",
      "情報開示や顧客説明に関する資料が更新された。開示項目の確認が必要。",
      "共有済み",
      "高",
      "鈴木",
      "規制",
      "規制,開示",
      "週次会議で共有済み。",
      dateOnly(2026, 5, 23),
      dateTime(2026, 5, 23, 13, 0),
    ],
    [
      "sample-004",
      dateTime(2026, 5, 21, 17, 35),
      "ペット 高齢化",
      "高齢ペット向けケアサービス、相談窓口の需要が増加",
      "生活産業サンプル",
      "https://example.com/senior-pet-care",
      "高齢ペットの通院、食事、介護に関する相談が増えている。",
      "未確認",
      "中",
      "未設定",
      "顧客ニーズ",
      "高齢化,ケア",
      "",
      dateOnly(2026, 5, 21),
      dateTime(2026, 5, 21, 18, 10),
    ],
    [
      "sample-005",
      dateTime(2026, 5, 18, 10, 5),
      "ペット保険 比較",
      "ペット保険比較サイト、補償条件の見せ方を刷新",
      "マーケティングサンプル",
      "https://example.com/pet-insurance-comparison-ui",
      "免責条件や通院補償を分かりやすく表示する改修が進んでいる。",
      "確認中",
      "低",
      "佐藤",
      "競合調査",
      "比較,UI",
      "",
      dateOnly(2026, 5, 18),
      dateTime(2026, 5, 18, 10, 40),
    ],
    [
      "sample-006",
      dateTime(2026, 5, 17, 14, 50),
      "ペット テック",
      "見守りデバイスの利用データ、健康管理サービスへ展開",
      "テック市場サンプル",
      "https://example.com/pet-tech-health-data",
      "活動量や食事量などのデータをもとに体調変化の把握を支援する。",
      "保留",
      "低",
      "田中",
      "周辺サービス",
      "データ,テック",
      "短期より中長期の参考。",
      dateOnly(2026, 5, 17),
      dateTime(2026, 5, 17, 15, 20),
    ],
  ];
  sheet.freezePanes.freezeRows(1);
  styleTable(sheet.getRange("A1:O7"));
  sheet.getRange("B2:B1000").format.numberFormat = "yyyy-mm-dd hh:mm";
  sheet.getRange("N2:N1000").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("O2:O1000").format.numberFormat = "yyyy-mm-dd hh:mm";
  sheet.getRange("H2:H1000").dataValidation = {
    rule: { type: "list", formula1: "settings!$A$2:$A$5" },
  };
  sheet.getRange("I2:I1000").dataValidation = {
    rule: { type: "list", formula1: "settings!$B$2:$B$4" },
  };
  sheet.getRange("J2:J1000").dataValidation = {
    rule: { type: "list", formula1: "team!$A$2:$A$20" },
  };
  sheet.getRange("K2:K1000").dataValidation = {
    rule: { type: "list", formula1: "categories!$A$2:$A$20" },
  };
  setWidths(sheet, {
    A: 140,
    B: 150,
    C: 150,
    D: 320,
    E: 160,
    F: 280,
    G: 360,
    H: 100,
    I: 90,
    J: 100,
    K: 120,
    L: 160,
    M: 260,
    N: 120,
    O: 150,
  });
}

function buildKeywords(sheet) {
  sheet.getRange("A1:I1").values = [[
    "searchTerm",
    "alertName",
    "frequency",
    "sources",
    "language",
    "region",
    "owner",
    "active",
    "notes",
  ]];
  sheet.getRange("A2:I8").values = [
    ["ペット保険", "Googleアラート: ペット保険", "1日1回", "ニュース", "日本語", "日本", "佐藤", "TRUE", "市場・競合の基本監視"],
    ["ペット保険 比較", "Googleアラート: ペット保険 比較", "1日1回", "ニュース", "日本語", "日本", "佐藤", "TRUE", "比較サイトや訴求変化を見る"],
    ["動物医療 DX", "Googleアラート: 動物医療 DX", "1日1回", "ニュース", "日本語", "日本", "田中", "TRUE", "周辺サービスと業務効率化"],
    ["ペット 高齢化", "Googleアラート: ペット 高齢化", "1日1回", "ニュース", "日本語", "日本", "未設定", "TRUE", "顧客ニーズの変化"],
    ["少額短期保険 ペット", "Googleアラート: 少額短期保険 ペット", "1日1回", "ニュース", "日本語", "日本", "鈴木", "TRUE", "規制・市場動向"],
    ["ペット テック", "Googleアラート: ペット テック", "週1回", "ニュース", "日本語", "日本", "田中", "TRUE", "中長期の周辺動向"],
    ["動物病院 経営", "Googleアラート: 動物病院 経営", "週1回", "ニュース", "日本語", "日本", "未設定", "TRUE", "医療提供側の変化"],
  ];
  sheet.freezePanes.freezeRows(1);
  styleTable(sheet.getRange("A1:I8"));
  sheet.getRange("C2:C100").dataValidation = {
    rule: { type: "list", formula1: "settings!$C$2:$C$4" },
  };
  sheet.getRange("G2:G100").dataValidation = {
    rule: { type: "list", formula1: "team!$A$2:$A$20" },
  };
  setWidths(sheet, { A: 180, B: 240, C: 100, D: 110, E: 100, F: 90, G: 100, H: 80, I: 280 });
}

function buildCategories(sheet) {
  sheet.getRange("A1:D1").values = [["category", "description", "defaultPriority", "owner"]];
  sheet.getRange("A2:D9").values = [
    ["市場動向", "加入動向、価格、商品設計など", "中", "佐藤"],
    ["競合調査", "他社商品、比較サイト、広告表現", "中", "佐藤"],
    ["規制", "行政、法令、監督指針、情報開示", "高", "鈴木"],
    ["医療DX", "動物病院、診療、受付、問診、データ連携", "中", "田中"],
    ["顧客ニーズ", "飼い主行動、高齢化、相談、ケア", "中", "未設定"],
    ["周辺サービス", "ペットテック、見守り、フード、介護", "低", "田中"],
    ["メディア露出", "ニュース、SNS、テレビ、特集", "低", "未設定"],
    ["その他", "分類未定", "低", "未設定"],
  ];
  sheet.freezePanes.freezeRows(1);
  styleTable(sheet.getRange("A1:D9"));
  setWidths(sheet, { A: 140, B: 360, C: 120, D: 100 });
}

function buildTeam(sheet) {
  sheet.getRange("A1:D1").values = [["owner", "role", "email", "notes"]];
  sheet.getRange("A2:D6").values = [
    ["佐藤", "市場・競合", "", ""],
    ["田中", "周辺サービス・DX", "", ""],
    ["鈴木", "規制・開示", "", ""],
    ["未設定", "未アサイン", "", "一時置き場"],
    ["共有用", "全体共有", "", ""],
  ];
  sheet.freezePanes.freezeRows(1);
  styleTable(sheet.getRange("A1:D6"));
  setWidths(sheet, { A: 110, B: 180, C: 220, D: 260 });
}

function buildSettings(sheet) {
  sheet.getRange("A1:E1").values = [["status", "priority", "frequency", "source", "boolean"]];
  sheet.getRange("A2:E6").values = [
    ["未確認", "高", "都度", "自動", "TRUE"],
    ["確認中", "中", "1日1回", "ニュース", "FALSE"],
    ["共有済み", "低", "週1回", "ブログ", ""],
    ["保留", "", "", "Web", ""],
    ["除外", "", "", "その他", ""],
  ];
  sheet.freezePanes.freezeRows(1);
  styleTable(sheet.getRange("A1:E6"));
  setWidths(sheet, { A: 110, B: 90, C: 100, D: 100, E: 90 });
}

function buildSetup(sheet) {
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").values = [["セットアップ手順"]];
  sheet.getRange("A1").format = {
    fill: colors.header,
    font: { bold: true, color: colors.white, size: 16 },
    verticalAlignment: "middle",
  };
  sheet.getRange("A1:F1").format.rowHeightPx = 38;
  sheet.getRange("A3:F10").values = [
    ["1", "Googleアラートを作成", "keywordsシートのsearchTermを参考に、Googleアラートをメール配信で作成します。", "", "", ""],
    ["2", "Apps Scriptを貼り付け", "apps-script/Code.gs と apps-script/Index.html をスプレッドシートに設定します。", "", "", ""],
    ["3", "setupを実行", "初回だけ権限を許可し、シート構成とラベルを作ります。", "", "", ""],
    ["4", "importGoogleAlertsを実行", "Gmail内のGoogleアラートメールをalertsシートへ取り込みます。", "", "", ""],
    ["5", "createHourlyTriggerを実行", "自動取り込みを1時間ごとに設定します。", "", "", ""],
    ["6", "Webアプリとして公開", "チーム用の確認ページを共有します。", "", "", ""],
    ["7", "運用ルールを決める", "重要度、担当、共有済みの判断基準をチームで揃えます。", "", "", ""],
    ["8", "週次で棚卸し", "未確認・保留を見直し、必要な記事を会議や資料に反映します。", "", "", ""],
  ];
  styleTable(sheet.getRange("A3:F10"));
  sheet.getRange("A3:A10").format = {
    fill: colors.soft,
    font: { bold: true, color: colors.brand },
    horizontalAlignment: "center",
  };
  setWidths(sheet, { A: 50, B: 180, C: 620, D: 40, E: 40, F: 40 });
}

function styleMetricBlock(range) {
  range.format = {
    fill: colors.white,
    font: { color: colors.ink },
    border: {
      top: { style: "continuous", color: colors.line },
      bottom: { style: "continuous", color: colors.line },
      left: { style: "continuous", color: colors.line },
      right: { style: "continuous", color: colors.line },
      insideHorizontal: { style: "continuous", color: colors.line },
      insideVertical: { style: "continuous", color: colors.line },
    },
  };
  range.getColumn(0).format = {
    fill: colors.soft,
    font: { bold: true, color: colors.brand },
  };
  range.getColumn(1).format = {
    font: { bold: true, size: 18, color: colors.ink },
    horizontalAlignment: "center",
  };
}

function styleTable(range) {
  range.format = {
    fill: colors.white,
    font: { color: colors.ink },
    wrapText: true,
    verticalAlignment: "top",
    border: {
      top: { style: "continuous", color: colors.line },
      bottom: { style: "continuous", color: colors.line },
      left: { style: "continuous", color: colors.line },
      right: { style: "continuous", color: colors.line },
      insideHorizontal: { style: "continuous", color: colors.line },
      insideVertical: { style: "continuous", color: colors.line },
    },
  };
  range.getRow(0).format = {
    fill: colors.header,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
  };
}

function sectionTitleFormat() {
  return {
    fill: colors.brand,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "left",
  };
}

function setWidths(sheet, widths) {
  for (const [letter, width] of Object.entries(widths)) {
    sheet.getRange(`${letter}:${letter}`).format.columnWidthPx = width;
  }
}

function dateTime(year, month, day, hour, minute) {
  return new Date(Date.UTC(year, month - 1, day, hour, minute));
}

function dateOnly(year, month, day) {
  return new Date(Date.UTC(year, month - 1, day));
}
