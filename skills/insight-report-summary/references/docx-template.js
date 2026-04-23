"use strict";
// ════════════════════════════════════════════════════════════
//  invosData 市場洞察分析報告 DOCX Template
//  修改 ── REPORT CONFIG ── 區塊後執行：node generate_insight_report.js
// ════════════════════════════════════════════════════════════
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, WidthType, BorderStyle,
  ShadingType, VerticalAlign, PageNumber, Header, Footer
} = require("docx");
const fs = require("fs");
const path = require("path");

// ── REPORT CONFIG ─────────────────────────────────────────────
const BRAND   = "美強生";                   // 品牌名稱
const PERIOD  = "2602MAT";                  // 報告期間標示
const DATE_RANGE = "2025/03 ～ 2026/02";    // 觀測期（本期）
const DATE_COMP  = "2024/03 ～ 2025/02";    // 觀測期（比較期）
const SAMPLE_N   = "55,076";               // 樣本數
const OUT = path.join(__dirname, `../../mnt/outputs/${BRAND}嬰幼兒奶粉市場洞察分析報告_invos.docx`);

// ── Color helpers ──────────────────────────────────────────────
const BLUE   = "2E75B6";
const BLUE_L = "D5E8F0";
const GREEN  = "16A34A";
const TEAL   = "0891B2";
const PURPLE = "7C3AED";
const GRAY_L = "F0F4F8";
const GRAY_B = "CCCCCC";
const DARK   = "1F2937";
const WHITE  = "FFFFFF";

// ── Shared border ─────────────────────────────────────────────
const bdr = (color = GRAY_B) => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = (c = GRAY_B) => ({ top: bdr(c), bottom: bdr(c), left: bdr(c), right: bdr(c) });
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helpers ───────────────────────────────────────────────────
const sp = (before = 0, after = 0) => ({ spacing: { before, after } });
const P = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, font: "Arial", ...opts })],
  ...sp(opts.before || 0, opts.after || 120)
});
const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: BLUE })],
  ...sp(360, 180)
});
const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: DARK })],
  ...sp(240, 120)
});
const BULLET = (text, bold = false) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text, font: "Arial", size: 22, bold })],
  ...sp(0, 60)
});
const BLANK = () => new Paragraph({ children: [new TextRun("")], ...sp(0, 0) });

// ── Source citation helper（每段數據／論述尾端附來源頁碼或標籤）─────
// 用法：new Paragraph({ children: [new TextRun({...主文...}), srcRun("P8")] })
// 標籤規則：P## = 原始頁碼 / P## + 推論 / 【策略推論】 / 【invos 服務】 / 【策略方向】
const srcRun = (src) => new TextRun({
  text: ` (${src})`, font: "Arial", size: 18, italics: true, color: "888888"
});

// P with source — 單段 Paragraph 主文 + 來源
const PSRC = (text, src, opts = {}) => new Paragraph({
  children: [
    new TextRun({ text, font: "Arial", size: opts.size || 22, ...opts }),
    srcRun(src)
  ],
  ...sp(opts.before || 0, opts.after || 120)
});

// ── Data Table (指標對比) ──────────────────────────────────────
// headers: string[]
// rows: string[][]
// colWidths: number[] (DXA, must sum to 9026 for A4 with 1" margins)
function dataTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: borders(GRAY_B),
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: BLUE_L, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 20 })] })]
    }))
  });
  const dataRows = rows.map(row => new TableRow({
    children: row.map((cell, i) => new TableCell({
      borders: borders(GRAY_B),
      width: { size: colWidths[i], type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 20 })] })]
    }))
  }));
  return new Table({ width: { size: totalW, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
}

// ── Insight Box (洞察框，灰底) ─────────────────────────────────
// title: "洞察 N 《標題》"
// sections: [{label: "▶ 數據發現", text: "..."}, ...]
function insightBox(title, sections) {
  const rows = [];
  // Title row
  rows.push(new TableRow({
    children: [new TableCell({
      borders: borders(BLUE),
      shading: { fill: BLUE_L, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      columnSpan: 1,
      children: [new Paragraph({
        children: [new TextRun({ text: title, font: "Arial", size: 24, bold: true, color: BLUE })],
        ...sp(0, 0)
      })]
    })]
  }));
  // Content rows
  sections.forEach(({ label, text }) => {
    rows.push(new TableRow({
      children: [new TableCell({
        borders: borders(GRAY_B),
        shading: { fill: GRAY_L, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [
          new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: 22, bold: true, color: DARK })], ...sp(0, 60) }),
          new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 22 })], ...sp(0, 0) })
        ]
      })]
    }));
  });
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows });
}

// ── Service Table (服務建議表格) ───────────────────────────────
// svc: { name, color, tagline }
// items: [{ title, text }]
function serviceTable(svc, items) {
  const leftW = 1600, rightW = 7426;
  const contentRows = items.map((item, i) => new TableRow({
    children: [
      // 左欄只在第一行顯示服務名
      new TableCell({
        borders: borders(),
        width: { size: leftW, type: WidthType.DXA },
        shading: { fill: i === 0 ? svc.color : "F8F8F8", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        rowSpan: i === 0 ? items.length : undefined,
        children: i === 0 ? [
          new Paragraph({ children: [new TextRun({ text: `invos ${svc.name}`, font: "Arial", size: 20, bold: true, color: WHITE })], ...sp(0, 40) }),
          new Paragraph({ children: [new TextRun({ text: svc.tagline, font: "Arial", size: 18, color: WHITE, italics: true })], ...sp(0, 0) })
        ] : [new Paragraph({ children: [] })],
      }),
      // 右欄：方案內容
      new TableCell({
        borders: borders(),
        width: { size: rightW, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [
          new Paragraph({ children: [new TextRun({ text: `◆ ${item.title}`, font: "Arial", size: 22, bold: true, color: DARK })], ...sp(0, 60) }),
          new Paragraph({ children: [new TextRun({ text: item.text, font: "Arial", size: 22 })], ...sp(0, 0) })
        ]
      })
    ]
  }));
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [leftW, rightW], rows: contentRows });
}

// ── Summary Table (策略重點彙總) ──────────────────────────────
// svcs: [{ name, color, items: string[] }]
function summaryTable(svcs) {
  const leftW = 1600, rightW = 7426;
  const rows = svcs.map(svc => new TableRow({
    children: [
      new TableCell({
        borders: borders(),
        width: { size: leftW, type: WidthType.DXA },
        shading: { fill: svc.color, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ children: [new TextRun({ text: `invos ${svc.name}`, font: "Arial", size: 20, bold: true, color: WHITE })], ...sp(0, 0) })]
      }),
      new TableCell({
        borders: borders(),
        width: { size: rightW, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: svc.items.map((item, i) => new Paragraph({
          children: [new TextRun({ text: `${["①","②","③","④"][i]} ${item}`, font: "Arial", size: 22 })],
          ...sp(0, i < svc.items.length - 1 ? 60 : 0)
        }))
      })
    ]
  }));
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [leftW, rightW], rows });
}

// ════════════════════════════════════════════════════════════
//  ▼▼▼  每次換新報告時，修改以下的報告內容  ▼▼▼
// ════════════════════════════════════════════════════════════

// ── 第一章：整體市場現況 ──────────────────────────────────────

// 1.1 市場成長結構表格
const mktTable = dataTable(
  ["指標", "2502MAT", "2602MAT", "YoY成長率"],
  [
    ["整體銷售額", "3.26 億", "3.61 億", "+11%"],
    ["購買人數（不重複）", "38,499 人", "39,053 人", "+1%"],
    ["人均購買次數", "3.79 次", "4.12 次", "+9%"],
    ["單次平均金額", "2,239 元", "2,242 元", "0%"],
  ],
  [3000, 1800, 1800, 2426]
);

// 1.2 各階奶粉表格
const stageTable = dataTable(
  ["階段", "銷售市佔", "人流滲透", "銷售額 YoY", "人流 YoY"],
  [
    ["一階", "44%", "48%", "+2%", "-4%"],
    ["三階", "44%", "55%", "+20%", "+11%"],
    ["四階", "5%", "13%", "+12%", "+10%"],
  ],
  [1400, 1800, 1800, 1900, 2126]
);

// ── 第二章：品牌專屬洞察 ──────────────────────────────────────
// 每則洞察用 insightBox() 生成，格式固定為三段：數據發現 / 市場洞察 / 突破方向

const insights = [
  {
    title: `洞察 1 「陷阱型成長」——銷售額成長掩蓋市佔流失危機`,
    sections: [
      { label: "▶ 數據發現", text: "市場 YoY +11%，美強生銷售額 YoY +7%，人均次數提升 +12%（達 3.87 次）。然而市佔率從 5.1% 悄然下滑至 4.9%，購買人數僅 +2%（2,587→2,646 人），遠落後於明治（+33%）與卡洛塔妮（+71%）的擴張速度。更值得警惕的是，優生新客人數 YoY 衰退 -14%，遠大於整體市場新客衰退 -8%。" },
      { label: "▶ 市場洞察", text: "美強生的「成長」是靠既有消費者「買更多次」製造的數字幻象，而非品牌影響力的真實擴張。在台灣出生率 2025 年大幅下滑 20% 的背景下，一旦這批舊客的孩子陸續斷奶，缺乏新客補充的美強生將面臨斷崖式流失。更危險的信號是：優生新客流失速度已超越市場惡化速度——這是品牌競爭力衰退的雙重早期警訊。" },
      { label: "▶ 突破方向", text: "美強生需要打破「銷售額有成長就是健康」的認知慣性。真正關鍵的指標是「新客獲取速度 vs 出生率下滑速度 vs 競品搶客速度」三維比值。當優生的新客流失率（-14%）已超出市場自然衰退（-8%），代表品牌正在額外流失原本可以留住的消費者，亟需找出並修補新客流失的核心斷點。" },
    ]
  },
  // ← 繼續加入洞察 2, 3, 4, 5，格式相同
];

// ── 第三章：invos 三大服務建議 ────────────────────────────────
const insightSvc = {
  name: "Insight",
  color: GREEN,
  tagline: "賦能智慧決策",
  items: [
    { title: `${BRAND}「健康成長指標」儀表板`, text: `建立追蹤「新客獲取速度 vs 出生率下滑速度 vs 競品市佔消長」的三合一品牌健康指數模型。每季提供結構性預警報告，讓${BRAND}管理層能在銷售額數字之外，清楚看見品牌真正的成長品質，提前制定防禦策略。` },
    { title: "三階奶粉競爭地圖與升級轉換率分析", text: `深度分析三階市場的通路結構、消費者輪廓、購買頻次，識別${BRAND}純睿/智睿最具機會的細分切入點。同步追蹤「一階→三階」消費者升級轉換率，比較${BRAND}與競品在升級旅程中的品牌留存率差異，精準定位流失節點。` },
    { title: "明治/卡洛塔妮崛起因素解碼報告", text: "拆解明治（銷售額 +33%）、卡洛塔妮（銷售額 +71%）快速成長的行為數據驅動因素：關鍵通路、目標客群輪廓、推動增長的購買行為模式，提供可對標學習的增長路徑參考。" },
  ]
};

const mediaSvc = {
  name: "Media",
  color: TEAL,
  tagline: "高效鎖客擴散",
  items: [
    { title: "高端線電商「數位信任感」種草計畫", text: `以鉑睿/純睿目標消費者行為輪廓（藥局高端購買記錄、30-39 歲女性）為基礎，在 momo、蝦皮官方旗艦等授權電商投放程序化廣告，搭配品牌教育內容（成分科普、媽媽真實評測影片）快速建立電商認知。` },
    { title: "親舒舊客「囤貨激活」精準行銷", text: "鎖定曾購買親舒但近期單次金額下滑的高忠誠舊客，在自然補貨週期前 2-3 週投放「家庭備貨組合優惠」廣告，激活從「少量多次補貨」→「一次囤足備貨」的消費習慣轉變，快速提升客單價。" },
    { title: "三階升級關鍵時點廣告攔截", text: "識別優生用戶中孩子即將進入 6-9 個月（一階→三階升級前夕）的目標族群，在升級決策視窗期（約 4-6 週）密集投放純睿/智睿的品牌教育廣告，率先建立「美強生三階是升級首選」的心智佔據。" },
  ]
};

const apiSvc = {
  name: "API",
  color: PURPLE,
  tagline: "私域深度經營",
  items: [
    { title: `${BRAND}全通路購買者私域識別池`, text: `整合大樹、蝦皮官方、momo、卡多摩、丁丁等授權通路的${BRAND}購買記錄，建立可識別、可觸達的「全通路${BRAND}買者私域名單」，為所有後續行銷計畫提供可自主運作的數據底座。同步可用於監測非授權通路異常流量，協助品牌掌握水貨對官方通路佈局的潛在影響。` },
    { title: "酷澎買者識別與授權通路回流引導", text: `透過電子發票數據識別「在酷澎有優生購買紀錄」的消費者，精準推送 momo 官方旗艦或品牌官網的「月配訂閱首購優惠」（每次配送設計在 1,100-1,300 元，貼近其熟悉的酷澎補貨單次金），搭配會員點數回饋，讓消費者覺得「訂閱授權通路更划算、更有歸屬感」。` },
    { title: "一階→三階品牌升級旅程個人化溝通", text: "以孩子月齡為主軸，系統性向一階奶粉購買者推送升級旅程內容：第 6 個月推送三階科普，第 9 個月推送純睿/智睿試用方案，第 12 個月推送升級專屬優惠，讓消費者在品牌旅程的每個節點都感受到美強生的陪伴。" },
    { title: "初乳/營養補充品交叉銷售識別", text: `數據顯示「幼兒黃金初乳蛋白粉」在奶粉買者的大樹併買排行中本期暴升 +247 名，躍入前 20。透過 invos API 識別${BRAND}買者中有此類跨品類需求的族群，協助品牌探索組合包／搭配推薦的新商業模式。` },
  ]
};

// ── 第四章：策略重點一覽 ──────────────────────────────────────
const summaryData = [
  { name: "Insight", color: GREEN, items: ["健康成長指標儀表板（洞察#1）", "三階市場競爭地圖與升級轉換率分析（洞察#5）", "明治/卡洛塔妮崛起因素拆解分析（洞察#1）"] },
  { name: "Media",   color: TEAL,   items: ["高端線授權電商種草計畫，鉑睿/純睿在 momo/蝦皮官方首發（洞察#3）", "親舒舊客囤貨激活廣告，補貨週期前精準干預（洞察#4）", "三階升級時點廣告攔截，孩子 6-9 個月決策視窗（洞察#5）"] },
  { name: "API",     color: PURPLE, items: ["全通路授權通路私域識別池整合，兼監測非授權通路異常（洞察#2#3）", "優生高頻買者訂閱方案推送，阻斷商品化漂移（洞察#2）", "一階→三階升級旅程個人化溝通 ＋ 初乳補充品交叉銷售識別（洞察#5）"] },
];

// ════════════════════════════════════════════════════════════
//  Document Assembly
// ════════════════════════════════════════════════════════════
async function main() {
  const children = [
    // ── 封面 ──
    BLANK(),
    new Paragraph({ children: [new TextRun({ text: "【引客數據 invos】", font: "Arial", size: 28, bold: true, color: BLUE })], alignment: AlignmentType.CENTER, ...sp(0, 80) }),
    new Paragraph({ children: [new TextRun({ text: `${PERIOD} 嬰幼兒奶粉市場洞察分析報告`, font: "Arial", size: 36, bold: true })], alignment: AlignmentType.CENTER, ...sp(80, 80) }),
    new Paragraph({ children: [new TextRun({ text: `${BRAND}品牌專屬市場洞察與 invos 服務策略建議`, font: "Arial", size: 24, italics: true, color: "555555" })], alignment: AlignmentType.CENTER, ...sp(0, 80) }),
    new Paragraph({ children: [new TextRun({ text: `數據觀測期：${DATE_RANGE} vs ${DATE_COMP}　｜　樣本：${SAMPLE_N} 人`, font: "Arial", size: 20, color: "888888" })], alignment: AlignmentType.CENTER, ...sp(0, 480) }),

    // ── 第一章 ──
    H1("一、嬰幼兒奶粉市場整體現況"),
    H2("1.1 市場成長結構"),
    P(`整體市場年銷售額 YoY +11%，但驅動力並非人口擴張，而是既有消費者「買更多次」。購買人數僅微幅成長 +1%，人均購買次數則提升 +9%（達 4.12 次）。這與台灣出生率大幅下滑高度相關——2025 年出生人口約 10.7 萬人，較 2024 年的 13.4 萬人銳減 20%，新生兒減少直接衝擊一階奶粉的購買人數，市場成長正在走向「靠舊客撐盤」的結構。`, { size: 22 }),
    BLANK(),
    mktTable,
    BLANK(),

    H2("1.2 各階奶粉表現"),
    P("三階奶粉在本期躍升為最大成長引擎，銷售規模已與一階並列（各 44% 市佔），人流與人均次數雙驅動。一階奶粉雖仍為新客購買主力，但受出生率下滑衝擊，人流持續萎縮。四階奶粉規模小但成長顯著。", { size: 22 }),
    BLANK(),
    stageTable,
    BLANK(),

    H2("1.3 通路分佈與競爭格局"),
    P("藥局仍是嬰幼兒奶粉核心戰場，銷售額佔 65%，觸及消費人數達 72%。2602MAT 藥局份額微降，部分市場被綜合電商承接——電商銷售額 YoY 暴增 +60%，份額從 9% 升至 13%，顯示電商通路已不再是平價品牌的專屬戰場，中高端品牌通路分佈同樣在往電商移動。", { size: 22 }),
    BULLET("大樹藥局：銷售額佔比 55%，觸及人數 54%，人均次數 3.95 次，黏著度最強"),
    BULLET("綜合電商（蝦皮、momo 等授權平台）：爆發性成長 +60%，消費者快速往電商補貨移動"),
    BULLET("卡多摩/婦嬰通路：單次金額高（約 4,000 元+），操作 APP 限定價/VIP 會員促囤貨"),
    BLANK(),

    H2("1.4 品牌競爭態勢"),
    P("市場形成三個競爭層級：高端小眾（啟賦、S-26 鉑臻靠高單價），中端主力（S-26、雀巢能恩人次兼顧），平價走量（雀巢能恩水解、明治、美強生優生靠廣泛滲透）。值得注意的是，明治（銷售額 +33%）、卡洛塔妮（+71%）正以快速人流擴張搶佔市場增量。", { size: 22 }),
    BLANK(),

    H2("1.5 新舊客買者狀態"),
    P("整體市場購買人數僅成長 +1%，但背後是舊客人流 +10%（20,529→22,501 人）與新客人流 -8%（17,970→16,552 人）相互抵銷的結果。整體市場其實正面臨新客萎縮的挑戰，只是被舊客成長掩蓋。考量到台灣 2025 年出生人口大幅下滑 20%，未來獲取新客將面臨更大挑戰。品牌的成長重點，應從「擴大人流」轉向「深化既有消費者的忠誠度與 LTV」。", { size: 22 }),
    BULLET("品類舊客（既有買者）：人流佔 58%（+10%），銷售佔 71%（+20%）；年均購買 5.07 次"),
    BULLET("品類新客（新進買者）：人流佔 42%（-8%），銷售佔 29%（-8%）；人數持續萎縮"),
    BULLET("新客購買力保守：一、三、四階新客單次金額均呈下滑趨勢，囤貨意願下降"),
    BLANK(),

    // ── 第二章 ──
    H1("二、美強生品牌專屬市場洞察"),
    P("以下五大洞察均基於美強生在數據中的獨特表現，具高度針對性。每一洞察均伴隨「突破方向」，指向可改變美強生既有認知的行動路徑。", { size: 22 }),
    BLANK(),
    ...insights.flatMap(ins => [insightBox(ins.title, ins.sections), BLANK()]),

    // ── 第三章 ──
    H1("三、invos 三大服務策略建議"),
    P("以下策略建議基於上述五大美強生專屬洞察，精準扣合引客數據三大服務功能，提供可落地執行的行動方向。", { size: 22 }),
    BLANK(),

    H2(`3.1 invos Insight｜市場機會洞察 × 賦能${BRAND}智慧決策`),
    new Paragraph({ children: [new TextRun({ text: "對應洞察：#1 陷阱型成長 ／ #5 三階升級斷層", font: "Arial", size: 20, italics: true, color: "666666" })], ...sp(0, 120) }),
    serviceTable(insightSvc, insightSvc.items),
    BLANK(),

    H2(`3.2 invos Media｜精準獲客轉換 × 高效鎖定${BRAND}目標客群`),
    new Paragraph({ children: [new TextRun({ text: "對應洞察：#2 優生商品化漂移 ／ #3 高端線電商空白 ／ #4 親舒隱形冠軍", font: "Arial", size: 20, italics: true, color: "666666" })], ...sp(0, 120) }),
    serviceTable(mediaSvc, mediaSvc.items),
    BLANK(),

    H2(`3.3 invos API｜全通路會員整合 × 建立${BRAND}私域流量護城河`),
    new Paragraph({ children: [new TextRun({ text: "對應洞察：#2 優生商品化漂移 ／ #4 親舒隱形冠軍 ／ #5 三階升級斷層", font: "Arial", size: 20, italics: true, color: "666666" })], ...sp(0, 120) }),
    serviceTable(apiSvc, apiSvc.items),
    BLANK(),

    // ── 第四章 ──
    H1("四、策略重點一覽"),
    summaryTable(summaryData),
    BLANK(),

    // ── 免責說明 ──
    new Paragraph({
      children: [new TextRun({ text: `本報告基於引客數據 invos 電子發票消費追蹤系統，觀測期 ${DATE_RANGE}（${PERIOD}） vs ${DATE_COMP}，樣本 ${SAMPLE_N} 人。所有洞察與策略建議均根據真實消費行為數據推導，具${BRAND}品牌專屬針對性，僅供品牌內部策略參考使用。`, font: "Arial", size: 18, italics: true, color: "888888" })],
      ...sp(240, 0)
    }),
  ];

  const doc = new Document({
    numbering: {
      config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }]
    },
    styles: {
      default: { document: { run: { font: "Arial", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Arial", color: BLUE }, paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: DARK }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      ]
    },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT, buffer);
  console.log("✅ Saved:", OUT);
}

main().catch(e => { console.error(e); process.exit(1); });
