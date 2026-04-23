"use strict";
// ════════════════════════════════════════════════════════════
//  invosData 一頁式行動建議 PPTX Template (2 slides)
//  使用方式：複製此檔案，修改下方 ── BRAND CONFIG ── 與 COLUMNS
//  即可為任何品牌生成「封面 + 一頁式行動建議整合頁」
//
//  視覺設計：深色底 + 透明卡片（只有外框）+ 彩色文字
//  — Header 卡：彩色外框 + 頂部細色條 + 分隔線，白字
//  — 關鍵資訊 / 行動建議卡：灰色細外框，字色用該欄 pastel 色
//  — Tag：白色細外框 + 白字
// ════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const path = require("path");

// ── BRAND CONFIG ─────────────────────────────────────────────
const BRAND_NAME  = "台塑生醫MD醫之方";
const CATEGORY    = "益生菌";
const REPORT_DATE = "2512MAT";
const SUBTITLE    = `${BRAND_NAME}｜通路客群行動建議｜當前瓶頸 x 核心機會 x invosData 數據解決方案`;
const COVER_BG_IMAGE = "";
const OUT = path.join(__dirname, `../../../output/action-recommendation/${BRAND_NAME}_行動建議_${REPORT_DATE}.pptx`);

// ─── Global Styles ───────────────────────────────────────────
const FONT = "Microsoft JhengHei";
const BODY_LINE_SP = 1.5;

// ─── Palette ─────────────────────────────────────────────────
const C = {
  bg:         "21253A",   // 深藍黑 slide background
  labelBar:   "202539",   // 「關鍵資訊」「行動建議」label bar bg
  cardEdge:   "64748B",   // 內容卡片灰色外框
  // Column main colors（header 邊框與 accent）
  cashcow:    "2563EB",   // 藍
  question:   "9333EA",   // 紫
  star:       "F59E0B",   // 橘
  dog:        "6B7280",   // 灰
  // 字體色
  white:      "FFFFFF",
  dim:        "94A3B8",
};

// ─── Column color presets ────────────────────────────────────
//   edge        = header 外框 / 頂部色條 / 分隔線
//   insightText = 關鍵資訊卡文字色（該欄 pastel 色）
//   actionText  = 行動建議卡文字色（該欄更淺 pastel）
// ─────────────────────────────────────────────────────────────
const COL_PRESETS = {
  cashcow:  { edge: C.cashcow,  insightText: "93C5FD", actionText: "93C5FD" },
  question: { edge: C.question, insightText: "C4B5FD", actionText: "BAACF1" },
  star:     { edge: C.star,     insightText: "FDE68A", actionText: "FCE589" },
  dog:      { edge: C.dog,      insightText: "D1D5DB", actionText: "C4C8CE" },
};

// ════════════════════════════════════════════════════════════
//  COLUMNS — 5 欄內容定義
// ════════════════════════════════════════════════════════════
const COLUMNS = [
  {
    preset: "cashcow",
    verb:   "守住",
    title:  "momo",
    subtitle: "高價值舊客基本盤",
    problem:  "問題不在消費力，在於留不住買者",
    goal:   [
      { label: "留存率 ", value: "36%" },
      { label: "、RPC 對標 ", value: "3,253 元" },
    ],
    insights: [
      { head: "客數流失，留存率僅 20%",
        body: "momo 客數 -18.6%、留存率 20%（低於市場均值 36%），流失主因為品類使用停止（69%）而非競品搶奪——不是產品不好，而是沒人提醒他再買" },
      { head: "momo 為最大營收通路",
        body: "momo 貢獻 MD醫之方主要營收份額，RPC 2,517元、舊客 RPC 達 3,243元（貼近通路均值 3,253元），留下來的人願意花錢，是產品線深度佈局的首選陣地" },
    ],
    actions: [
      { head: "品類喚醒：", body: "圈選 69% 「momo MD醫之方」品類流失買者投放廣告，推送品類需求內容（非折扣），重建使用習慣後再推產品，目標留存率 36%" },
      { head: "增加留存：", body: "設計定期購／組合包／發票回饋，目標提升 RPC 貼近通路均值 3,253 元" },
      { head: "自動補貨提醒：", body: "依消耗週期預估用罄時間 → 用罄前 7 天推播提醒，目標留住 RPC 3,243 元舊客" },
    ],
    tags: ["精準受眾", "會員留存"],
  },
  {
    preset: "question",
    verb:   "新客入口｜",
    title:  "康是美×寶雅",
    subtitle: "藥妝新客入口",
    problem:  "新客入口成立，但缺乏回購銜接",
    goal:   [
      { label: "", value: "88%" },
      { label: "＋新客、PRC 對標 ", value: "2,416 元" },
    ],
    insights: [
      { head: "康是美逆勢成長　留存結構不穩",
        body: "康是美 MD醫之方人流 +32.8%、營收 +37.3%（相較通路整體逆勢成長）、88% 新客，且這群買者在整體市場人均花 2,416 元買益生菌（MD醫之方只拿到 1,317 元），顯示他們有錢有需求，但買完就離開，缺乏橋樑接口" },
      { head: "寶雅正在長　競品不同",
        body: "寶雅競品是新普利／我的健康日記／大研生醫（非康是美 BHK's/倍適），顯示不同通路的人在比較不同品牌，產品佈局不能複製貼上" },
    ],
    actions: [
      { head: "新客入口：", body: "設計康是美／寶雅「品牌首購發票回饋」活動，目標讓 88%＋新客認識 MD醫之方" },
      { head: "回購橋樑系統：", body: "首購後透過發票存摺加入品牌會員，後續發票資料自動存入品牌資料庫，以利補貨週期自動化追蹤，目標留下該群買者達到 PRC 均值 2,416 元" },
      { head: "競品對標：", body: "針對通路內主要競品設計首購方案、試用包，目標截獲在選牌比較當下猶豫的新客" },
    ],
    tags: ["新客導流", "發票回饋"],
  },
  {
    preset: "star",
    verb:   "舊客留存×",
    title:  "回購深化",
    subtitle: "最高槓桿成長引擎",
    problem:  "舊客 PRC 為新客 2 倍，短期即可變現",
    goal:   [
      { label: "舊客佔比 ", value: "37%" },
      { label: " 、營收 ", value: "+14.7%" },
    ],
    insights: [
      { head: "舊客留存×回購深化最高槓桿",
        body: "MD醫之方舊客佔比 20%（低於市場舊客比例 37%）、舊客人均次 1.56 次（低於市場舊客均值 1.98 次），但 MD醫之方舊客 RPC 是新客 2.04 倍（整體市場新舊客 RPC 僅差 1.86 倍）——代表 MD醫之方舊客忠誠度不差、願意花得更多，每提升 1% 留存率的效益遠大於等量新客投入" },
    ],
    actions: [
      { head: "舊客會員經營：", body: "以 RFM 分層推不同產品，高頻高額客推升級試用、中頻客推回購組合包、沉默客推品類需求喚醒內容，目標舊客佔比從 20% 拉到市場水平 37%，若總人數不變，營收預期增加 14.7%" },
      { head: "自動補貨提醒：", body: "建立「補貨週期計算」規則預估用罄時間，提前推送發票推播，目標提升舊客購買頻率貼近市場舊客均值 1.98 次" },
    ],
    tags: ["會員留存", "LTV深化"],
  },
  {
    preset: "star",
    verb:   "",
    title:  "競品轉換客",
    subtitle: "",
    problem:  "舊客留存＋競品觸及 = 最高效益引擎",
    goal:   [
      { label: "競品客 ", value: "+16%" },
      { label: "、識別精準受眾", value: "" },
    ],
    insights: [
      { head: "競品轉換人流 +16%",
        body: "競品客人流成長，證明品牌有吸引力。發票可識別跨品牌購買者精準定向投放——他們正在比較，只要在對的時機推對的產品，就有機會轉換" },
    ],
    actions: [
      { head: "精準受眾：", body: "發票識別「互動高／康是美寶雅主要競品」買者，投放 MD醫之方品牌 Branding 廣告，提升品牌識別，創造人流入口" },
      { head: "購買誘因：", body: "在康是美寶雅貨架鄰近位置加強 MD醫之方露出、設計「轉換體驗包」、「發票回饋」競品客專屬獎勵，搶佔「比較購買」場景的最後一哩決策，目標維持競品客人流成長 16%" },
    ],
    tags: ["精準受眾", "發票回饋"],
  },
  {
    preset: "dog",
    verb:   "",
    title:  "酷澎水貨×品牌溢價",
    subtitle: "",
    problem:  "非授權通路治理・導流正規",
    goal:   [
      { label: "佔比 ", value: "17%" },
      { label: "、RPC 僅 ", value: "1,160 元" },
    ],
    insights: [
      { head: "酷澎低價稀釋品牌溢價",
        body: "酷澎 MD醫之方客數暴增 3.7 倍至 855 人，但 RPC 僅 1,160 元，為 momo 的 46%，82% 是品牌新客，被低價而非品牌力吸引，難以拉高客單" },
      { head: "銷額佔比成長至 17%",
        body: "酷澎銷額 YoY +288%，高回購率 35% 反映的是價格敏感者穩定性而非品牌忠誠" },
    ],
    actions: [
      { head: "監測量化：", body: "追蹤酷澎交易規模，作為通路保護與定價追蹤" },
      { head: "導流正規：", body: "以官方獨家發票回饋（momo／官網）、官方專屬會員積點等發票活動，吸引價格敏感客回流" },
      { head: "受眾圈選：", body: "發票圈選酷澎買者至品牌可控通路建立會員關係" },
    ],
    tags: ["通路監測", "精準受眾"],
  },
];

// ════════════════════════════════════════════════════════════
//  SOURCES — Slide 3「附數據來源」所需，與 COLUMNS 一對一
//  P## = 原始報告頁碼；
//  【策略推論】= 非原始數據、為分析推論（寫入前必須先跟使用者確認）
//  【invos 服務】= invos 產品能力，非原始數據
//  【策略方向】= goal 為方向性語言、無具體 benchmark
//
//  （下方範例為 MD醫之方 益生菌 COLUMNS 對應；換新品牌時整組改）
// ════════════════════════════════════════════════════════════
const SOURCES = [
  { // Col 1 守住 momo
    problem: "報告 momo 通路頁",
    goal: "報告 momo 通路頁（留存率 36%、RPC 3,253）",
    insights: [
      { body: "報告 momo 通路頁（客數 -18.6%、留存率 20%）" },
      { body: "報告 momo 通路頁（RPC 2,517、舊客 RPC 3,243、通路均值 3,253）" },
    ],
    actions: [
      { body: "【invos 服務】+ 策略推論" },
      { body: "報告 momo 通路頁 + 策略推論" },
      { body: "【invos 服務】+ 策略推論" },
    ],
  },
  { // Col 2 新客入口 康是美×寶雅
    problem: "報告 康是美 / 寶雅 通路頁",
    goal: "報告 康是美 / 寶雅 通路頁",
    insights: [
      { body: "報告 康是美 通路頁（+32.8%人流、+37.3%營收、88%新客、PRC 2,416 vs 1,317）" },
      { body: "報告 寶雅 通路頁（主要競品非康是美同款）" },
    ],
    actions: [
      { body: "【invos 服務】+ 策略推論" },
      { body: "【invos 服務】+ 策略推論" },
      { body: "報告 寶雅 通路頁 + 策略推論" },
    ],
  },
  { // Col 3 舊客留存×回購深化
    problem: "報告 舊客分析頁",
    goal: "報告 市場舊客佔比 + 估算推論",
    insights: [
      { body: "報告 舊客分析頁（舊客 20% vs 市場 37%、舊客 RPC 為新客 2.04 倍 vs 市場 1.86 倍）" },
    ],
    actions: [
      { body: "【invos 服務】+ 策略推論" },
      { body: "【invos 服務】+ 策略推論" },
    ],
  },
  { // Col 4 競品轉換客
    problem: "報告 競品客分析頁",
    goal: "報告 競品客 +16%",
    insights: [
      { body: "報告 競品客分析頁（人流 +16%）" },
    ],
    actions: [
      { body: "【invos 服務】+ 策略推論" },
      { body: "【invos 服務】+ 策略推論" },
    ],
  },
  { // Col 5 酷澎水貨×品牌溢價
    problem: "報告 酷澎 通路頁",
    goal: "報告 酷澎 通路頁（佔比 17%、RPC 1,160）",
    insights: [
      { body: "報告 酷澎 通路頁（客數 x3.7、RPC 1,160、82% 新客）" },
      { body: "報告 酷澎 通路頁（銷額 +288%、回購率 35%）" },
    ],
    actions: [
      { body: "【invos 服務】+ 策略推論" },
      { body: "【invos 服務】+ 策略推論" },
      { body: "【invos 服務】+ 策略推論" },
    ],
  },
];

// ════════════════════════════════════════════════════════════
//  Helpers
// ════════════════════════════════════════════════════════════
function drawPageNum(sl, num) {
  sl.addText(String(num), {
    x: 0.15, y: 7.1, w: 0.4, h: 0.3, margin: 0,
    fontFace: FONT, fontSize: 8, color: C.white, align: "left", valign: "bottom"
  });
}

function drawLabelBar(sl, x, y, w, text) {
  // Dark bar with gray border — matches sample's label bar style
  sl.addShape("rect", {
    x, y, w, h: 0.33,
    fill: { color: C.labelBar },
    line: { color: C.cardEdge, width: 0.5 }
  });
  sl.addText(text, {
    x: x + 0.06, y: y + 0.03, w: w - 0.12, h: 0.27, margin: 0,
    fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle"
  });
}

// ════════════════════════════════════════════════════════════
//  SLIDE 1 — 封面頁
// ════════════════════════════════════════════════════════════
function slideCover(pres) {
  const sl = pres.addSlide();
  if (COVER_BG_IMAGE) {
    sl.background = { path: COVER_BG_IMAGE };
    sl.addShape("rect", { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: C.bg, transparency: 25 } });
  } else {
    sl.background = { color: C.bg };
  }

  // 中央長條圖 icon
  const iconCX = 6.65, iconCY = 2.8;
  const barW = 0.22, barGap = 0.08;
  const bars = [
    { h: 0.45, y: iconCY + 0.55 },
    { h: 0.70, y: iconCY + 0.30 },
    { h: 1.00, y: iconCY },
  ];
  sl.addShape("rect", { x: iconCX - 0.55, y: iconCY - 0.15, w: 0.06, h: 1.25, fill: { color: C.white }, line: { color: C.white, width: 0 } });
  sl.addShape("rect", { x: iconCX - 0.55, y: iconCY + 1.04, w: 1.60, h: 0.06, fill: { color: C.white }, line: { color: C.white, width: 0 } });
  bars.forEach((b, i) => {
    const bx = iconCX - 0.15 + i * (barW + barGap);
    sl.addShape("rect", { x: bx, y: b.y, w: barW, h: b.h, fill: { color: C.white }, line: { color: C.white, width: 0 } });
  });
  sl.addText("✦", {
    x: iconCX + 0.45, y: iconCY - 0.35, w: 0.35, h: 0.35, margin: 0,
    fontFace: FONT, fontSize: 12, color: C.white, align: "center", valign: "middle"
  });
  sl.addText("行動建議", {
    x: 2.5, y: 3.65, w: 8.3, h: 1.2, margin: 0,
    fontFace: FONT, fontSize: 36, bold: true, color: C.white, align: "center", valign: "middle"
  });
  drawPageNum(sl, 1);
}

// ════════════════════════════════════════════════════════════
//  SLIDE 2 — 一頁式行動建議整合頁
// ════════════════════════════════════════════════════════════
function slideOnePager(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };

  // ── Layout anchors（以樣本檔座標為準）──
  const X0 = 0.20;
  const CW = 2.54;
  const GAP = 0.04;
  const TOTAL_W = CW * 5 + GAP * 4;  // 12.86

  // ── Top title bar ──
  sl.addShape("rect", {
    x: X0, y: 0.10, w: TOTAL_W, h: 0.28,
    fill: { color: C.labelBar }, line: { color: C.cardEdge, width: 0.5 }
  });
  sl.addText(SUBTITLE, {
    x: X0 + 0.06, y: 0.10, w: TOTAL_W - 0.12, h: 0.28, margin: 0,
    fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle"
  });

  // ── Row 1: Header 目標卡（透明 + 彩色外框 + 頂條 + 分隔線）──
  const HDR_Y = 0.46, HDR_H = 0.75;
  const TOP_STRIP_H = 0.05;
  const DIVIDER_Y_OFFSET = 0.46;  // 分隔線位於 header 內的 y 偏移
  const DIVIDER_H = 0.02;

  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];

    // 透明外框（只有線）
    sl.addShape("rect", {
      x: cx, y: HDR_Y, w: CW, h: HDR_H,
      fill: { type: "none" }, line: { color: p.edge, width: 0.8 }
    });
    // 頂部色條
    sl.addShape("rect", {
      x: cx, y: HDR_Y, w: CW, h: TOP_STRIP_H,
      fill: { color: p.edge }, line: { color: p.edge, width: 0 }
    });
    // 底部分隔線
    sl.addShape("rect", {
      x: cx + 0.06, y: HDR_Y + DIVIDER_Y_OFFSET, w: CW - 0.14, h: DIVIDER_H,
      fill: { color: p.edge }, line: { color: p.edge, width: 0 }
    });

    // 標題行：verb + title
    const titleRuns = [];
    if (col.verb) titleRuns.push({ text: col.verb, options: { fontSize: 10, bold: true, color: C.white } });
    titleRuns.push({ text: col.title, options: { fontSize: 10, bold: true, color: C.white } });
    if (col.subtitle) titleRuns.push({ text: " " + col.subtitle, options: { fontSize: 10, bold: true, color: C.white } });
    sl.addText(titleRuns, {
      x: cx + 0.08, y: HDR_Y + 0.06, w: CW - 0.16, h: 0.24, margin: 0,
      fontFace: FONT, valign: "middle"
    });
    // Problem 副標
    sl.addText(col.problem, {
      x: cx + 0.08, y: HDR_Y + 0.28, w: CW - 0.16, h: 0.18, margin: 0,
      fontFace: FONT, fontSize: 8, color: C.white, valign: "middle"
    });
    // 目標行
    const goalRuns = [{ text: "目標｜", options: { fontSize: 8, bold: true, color: C.white } }];
    col.goal.forEach(g => {
      if (g.label) goalRuns.push({ text: g.label, options: { fontSize: 8, color: C.white } });
      if (g.value) goalRuns.push({ text: g.value, options: { fontSize: 8, bold: true, color: C.white } });
    });
    sl.addText(goalRuns, {
      x: cx + 0.08, y: HDR_Y + 0.50, w: CW - 0.16, h: 0.22, margin: 0,
      fontFace: FONT, valign: "middle"
    });
  });

  // ── Label bar: 關鍵資訊 ──
  const KEY_LABEL_Y = 1.28;
  drawLabelBar(sl, X0, KEY_LABEL_Y, TOTAL_W, "關鍵資訊");

  // ── Row 2: 關鍵資訊卡（透明 + 灰邊 + 彩色文字）──
  const KEY_Y = 1.64, KEY_H = 2.50;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    // 透明外框
    sl.addShape("rect", {
      x: cx, y: KEY_Y, w: CW, h: KEY_H,
      fill: { type: "none" }, line: { color: C.cardEdge, width: 0.5 }
    });
    // 內容：粗體小標 + 說明段，字色用 pastel
    const runs = [];
    col.insights.forEach((insight, j) => {
      const isLast = (j === col.insights.length - 1);
      runs.push({ text: insight.head, options: { fontSize: 10, bold: true, color: p.insightText, breakLine: true } });
      runs.push({ text: insight.body, options: { fontSize: 8, color: C.white, breakLine: !isLast, paraSpaceAfter: !isLast ? 4 : 0 } });
    });
    sl.addText(runs, {
      x: cx + 0.10, y: KEY_Y + 0.08, w: CW - 0.20, h: KEY_H - 0.16, margin: 0,
      fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP
    });
  });

  // ── Label bar: 行動建議 ──
  const ACT_LABEL_Y = 4.24;
  drawLabelBar(sl, X0, ACT_LABEL_Y, TOTAL_W, "行動建議");

  // ── Row 3: 行動建議卡（透明 + 灰邊 + 彩色文字）──
  const ACT_Y = 4.60, ACT_H = 2.45;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    sl.addShape("rect", {
      x: cx, y: ACT_Y, w: CW, h: ACT_H,
      fill: { type: "none" }, line: { color: C.cardEdge, width: 0.5 }
    });
    const runs = [];
    col.actions.forEach((a, j) => {
      const isLast = (j === col.actions.length - 1);
      runs.push({ text: a.head, options: { bullet: true, fontSize: 8, bold: true, color: p.actionText } });
      runs.push({ text: a.body, options: { fontSize: 8, color: C.white, breakLine: !isLast, paraSpaceAfter: !isLast ? 4 : 0 } });
    });
    sl.addText(runs, {
      x: cx + 0.10, y: ACT_Y + 0.08, w: CW - 0.20, h: ACT_H - 0.16, margin: 0,
      fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP
    });
  });

  // ── Bottom tag bar（透明 + 白框 + 白字）──
  const TAG_Y = 7.09, TAG_H = 0.20, TAG_W = 1.10, TAG_GAP = 0.10;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    col.tags.forEach((tag, t) => {
      const tx = cx + 0.08 + t * (TAG_W + TAG_GAP);
      sl.addShape("rect", {
        x: tx, y: TAG_Y, w: TAG_W, h: TAG_H,
        fill: { type: "none" }, line: { color: C.white, width: 1 }
      });
      sl.addText(tag, {
        x: tx, y: TAG_Y, w: TAG_W, h: TAG_H, margin: 0,
        fontFace: FONT, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle"
      });
    });
  });

  drawPageNum(sl, 2);
}

// ════════════════════════════════════════════════════════════
//  SLIDE 3 — 一頁式整合頁（附數據來源）
//  與 slide 2 內容完全一致，每段 problem/goal/insight body/action body
//  末尾加上淡灰色來源標籤 (P##) / 【策略推論】/【invos 服務】/【策略方向】
// ════════════════════════════════════════════════════════════
function slideOnePagerWithSources(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };

  const X0 = 0.20;
  const CW = 2.54;
  const GAP = 0.04;
  const TOTAL_W = CW * 5 + GAP * 4;

  const SUBTITLE_SOURCED = `${BRAND_NAME}｜通路客群行動建議【附數據來源】｜每段標示原始報告頁碼 (Pxx)、策略推論、invos 服務能力`;
  sl.addShape("rect", {
    x: X0, y: 0.10, w: TOTAL_W, h: 0.28,
    fill: { color: C.labelBar }, line: { color: C.cardEdge, width: 0.5 }
  });
  sl.addText(SUBTITLE_SOURCED, {
    x: X0 + 0.06, y: 0.10, w: TOTAL_W - 0.12, h: 0.28, margin: 0,
    fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle"
  });

  const HDR_Y = 0.46, HDR_H = 0.75;
  const TOP_STRIP_H = 0.05;
  const DIVIDER_Y_OFFSET = 0.46;
  const DIVIDER_H = 0.02;

  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    const src = SOURCES[i];

    sl.addShape("rect", { x: cx, y: HDR_Y, w: CW, h: HDR_H, fill: { type: "none" }, line: { color: p.edge, width: 0.8 } });
    sl.addShape("rect", { x: cx, y: HDR_Y, w: CW, h: TOP_STRIP_H, fill: { color: p.edge }, line: { color: p.edge, width: 0 } });
    sl.addShape("rect", { x: cx + 0.06, y: HDR_Y + DIVIDER_Y_OFFSET, w: CW - 0.14, h: DIVIDER_H, fill: { color: p.edge }, line: { color: p.edge, width: 0 } });

    // 標題行（不加來源；標題本身非數據）
    const titleRuns = [];
    if (col.verb) titleRuns.push({ text: col.verb, options: { fontSize: 10, bold: true, color: C.white } });
    titleRuns.push({ text: col.title, options: { fontSize: 10, bold: true, color: C.white } });
    if (col.subtitle) titleRuns.push({ text: " " + col.subtitle, options: { fontSize: 10, bold: true, color: C.white } });
    sl.addText(titleRuns, {
      x: cx + 0.08, y: HDR_Y + 0.06, w: CW - 0.16, h: 0.24, margin: 0,
      fontFace: FONT, valign: "middle"
    });

    // Problem + 來源
    sl.addText([
      { text: col.problem, options: { fontSize: 8, color: C.white } },
      { text: ` (${src.problem})`, options: { fontSize: 7, color: C.dim } }
    ], {
      x: cx + 0.08, y: HDR_Y + 0.28, w: CW - 0.16, h: 0.18, margin: 0,
      fontFace: FONT, valign: "middle"
    });

    // Goal + 來源
    const goalRuns = [{ text: "目標｜", options: { fontSize: 8, bold: true, color: C.white } }];
    col.goal.forEach(g => {
      if (g.label) goalRuns.push({ text: g.label, options: { fontSize: 8, color: C.white } });
      if (g.value) goalRuns.push({ text: g.value, options: { fontSize: 8, bold: true, color: C.white } });
    });
    goalRuns.push({ text: ` (${src.goal})`, options: { fontSize: 7, color: C.dim } });
    sl.addText(goalRuns, {
      x: cx + 0.08, y: HDR_Y + 0.50, w: CW - 0.16, h: 0.22, margin: 0,
      fontFace: FONT, valign: "middle"
    });
  });

  const KEY_LABEL_Y = 1.28;
  drawLabelBar(sl, X0, KEY_LABEL_Y, TOTAL_W, "關鍵資訊（附來源）");

  const KEY_Y = 1.64, KEY_H = 2.50;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    const src = SOURCES[i];
    sl.addShape("rect", { x: cx, y: KEY_Y, w: CW, h: KEY_H, fill: { type: "none" }, line: { color: C.cardEdge, width: 0.5 } });
    const runs = [];
    col.insights.forEach((insight, j) => {
      const isLast = (j === col.insights.length - 1);
      runs.push({ text: insight.head, options: { fontSize: 10, bold: true, color: p.insightText, breakLine: true } });
      runs.push({ text: insight.body, options: { fontSize: 8, color: C.white } });
      runs.push({
        text: ` (${src.insights[j].body})`,
        options: { fontSize: 7, color: C.dim, breakLine: !isLast, paraSpaceAfter: !isLast ? 4 : 0 }
      });
    });
    sl.addText(runs, {
      x: cx + 0.10, y: KEY_Y + 0.08, w: CW - 0.20, h: KEY_H - 0.16, margin: 0,
      fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP
    });
  });

  const ACT_LABEL_Y = 4.24;
  drawLabelBar(sl, X0, ACT_LABEL_Y, TOTAL_W, "行動建議（附來源）");

  const ACT_Y = 4.60, ACT_H = 2.45;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    const src = SOURCES[i];
    sl.addShape("rect", { x: cx, y: ACT_Y, w: CW, h: ACT_H, fill: { type: "none" }, line: { color: C.cardEdge, width: 0.5 } });
    const runs = [];
    col.actions.forEach((a, j) => {
      const isLast = (j === col.actions.length - 1);
      runs.push({ text: a.head, options: { bullet: true, fontSize: 8, bold: true, color: p.actionText } });
      runs.push({ text: a.body, options: { fontSize: 8, color: C.white } });
      runs.push({
        text: ` (${src.actions[j].body})`,
        options: { fontSize: 7, color: C.dim, breakLine: !isLast, paraSpaceAfter: !isLast ? 4 : 0 }
      });
    });
    sl.addText(runs, {
      x: cx + 0.10, y: ACT_Y + 0.08, w: CW - 0.20, h: ACT_H - 0.16, margin: 0,
      fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP
    });
  });

  const TAG_Y = 7.09, TAG_H = 0.20, TAG_W = 1.10, TAG_GAP = 0.10;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    col.tags.forEach((tag, t) => {
      const tx = cx + 0.08 + t * (TAG_W + TAG_GAP);
      sl.addShape("rect", {
        x: tx, y: TAG_Y, w: TAG_W, h: TAG_H,
        fill: { type: "none" }, line: { color: C.white, width: 1 }
      });
      sl.addText(tag, {
        x: tx, y: TAG_Y, w: TAG_W, h: TAG_H, margin: 0,
        fontFace: FONT, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle"
      });
    });
  });

  drawPageNum(sl, 3);
}

// ════════════════════════════════════════════════════════════
//  MAIN
// ════════════════════════════════════════════════════════════
async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = `${BRAND_NAME} 行動建議（${REPORT_DATE}）`;
  pres.author = "invosData";

  slideCover(pres);
  slideOnePager(pres);
  slideOnePagerWithSources(pres);

  await pres.writeFile({ fileName: OUT });
  console.log("✅  Saved:", OUT);
}

main().catch(e => { console.error(e); process.exit(1); });
