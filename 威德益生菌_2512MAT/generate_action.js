"use strict";
const pptxgen = require("pptxgenjs");
const path = require("path");

// ── BRAND CONFIG ─────────────────────────────────────────────
const BRAND_NAME  = "威德";
const CATEGORY    = "益生菌";
const REPORT_DATE = "2512MAT";
const SUBTITLE    = `${BRAND_NAME}｜通路客群行動建議｜當前瓶頸 x 核心機會 x invosData 數據解決方案`;
const COVER_BG_IMAGE = "";
const OUT = path.join(__dirname, `${BRAND_NAME}_行動建議_${REPORT_DATE}.pptx`);

// ─── Global Styles ───────────────────────────────────────────
const FONT = "Microsoft JhengHei";
const BODY_LINE_SP = 1.5;

// ─── Palette ─────────────────────────────────────────────────
const C = {
  bg:         "21253A",
  labelBar:   "202539",
  cardEdge:   "64748B",
  cashcow:    "2563EB",
  question:   "9333EA",
  star:       "F59E0B",
  dog:        "6B7280",
  white:      "FFFFFF",
  dim:        "94A3B8",
};

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
    verb:   "守住｜",
    title:  "好市多舊客基本盤",
    subtitle: "",
    problem:  "舊客穩健但品牌過度依賴單一客群",
    goal:   [
      { label: "忠誠度維持 ", value: "87%+" },
      { label: "、人均次對標市場 ", value: "1.8 次" },
    ],
    insights: [
      { head: "舊客+23%撐住銷額，行為穩健",
        body: "威德既有買者人數+23%、銷額+19%，貢獻57%銷額，舊客人均次+3%、單次金+5%，行為穩健。但品牌整體人均次-6%是被新客（-3%）和競品客（-6%）拖累，品牌過度依賴舊客撐盤" },
      { head: "忠誠度87.8%，但輕度買者流失",
        body: "好市多內威德買者忠誠度87.8%（僅12%兼買競品），基本盤穩固。但輕度買者銷額-5%，買者離開品類人數是流失至競品的5倍——最大威脅不是被搶走，是直接不買了" },
    ],
    actions: [
      { head: "補貨提醒：", body: "建立消耗週期計算，購買後3-4個月自動推送補貨提醒，目標人均年購次數從1.5次提升至對標市場1.8次" },
      { head: "品類喚醒：", body: "鎖定「已6個月未購買且無競品購買紀錄」的威德流失買者，推送品類需求內容召回，攔截品類離開" },
      { head: "舊客分層：", body: "RFM分層經營——高頻客推家庭備貨組合、中頻客推回購提醒、沉默客推品類科普喚醒" },
    ],
    tags: ["會員留存", "LTV深化"],
  },
  {
    preset: "question",
    verb:   "搶回｜",
    title:  "好市多新客入口",
    subtitle: "",
    problem:  "新客斷崖下滑，善存攔截源頭",
    goal:   [
      { label: "首購佔比回升至 ", value: "40%+" },
      { label: "（對標2024/05）", value: "" },
    ],
    insights: [
      { head: "新客-16%，首購佔比44.7%→30.8%",
        body: "威德新客人數YoY -16%（市場僅-4%），好市多檔期首購佔比從44.7%連續下滑至30.8%，檔期正從「開源機器」退化為「舊客補貨日」。新客人均次-3%、單次金-6%，投入金額也在下降" },
      { head: "善存攔截潛在新客水源",
        body: "善存86%新客是好市多舊客但前期沒買益生菌——這正是威德最需要的增量客群。流出至善存的威德買者中53%是前期新客，新客留存有漏洞" },
    ],
    actions: [
      { head: "搶先觸達：", body: "識別好市多會員中「有買保健品但沒買過益生菌」族群，檔期前2-4週投放威德品牌廣告，搶先善存建立心智佔據" },
      { head: "新客留存：", body: "首購後30/60/90天分階段投放留存廣告——使用提醒→回購優惠→檔期預告，目標首購回購率從18.9%提升至對標舊客水平" },
      { head: "檔期新客激勵：", body: "設計首購專屬小包裝體驗組，降低嘗試門檻，目標首購佔比回升至40%以上" },
    ],
    tags: ["新客導流", "精準受眾"],
  },
  {
    preset: "star",
    verb:   "開拓｜",
    title:  "電商第二戰場",
    subtitle: "",
    problem:  "96%銷額鎖死好市多，無備用戰場",
    goal:   [
      { label: "好市多以外佔比 ", value: ">10%" },
      { label: "（對標品類均值）", value: "" },
    ],
    insights: [
      { head: "通路集中度達危險水位",
        body: "威德96%銷額來自好市多，全家-76%，電商/藥妝幾乎無量。市場前十大品牌中通路最單一，好市多任何政策變動直接決定威德生死" },
      { head: "電商已有消費者基礎",
        body: "益生菌買者在momo併買以保健食品為主（魚油/鈣/葉黃素），消費者已有電商買保健品習慣。配方時代官網+226%證明電商可行，威德卻缺席" },
    ],
    actions: [
      { head: "momo旗艦佈局：", body: "以好市多同款產品的小包裝/試用組切入momo官方旗艦，降低電商首購門檻，建立「電商試用→好市多回購」跨通路路徑" },
      { head: "電商種草：", body: "鎖定momo/蝦皮上有益生菌購買紀錄但未買過威德的消費者，投放品牌教育廣告+試用優惠" },
    ],
    tags: ["新客導流", "精準受眾"],
  },
  {
    preset: "star",
    verb:   "深化｜",
    title:  "舊客回購頻率",
    subtitle: "",
    problem:  "半年一次的囤貨節奏無法建立黏著",
    goal:   [
      { label: "人均次 ", value: "1.8次" },
      { label: "（對標市場均值）", value: "" },
    ],
    insights: [
      { head: "非檔期月回購率僅1-3%",
        body: "威德消費集中在5月/10月檔期，中間半年幾乎不購買。非檔期月回購率僅1-3%，代表「囤貨型購買」天然不利建立品牌黏著度，消費者只在打折時出現" },
      { head: "整體市場重度買者人數膨脹但品質稀釋",
        body: "市場整體重度買者人數+13%但人均次-4%、單次金-5%，顯示中度買者升級進來拉低平均值。威德舊客本身穩健，但若新客持續萎縮，未來舊客自然流失後將無後備力量" },
    ],
    actions: [
      { head: "週期補貨機制：", body: "依上次購買時間推估食用完畢日，提前7天推送補貨提醒，將「半年囤一次」轉化為「每3-4個月補一次」，目標人均次從1.5次提升至1.8次" },
      { head: "中間檔體驗：", body: "在5月/10月大檔之間增設小型體驗檔，不靠折扣而是以新口味/新規格試用切入，擴大非檔期接觸面" },
    ],
    tags: ["會員留存", "LTV深化"],
  },
  {
    preset: "dog",
    verb:   "重構｜",
    title:  "檔期新客策略",
    subtitle: "",
    problem:  "放寬限購無效，需從拉深度轉向拉廣度",
    goal:   [
      { label: "檔期新客數回升至 ", value: "8,000+" },
      { label: "（對標2024/10）", value: "" },
    ],
    insights: [
      { head: "放寬限購盒數，人均盒數反降",
        body: "2025/10檔期放寬限購6→12盒，但人均盒數2.78（低於2025/05的2.90），7成消費者每檔不超過3盒——這是需求自然上限，不是限購造成的" },
      { head: "檔期銷額下滑的根因是人不夠",
        body: "2025/10總買者-11.5%、銷售額-15.3%，首購新客從44.7%降至30.8%。檔期策略的瓶頸不在「買的量不夠」而在「來的人不夠」" },
    ],
    actions: [
      { head: "新客首購激勵：", body: "將檔期部分資源轉移到首購激勵——首購專屬小包裝、好市多APP推播「有買保健品沒買過益生菌」會員，目標檔期新客數回升至8,000+人" },
      { head: "檔期效率追蹤：", body: "每檔期提供「首購數/回購率/喚回率/人均盒數」四維報告，取代看總銷額的粗放評估" },
    ],
    tags: ["新客導流", "通路監測"],
  },
];

// ════════════════════════════════════════════════════════════
//  SOURCES — Slide 3「附數據來源」
// ════════════════════════════════════════════════════════════
const SOURCES = [
  { // Col 1 守住好市多舊客基本盤
    problem: "P4, P60",
    goal: "P55（忠誠度87.8%）、P10（市場1.8次）",
    insights: [
      { body: "P38（舊客+23%、銷額+19%、舊客人均次+3%、單次金+5%）、P4（品牌整體人均次-6%）" },
      { body: "P55（忠誠度87.8%）、P60（輕度-5%）、P49（離開品類為流失競品5倍）" },
    ],
    actions: [
      { body: "P67（非檔期月回購1-3%）+ 【invos 服務】" },
      { body: "P49（離開品類人數）+ 【invos 服務】" },
      { body: "【invos 服務】+ 策略推論" },
    ],
  },
  { // Col 2 搶回好市多新客入口
    problem: "P38, P66",
    goal: "P66（2024/05首購佔比44.7%）【策略方向】",
    insights: [
      { body: "P38（新客-16%）、P13（市場-4%）、P66（44.7%→30.8%）" },
      { body: "P56（善存86%新客來源）、P52（53%流出新客）" },
    ],
    actions: [
      { body: "P56 + 【invos 服務】" },
      { body: "P81（回購率18.9%）+ 【invos 服務】" },
      { body: "【策略推論】" },
    ],
  },
  { // Col 3 開拓電商第二戰場
    problem: "P4, P33",
    goal: "【策略方向】",
    insights: [
      { body: "P4（96%好市多）、P22（全家-76%）、P33（通路佈局）" },
      { body: "P24（momo併買保健品）、P21（配方時代+226%）" },
    ],
    actions: [
      { body: "P24 + 【策略推論】" },
      { body: "【invos 服務】+ 策略推論" },
    ],
  },
  { // Col 4 深化舊客回購頻率
    problem: "P67, P80",
    goal: "P4（威德1.5次）、P10（市場1.8次）",
    insights: [
      { body: "P67（非檔期月回購率1-3%）" },
      { body: "P18（重度+13%、人均次-4%、單次金-5%）" },
    ],
    actions: [
      { body: "P67 + 【invos 服務】" },
      { body: "【策略推論】" },
    ],
  },
  { // Col 5 重構檔期新客策略
    problem: "P64, P65",
    goal: "P66（2024/10首購約8,190人）【策略方向】",
    insights: [
      { body: "P65（人均盒數2.78 vs 2.90、7成≤3盒）" },
      { body: "P64（買者-11.5%、銷額-15.3%）、P66（首購44.7%→30.8%）" },
    ],
    actions: [
      { body: "【策略推論】" },
      { body: "【invos 服務】" },
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
  sl.background = { color: C.bg };

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

  const X0 = 0.20;
  const CW = 2.54;
  const GAP = 0.04;
  const TOTAL_W = CW * 5 + GAP * 4;

  // Top title bar
  sl.addShape("rect", {
    x: X0, y: 0.10, w: TOTAL_W, h: 0.28,
    fill: { color: C.labelBar }, line: { color: C.cardEdge, width: 0.5 }
  });
  sl.addText(SUBTITLE, {
    x: X0 + 0.06, y: 0.10, w: TOTAL_W - 0.12, h: 0.28, margin: 0,
    fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle"
  });

  // Row 1: Header 目標卡
  const HDR_Y = 0.46, HDR_H = 0.75;
  const TOP_STRIP_H = 0.05;
  const DIVIDER_Y_OFFSET = 0.46;
  const DIVIDER_H = 0.02;

  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];

    sl.addShape("rect", { x: cx, y: HDR_Y, w: CW, h: HDR_H, fill: { type: "none" }, line: { color: p.edge, width: 0.8 } });
    sl.addShape("rect", { x: cx, y: HDR_Y, w: CW, h: TOP_STRIP_H, fill: { color: p.edge }, line: { color: p.edge, width: 0 } });
    sl.addShape("rect", { x: cx + 0.06, y: HDR_Y + DIVIDER_Y_OFFSET, w: CW - 0.14, h: DIVIDER_H, fill: { color: p.edge }, line: { color: p.edge, width: 0 } });

    const titleRuns = [];
    if (col.verb) titleRuns.push({ text: col.verb, options: { fontSize: 10, bold: true, color: C.white } });
    titleRuns.push({ text: col.title, options: { fontSize: 10, bold: true, color: C.white } });
    if (col.subtitle) titleRuns.push({ text: " " + col.subtitle, options: { fontSize: 10, bold: true, color: C.white } });
    sl.addText(titleRuns, {
      x: cx + 0.08, y: HDR_Y + 0.06, w: CW - 0.16, h: 0.24, margin: 0,
      fontFace: FONT, valign: "middle"
    });

    sl.addText(col.problem, {
      x: cx + 0.08, y: HDR_Y + 0.28, w: CW - 0.16, h: 0.18, margin: 0,
      fontFace: FONT, fontSize: 8, color: C.white, valign: "middle"
    });

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

  // Label bar: 關鍵資訊
  const KEY_LABEL_Y = 1.28;
  drawLabelBar(sl, X0, KEY_LABEL_Y, TOTAL_W, "關鍵資訊");

  // Row 2: 關鍵資訊卡
  const KEY_Y = 1.64, KEY_H = 2.50;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    sl.addShape("rect", { x: cx, y: KEY_Y, w: CW, h: KEY_H, fill: { type: "none" }, line: { color: C.cardEdge, width: 0.5 } });
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

  // Label bar: 行動建議
  const ACT_LABEL_Y = 4.24;
  drawLabelBar(sl, X0, ACT_LABEL_Y, TOTAL_W, "行動建議");

  // Row 3: 行動建議卡
  const ACT_Y = 4.60, ACT_H = 2.45;
  COLUMNS.forEach((col, i) => {
    const cx = X0 + i * (CW + GAP);
    const p = COL_PRESETS[col.preset];
    sl.addShape("rect", { x: cx, y: ACT_Y, w: CW, h: ACT_H, fill: { type: "none" }, line: { color: C.cardEdge, width: 0.5 } });
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

  // Bottom tag bar
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

    const titleRuns = [];
    if (col.verb) titleRuns.push({ text: col.verb, options: { fontSize: 10, bold: true, color: C.white } });
    titleRuns.push({ text: col.title, options: { fontSize: 10, bold: true, color: C.white } });
    if (col.subtitle) titleRuns.push({ text: " " + col.subtitle, options: { fontSize: 10, bold: true, color: C.white } });
    sl.addText(titleRuns, {
      x: cx + 0.08, y: HDR_Y + 0.06, w: CW - 0.16, h: 0.24, margin: 0,
      fontFace: FONT, valign: "middle"
    });

    sl.addText([
      { text: col.problem, options: { fontSize: 8, color: C.white } },
      { text: ` (${src.problem})`, options: { fontSize: 7, color: C.dim } }
    ], {
      x: cx + 0.08, y: HDR_Y + 0.28, w: CW - 0.16, h: 0.18, margin: 0,
      fontFace: FONT, valign: "middle"
    });

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
