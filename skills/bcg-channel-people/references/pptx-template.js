"use strict";
// ════════════════════════════════════════════════════════════
//  invosData BCG 通路佈局×客群配對 PPTX Template
//  使用方式：複製此檔案，修改下方 ── BRAND CONFIG ── 區塊
//  即可為任何品牌生成四頁通路佈局 BCG 策略投影片
// ════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const path = require("path");

// ── BRAND CONFIG ─────────────────────────────────────────────
// 修改以下變數，其餘程式碼保持不變
const BRAND_NAME  = "台塑生醫MD醫之方";  // 品牌名稱（用於標題與檔名）
const CATEGORY    = "益生菌";            // 品類名稱
const REPORT_DATE = "2512MAT";           // 報告期間標示
const COVER_BG_IMAGE = "";               // 封面背景圖片路徑（留空則用純色背景）
const OUT = path.join(__dirname, `../../mnt/outputs/${BRAND_NAME}_BCG策略建議_通路找人版.pptx`);

// ─── Global Styles ────────────────────────────────────────────
const FONT = "Microsoft JhengHei";     // 全域字型：微軟正黑體
const BODY_LINE_SP = 1.5;              // 全域內文行距倍數

// ─── Palette ───────────────────────────────────────────────
const C = {
  bg:         "21253A",
  bgPanel:    "1C2440",
  bgCard:     "232C4A",
  bgDark:     "0E1220",
  star:       "F59E0B",
  cashcow:    "2563EB",
  question:   "9333EA",
  dog:        "6B7280",
  insight:    "16A34A",
  media:      "0891B2",
  api:        "7C3AED",
  accent:     "F28165",
  green:      "3BA676",
  teal:       "008786",
  warn:       "EF4444",
  white:      "FFFFFF",
  offWhite:   "E2E8F0",
  muted:      "94A3B8",
  dim:        "64748B",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.25 });

// ── 投影片頁碼（左下角） ───────────────────────────────────────
function drawPageNum(sl, num) {
  sl.addText(String(num), {
    x: 0.15, y: 7.1, w: 0.4, h: 0.3, margin: 0,
    fontFace: FONT, fontSize: 8, color: C.white, align: "left", valign: "bottom"
  });
}

// ════════════════════════════════════════════════════════════
//  SLIDE 1 — 封面頁
// ════════════════════════════════════════════════════════════
function slideCover(pres) {
  const sl = pres.addSlide();

  // ── 背景：若有指定圖片則用圖片＋深色覆蓋層，否則用純色 ──
  if (COVER_BG_IMAGE) {
    sl.background = { path: COVER_BG_IMAGE };
    sl.addShape("rect", {
      x: 0, y: 0, w: 13.33, h: 7.5,
      fill: { color: C.bg, transparency: 25 }
    });
  } else {
    sl.background = { color: C.bg };
  }

  // ── 中央圖示：用 shapes 組合模擬長條圖 icon ──
  const iconCX = 6.65, iconCY = 2.8;
  const barW = 0.22, barGap = 0.08;
  const bars = [
    { h: 0.45, y: iconCY + 0.55 },  // short bar
    { h: 0.7,  y: iconCY + 0.3 },   // medium bar
    { h: 1.0,  y: iconCY },          // tall bar
  ];
  // Chart frame (L-shaped axes)
  sl.addShape("rect", { x: iconCX - 0.55, y: iconCY - 0.15, w: 0.06, h: 1.25, fill: { color: C.white }, line: { color: C.white, width: 0 } });
  sl.addShape("rect", { x: iconCX - 0.55, y: iconCY + 1.04, w: 1.6, h: 0.06, fill: { color: C.white }, line: { color: C.white, width: 0 } });
  // Bars
  bars.forEach((b, i) => {
    const bx = iconCX - 0.15 + i * (barW + barGap);
    sl.addShape("rect", { x: bx, y: b.y, w: barW, h: b.h, fill: { color: C.white }, line: { color: C.white, width: 0 } });
  });
  // Sparkle / star dot (top-right of tallest bar)
  sl.addText("✦", {
    x: iconCX + 0.45, y: iconCY - 0.35, w: 0.35, h: 0.35, margin: 0,
    fontFace: FONT, fontSize: 12, color: C.white, align: "center", valign: "middle"
  });

  // ── 標題文字 ──
  sl.addText("行動建議", {
    x: 2.5, y: 3.65, w: 8.3, h: 1.2, margin: 0,
    fontFace: FONT, fontSize: 36, bold: true, color: C.white, align: "center", valign: "middle"
  });

  drawPageNum(sl, 1);
}

// ════════════════════════════════════════════════════════════
//  SLIDE 2 — BCG 矩陣 + 策略定位說明（合併版）
// ════════════════════════════════════════════════════════════
function slide2(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText([
    { text: `${BRAND_NAME}${CATEGORY}｜通路佈局 × 客群配對策略矩陣`, options: { fontSize: 12, bold: true, color: C.white } },
    { text: `   ${REPORT_DATE}｜100+ 產品如何在對的通路找到對的人？`, options: { fontSize: 8, color: C.white } },
  ], { x: 0.3, y: 0.14, w: 12.7, h: 0.42, margin: 0, fontFace: FONT, valign: "middle" });

  const LX = 0.2, LY = 0.78;
  const QW = 3.45, QH = 3.2, GAP = 0.1;

  sl.addText("← 客群掌握度低（不知道誰在買）", { x: LX, y: LY - 0.22, w: QW, h: 0.2, margin: 0, align: "center", fontFace: FONT, fontSize: 8, color: C.white, italic: true });
  sl.addText("客群掌握度高（知道誰在買）→", { x: LX + QW + GAP, y: LY - 0.22, w: QW, h: 0.2, margin: 0, align: "center", fontFace: FONT, fontSize: 8, color: C.white, italic: true });

  // Y-axis labels (left side, vertical text)
  sl.addText("通路規模大但成長受限", { x: -0.08, y: LY + QH * 0.1, w: 0.3, h: QH * 0.8, margin: 0, fontFace: FONT, fontSize: 8, color: C.white, italic: true, align: "center", valign: "middle", isTextBox: true, vert: "eaVert" });
  sl.addText("通路規模小但潛力高", { x: -0.08, y: LY + QH + GAP + QH * 0.05, w: 0.3, h: QH * 0.9, margin: 0, fontFace: FONT, fontSize: 8, color: C.white, italic: true, align: "center", valign: "middle", isTextBox: true, vert: "eaVert" });

  // Dogs (top-left)
  sl.addShape("rect", { x: LX, y: LY, w: QW, h: QH, line: { color: C.dog, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: LX, y: LY, w: QW, h: 0.35, fill: { color: C.dog }, line: { color: C.dog, width: 0 } });
  sl.addText("🐶  酷澎：看得到人流，抓不住客人", { x: LX + 0.1, y: LY, w: QW - 0.2, h: 0.35, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle" });
  sl.addText([
    { text: "客數暴增但「找不到對的人」的典型", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "855人（+3.7倍），但 RPC 僅 1,160元（momo 的 46%）。82% 是品牌新客，被低價而非品牌力吸引——放再多產品也拉不高客單", options: { fontSize: 8, color: C.white, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "水貨通路無法做產品×客群配對", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "品牌無法控制上架品項、定價與陳列，100+ 產品的佈局策略在此通路完全失效，且低價認知正侵蝕品牌溢價", options: { fontSize: 8, color: C.white } },
  ], { x: LX + 0.12, y: LY + 0.42, w: QW - 0.22, h: QH - 0.5, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP });

  // Cash Cow (top-right)
  const CX = LX + QW + GAP;
  sl.addShape("rect", { x: CX, y: LY, w: QW, h: QH, line: { color: C.cashcow, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: CX, y: LY, w: QW, h: 0.35, fill: { color: C.cashcow }, line: { color: C.cashcow, width: 0 } });
  sl.addText("🐮  momo：最知道誰在買，但人正在走", { x: CX + 0.1, y: LY, w: QW - 0.2, h: 0.35, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle" });
  sl.addText([
    { text: "品牌「對的人」最集中的通路", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "RPC 2,517元全通路最高，舊客年均 $3,243、買者全市場均消 $3,679。這裡的人願意花錢、主動搜尋，是產品線深度佈局的首選陣地", options: { fontSize: 8, color: C.white, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "人在走：客數 -18.6%，但不是被搶走", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "流失 807 人中 69% 直接離開品類、僅 23% 轉買競品。問題是「買了一次就停」，不是產品不好而是沒人提醒他再買", options: { fontSize: 8, color: C.white } },
  ], { x: CX + 0.12, y: LY + 0.42, w: QW - 0.22, h: QH - 0.5, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP });

  // Question Marks (bottom-left)
  const QBY = LY + QH + GAP;
  sl.addShape("rect", { x: LX, y: QBY, w: QW, h: QH, line: { color: C.question, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: LX, y: QBY, w: QW, h: 0.35, fill: { color: C.question }, line: { color: C.question, width: 0 } });
  sl.addText("❓  藥妝通路：新人一直來，但都只買一次", { x: LX + 0.1, y: QBY, w: QW - 0.2, h: 0.35, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: C.white, valign: "middle" });
  sl.addText([
    { text: "康是美：天然的「品牌體驗入口」", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "312人（+32.8%），88% 新客。但回購率僅 11.9%——通路幫你找到了新人，問題是買完就散。RPC 1,317元，同群客全市場年均花 2,416元，你只拿到 55%", options: { fontSize: 8, color: C.white, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "寶雅：正在長但競品不同", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "份額升至 7%，競品是新普利/我的健康日記/大研生醫（非康是美的 BHK's/倍適）。不同通路的人在比較不同品牌——產品佈局不能複製貼上", options: { fontSize: 8, color: C.white } },
  ], { x: LX + 0.12, y: QBY + 0.42, w: QW - 0.22, h: QH - 0.5, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP });

  // Star (bottom-right)
  sl.addShape("rect", { x: CX, y: QBY, w: QW, h: QH, line: { color: C.star, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: CX, y: QBY, w: QW, h: 0.35, fill: { color: C.star }, line: { color: C.star, width: 0 } });
  sl.addText("⭐  舊客＋競品客：已知道是誰，能主動出擊", { x: CX + 0.1, y: QBY, w: QW - 0.2, h: 0.35, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: "1A1A1A", valign: "middle" });
  sl.addText([
    { text: "舊客：最值得被「找到」的人", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "人均 $2,800/年是新客 2 倍。但留存率僅 22%，60% 前期買者直接離開品類。100+ 產品的交叉銷售/產品線延伸對舊客最有效，但現在根本沒機制觸達他們", options: { fontSize: 8, color: C.white, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "競品客：可以被識別、被轉換", options: { bold: true, fontSize: 10, color: C.white, breakLine: true } },
    { text: "BHK's/善存為主要交叉購買品牌，發票數據可識別「同時買兩個品牌」的人——他們正在比較，你只要在對的時機推對的產品", options: { fontSize: 8, color: C.white } },
  ], { x: CX + 0.12, y: QBY + 0.42, w: QW - 0.22, h: QH - 0.5, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP });

  // BCG cross divider
  const MX = LX + QW + GAP / 2;
  const MY = LY + QH + GAP / 2;
  sl.addShape("rect", { x: MX - 0.018, y: LY, w: 0.036, h: QH * 2 + GAP, fill: { color: C.dim }, line: { color: C.dim, width: 0 } });
  sl.addShape("rect", { x: LX, y: MY - 0.018, w: QW * 2 + GAP, h: 0.036, fill: { color: C.dim }, line: { color: C.dim, width: 0 } });

  // ── RIGHT: STRATEGY POSITION BOXES (4 equal: STAR / QM / Cash Cow / Dogs) ──
  const RX = 7.6, RY = 0.78, RW = 5.55;
  const boxH4 = (QH * 2 + GAP - GAP * 3) / 4;

  sl.addText("策略定位說明", {
    x: RX, y: LY - 0.22, w: RW, h: 0.2, margin: 0,
    fontFace: FONT, fontSize: 10, bold: true, color: C.white, italic: true
  });

  const HDR = 0.27, DESC_Y = 0.30, DESC_H = 0.33, DIV_Y = 0.65, LBL_Y = 0.68, BULL_Y = 0.87;

  const stratBoxes = [
    {
      color: C.star, titleColor: "1A1A1A",
      title: "⭐  STAR — 讓舊客買更多、讓競品客改買你",
      desc:  "100+ 產品的最大價值在「已認識品牌的人」身上：舊客的交叉銷售、產品線延伸，以及競品客的精準轉換。這兩群人已可識別、可觸達。",
      labelColor: C.star,
      bullets: [
        "舊客留存率22%→30%：每提升1%等於新增一批$2,800/年高價值客，是100+產品最直接的銷售對象",
        "競品轉換：BHK's/善存交叉買者可識別，在對的通路推對的產品搶佔選牌決策",
      ]
    },
    {
      color: C.question, titleColor: C.white,
      title: "❓  Question — 藥妝找到新人，但需要橋樑",
      desc:  "康是美/寶雅幫品牌做了「品牌滲透」——88%新客證明通路能觸及新人。瓶頸是回購率12%，需要「首購→回購」的橋樑把新人導向線上。",
      labelColor: C.question,
      bullets: [
        "產品策略：藥妝放明星單品/試用裝→掃碼加會員→次購優惠導流momo/91APP，打通線下首購→線上留客",
        "通路差異化：康是美競品=BHK's/倍適；寶雅=新普利/我的健康日記——不同通路放不同產品",
      ]
    },
    {
      color: C.cashcow, titleColor: C.white,
      title: "🐮  Cash Cow — momo 是最該深耕的產品線陣地",
      desc:  "momo 買者 RPC 最高（2,517元）、消費力最強（全市場均消3,679元），是放「高單價組合包/進階功能訴求品項」的首選。問題是留不住人。",
      labelColor: C.white,
      bullets: [
        "產品線深度佈局：100+產品中高單價/組合包/進階功能品項集中在momo，匹配高消費力客群",
        "留客機制優先：流失主因是品類停用非競品搶奪，建立定期購+補貨提醒，留住一個舊客勝過獲五個新客",
      ]
    },
    {
      color: C.dog, titleColor: C.white,
      title: "🐶  Dogs — 酷澎有人流但無法佈局產品",
      desc:  "水貨通路無法控制品項、定價、陳列——100+ 產品的佈局策略在此完全失效。客數暴增但客單僅 1,160 元，正在培養「MD=便宜」的認知。",
      labelColor: "9CA3AF",
      bullets: [
        "不投入不競爭：酷澎82%新客被低價吸引，正面搶客只會拉低品牌溢價，策略重心放在主力通路留存",
        "監測+談判：用數據量化水貨侵蝕程度，作為通路保護與品牌定價策略的協商籌碼",
      ]
    },
  ];

  stratBoxes.forEach((b, i) => {
    const BY = RY + i * (boxH4 + GAP);
    sl.addShape("rect", { x: RX, y: BY, w: RW, h: boxH4, line: { color: b.color, width: 2 }, shadow: makeShadow() });
    sl.addShape("rect", { x: RX, y: BY, w: RW, h: HDR, fill: { color: b.color }, line: { color: b.color, width: 0 } });
    sl.addText(b.title, { x: RX + 0.1, y: BY, w: RW - 0.15, h: HDR, margin: 0, valign: "middle", fontFace: FONT, fontSize: 10, bold: true, color: b.titleColor });
    sl.addText(b.desc, { x: RX + 0.15, y: BY + DESC_Y, w: RW - 0.28, h: DESC_H, margin: 0, fontFace: FONT, fontSize: 8, color: C.white, wrap: true, lineSpacingMultiple: BODY_LINE_SP });
    sl.addShape("rect", { x: RX + 0.15, y: BY + DIV_Y, w: RW - 0.28, h: 0.022, fill: { color: b.color }, line: { color: b.color, width: 0 } });
    sl.addText("關鍵資訊", { x: RX + 0.15, y: BY + LBL_Y, w: RW - 0.28, h: 0.18, margin: 0, fontFace: FONT, fontSize: 8, bold: true, color: b.labelColor });
    sl.addText(
      b.bullets.map((t, j) => ({ text: t, options: { bullet: true, fontSize: 8, color: C.white, breakLine: j < b.bullets.length - 1, paraSpaceAfter: 3 } })),
      { x: RX + 0.15, y: BY + BULL_Y, w: RW - 0.28, h: boxH4 - BULL_Y - 0.04, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP }
    );
  });

  drawPageNum(sl, 2);
}

// ════════════════════════════════════════════════════════════
//  COLUMN CONFIG (shared by slides 3 & 4)
// ════════════════════════════════════════════════════════════
const COLS = [
  { label: "momo 產品線深耕陣地",  sub: "高單價 / 組合包 / 進階功能",     badge: "🐮 Cash Cow",       badgeColor: C.cashcow,  note: "RPC 2,517元 | 客數-18.6%" },
  { label: "康是美 品牌體驗入口",   sub: "明星單品 / 試用裝 / 首購導流",   badge: "❓ Question Mark",  badgeColor: C.question, note: "88%新客 | 回購率12%" },
  { label: "寶雅 第二藥妝佈點",    sub: "差異化品項 / 打不同競品",        badge: "❓ Question Mark",  badgeColor: C.question, note: "份額↑7% | 競品≠康是美" },
  { label: "舊客回購×產品延伸",    sub: "留存深化 / 交叉銷售 / 升級",     badge: "⭐ Star",            badgeColor: C.star,     note: "留存率22% | 人均$2,800" },
  { label: "競品客轉換×通路防守",   sub: "BHK's/善存/酷澎對策",          badge: "⭐ Star + 🐶 Dogs", badgeColor: C.star,     note: "可識別 | 可定向" },
];

function drawColHeaders(sl, X0, Y0, CW, full) {
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW, col = COLS[c];
    const hh = full ? 0.75 : 0.52;
    sl.addShape("rect", { x: cx, y: Y0, w: CW - 0.04, h: hh, fill: { color: C.bgPanel }, line: { color: col.badgeColor, width: 0.8 } });
    sl.addShape("rect", { x: cx, y: Y0, w: CW - 0.04, h: 0.05, fill: { color: col.badgeColor }, line: { color: col.badgeColor, width: 0 } });
    sl.addText(col.label, { x: cx + 0.06, y: Y0 + 0.06, w: CW - 0.16, h: 0.24, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: C.white });
    sl.addText(col.sub,   { x: cx + 0.06, y: Y0 + 0.28, w: CW - 0.16, h: 0.18, margin: 0, fontFace: FONT, fontSize: 8, color: C.white });
    if (full) {
      sl.addShape("rect", { x: cx + 0.06, y: Y0 + 0.46, w: CW - 0.18, h: 0.02, fill: { color: col.badgeColor }, line: { color: col.badgeColor, width: 0 } });
      sl.addText(col.badge, { x: cx + 0.06, y: Y0 + 0.5, w: CW * 0.66, h: 0.2, margin: 0, fontFace: FONT, fontSize: 8, bold: true, color: col.badgeColor });
      sl.addText(col.note,  { x: cx + 0.06, y: Y0 + 0.5, w: CW - 0.14, h: 0.2, margin: 0, align: "right", fontFace: FONT, fontSize: 8, color: C.white });
    }
  }
}

// ────────────────────────────────────────────────────────────
//  drawRow — data items can be plain strings OR [text, pageRef] tuples.
//  pageRef is rendered inline as small gray italic text after the bullet.
// ────────────────────────────────────────────────────────────
function drawRow(sl, label, color, bg, data, X0, rowY, CW, rowH) {
  const TW = CW * 5;
  sl.addShape("rect", { x: X0, y: rowY, w: TW, h: 0.27, fill: { color: bg }, line: { color, width: 0.5 } });
  sl.addText(label, { x: X0 + 0.1, y: rowY, w: TW - 0.2, h: 0.27, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color, valign: "middle" });
  const cY = rowY + 0.27;
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW;
    sl.addShape("rect", { x: cx, y: cY, w: CW - 0.04, h: rowH, line: { color: C.dim, width: 0.5 } });
    const bArr = [];
    data[c].forEach((item, i) => {
      const isLast = (i === data[c].length - 1);
      const [txt, ref] = Array.isArray(item) ? item : [item, null];
      if (ref) {
        bArr.push({ text: txt + "  ", options: { bullet: true, fontSize: 8, color: C.white } });
        bArr.push({ text: ref, options: {
          fontSize: 8, color: C.dim, italic: true,
          breakLine: !isLast,
          paraSpaceAfter: !isLast ? 3.5 : 0
        }});
      } else {
        bArr.push({ text: txt, options: {
          bullet: true, fontSize: 8, color: C.white,
          breakLine: !isLast,
          paraSpaceAfter: !isLast ? 3.5 : 0
        }});
      }
    });
    sl.addText(bArr, { x: cx + 0.08, y: cY + 0.09, w: CW - 0.2, h: rowH - 0.14, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP });
  }
}

// ════════════════════════════════════════════════════════════
//  SLIDE 3 — 當前瓶頸 × 核心機會
// ════════════════════════════════════════════════════════════
function slide3(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  sl.addText(`${BRAND_NAME}${CATEGORY}｜為什麼 100+ 產品還是找不到對的人？`, {
    x: 0.25, y: 0.1, w: 12.8, h: 0.36, margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: C.white
  });

  const X0 = 0.2, Y0 = 0.5, CW = (13.3 - X0 * 2) / 5;

  drawColHeaders(sl, X0, Y0, CW, true);

  // ════════════════════════════════════════════════════════
  //  ▼▼▼  每次換新報告時，修改以下兩個陣列  ▼▼▼
  //  格式：["句子內容", "報告 p.X"]  — 每條 bullet + 頁碼引用
  //  每欄 3 條，共 5 欄 × 2 列 = 30 條
  // ════════════════════════════════════════════════════════

  // ── 當前瓶頸 ──────────────────────────────────────────────
  const bottlenecks = [
    // Col 1: 藥局婦嬰通路整體
    [
      ["大樹/卡多摩/丁丁 三大主力通路購買人數均下滑，通路過度集中，缺乏備援增長引擎",  "報告 p.8"],
      ["消費者在實體通路購買後無任何品牌數位留存機制，每賣一罐等於流失一個可回購客戶",   "報告 p.26"],
      ["無法識別同一消費者的跨通路購買行為，三個通路的消費者資料各自孤立無法整合",       "報告 p.3"],
    ],
    // Col 2: 品類新客
    [
      ["卡洛塔妮銷售額 +71%（1,365萬）、人流 +16%，正快速截獲初次選牌的新手父母",       "報告 p.11"],
      ["美強生無針對「懷孕期/新生兒期父母」的精準數位觸達機制，第一次品牌接觸門檻高",   "報告 p.30"],
      ["品類新客在電商完成選牌比較，但美強生官方旗艦能見度低，在關鍵決策時刻無法曝光",   "報告 p.10, 16"],
    ],
    // Col 3: 品牌舊客 × 升級
    [
      ["優生→鉑睿升級路徑完全缺失：無 LINE 推播、無回購提醒、無差異說明，買家靠自主記憶", "報告 p.32, 35"],
      ["鉑睿買家人數持續下滑 + 單次金衰退，高端線雙重失血，但沒有優生舊客補進來",        "報告 p.32"],
      ["消費者品牌認知停留在「優生=美強生」，不知道有鉑睿此高端選項，且鉑睿電商幾乎無露出", "報告 p.16"],
    ],
    // Col 4: 競品轉換客
    [
      ["卡洛塔妮每月持續新增買家，部分來自原美強生客群，但美強生無任何機制追蹤這個流失",  "報告 p.13, 30"],
      ["競品客群一旦固化，轉換成本急劇升高；現在是轉換窗口，越晚介入越困難",             "報告 p.13"],
      ["美強生缺乏競品受眾識別與定向投放工具，既看不見流失，也沒有方式截獲",             "報告 p.34"],
    ],
    // Col 5: 電商第二通路
    [
      ["電商以酷澎最低單次金 NT$1,214 為主要露出，水貨疑慮已損害品牌定價秩序",           "報告 p.8, 16"],
      ["momo/蝦皮官方旗艦存在感極低，消費者電商搜尋時優先看到卡洛塔妮與雀巢",           "報告 p.15, 16"],
      ["實體舊客轉移至電商購買時，無跨通路識別機制，等於再次流失一次購買行為",           "報告 p.16"],
    ],
  ];

  // ── 核心機會 ──────────────────────────────────────────────
  const opps = [   // ← 同樣格式：["句子", "報告 p.X"]
    // Col 1: 藥局婦嬰通路整體
    [
      ["三大通路雖在下滑，仍是最大人流基礎；核心機會是將現有客「資產化」，從被動等待轉為主動留存", "報告 p.8, 26"],
      ["以發票整合跨通路購買行為，建立完整顧客輪廓，讓分散於三通路的消費者資料成為品牌資產",       "報告 p.3, 19"],
      ["競品尚未大舉布局實體通路，趁現在先行建立私域壁壘，是美強生的時間視窗",                   "報告 p.14"],
    ],
    // Col 2: 品類新客
    [
      ["以「購買尿布/奶瓶/月子產品但尚未購奶粉」的跨品類行為識別潛在新手父母，提前觸及",           "報告 p.19"],
      ["首購品牌黏著性極高（一旦選定通常使用至換階），比任何其他階段更值得投入行銷資源",           "報告 p.21, 22"],
      ["優生單次金 NT$931 門檻低，作為首購入口；首購加入會員後，後續升級路徑可自動化觸發",         "報告 p.35"],
    ],
    // Col 3: 品牌舊客 × 升級
    [
      ["優生1,677人購買次數持續增加，黏性已確認；若引導10%升鉑睿，每次消費金額提升3–4倍",         "報告 p.35"],
      ["親舒425人特配客群需求剛性強，計算補貨週期後自動提醒，無需折扣即可提升回購頻率",           "報告 p.32"],
      ["識別「優生3次以上購買者但從未接觸鉑睿」的精準升級受眾，是轉化成本最低的廣告目標群",       "報告 p.34, 35"],
    ],
    // Col 4: 競品轉換客
    [
      ["以發票識別「同時購買卡洛塔妮與優生」的交叉購買者——尚未品牌化的最高轉換潛力族群",         "報告 p.34"],
      ["美強生在藥局婦嬰通路的覆蓋深度，是識別競品交叉購買者的獨特優勢，競品完全沒有此洞察",     "報告 p.14, 34"],
      ["競品轉換窗口有限，越早建立識別與定向投放機制，截獲成本越低、轉換率越高",                 "報告 p.13"],
    ],
    // Col 5: 電商第二通路
    [
      ["momo 消費者主動搜尋意圖強，購買決策在電商完成，是截獲新客的最佳場景",                   "報告 p.7, 10"],
      ["電商可追蹤消費者數位行為，首購加入會員後，是建立美強生私域的高效入口",                   "報告 p.7"],
      ["非授權通路（酷澎）監測可量化灰色市場規模，作為保護官方定價的通路協商籌碼",               "報告 p.8, 15"],
    ],
  ];

  const rowY1 = Y0 + 0.78, rowH = 2.28;
  drawRow(sl, "為什麼找不到對的人？", C.warn, "1F0909", bottlenecks, X0, rowY1, CW, rowH);
  drawRow(sl, "如何用數據找到對的人", "22C55E", "091F09", opps, X0, rowY1 + 0.27 + rowH + 0.05, CW, rowH);

  drawPageNum(sl, 3);
}

// ════════════════════════════════════════════════════════════
//  SLIDE 4 — 行動建議 × invosData 方案
// ════════════════════════════════════════════════════════════
function slide4(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  sl.addText(`${BRAND_NAME}${CATEGORY}｜通路佈局行動方案 × invosData 幫你在每個通路找到對的人`, {
    x: 0.25, y: 0.1, w: 12.8, h: 0.36, margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: C.white
  });

  const X0 = 0.2, Y0 = 0.5, CW = (13.3 - X0 * 2) / 5;
  drawColHeaders(sl, X0, Y0, CW, false);

  const actions = [
    [
      "各通路包裝統一加入 QR Code，消費者掃碼上傳發票整合為同一份顧客資料，三通路買家合一管理",
      "建立促銷檔期前後廣告波段：促銷期衝新客首購，促銷結束後1個月精準推送舊客補貨提醒",
      "以發票監測三大通路人流衰退速度，提前識別高流失風險期，即時調整通路促銷排程",
    ],
    [
      "以跨品類發票行為（尿布/奶瓶/月子產品）識別潛在新手父母，在電商廣告投放優生首購優惠",
      "設計「新生兒禮包」：首購優生加贈鉑睿試用罐，一次完成品牌階梯教育，建立高端認知起點",
      "優先布局 momo/蝦皮官方旗艦，提升搜尋排名，讓初次選牌父母在電商能找到美強生",
    ],
    [
      "優生包裝附 QR Code 加入會員；第3次購買後自動觸發鉑睿升級禮方案推播（半價首盒）",
      "親舒特配依個人補貨週期，預估用罄前7天自動 LINE 推播補貨提醒，轉被動回購為主動留客",
      "圈選「優生3次以上×從未接觸鉑睿」精準受眾，在 Meta/Google 投放鉑睿差異化升級廣告",
    ],
    [
      "以發票識別「同時購買卡洛塔妮與優生」的交叉購買者，定向投放鉑睿轉換廣告",
      "在卡洛塔妮高滲透電商通路布局鉑睿比較廣告，以「成分科學差異」搶奪選牌決策時刻",
      "設計鉑睿30天試用計劃，以低門檻體驗取代高門檻說服，降低競品客轉換摩擦",
    ],
    [
      "優先建立 momo/蝦皮官方旗艦，鉑睿高端線主力曝光；電商獨家「優生→鉑睿升級包」提升客單",
      "圈選「藥局婦嬰通路購美強生但從未在電商購買」的消費者，在 momo 投放限時首購優惠廣告",
      "以 invos API 監測酷澎等非授權通路的美強生交易量，量化灰色市場侵蝕，作為通路保護談判工具",
    ],
  ];

  const actionTags = [
    [{ label: "舊客留存", color: C.cashcow }, { label: "通路管理", color: C.insight }],
    [{ label: "新客獲取", color: C.question }, { label: "品類教育", color: C.media }],
    [{ label: "舊客升級", color: C.star },    { label: "私域留存", color: C.api }],
    [{ label: "競品轉換", color: C.star },    { label: "新客獲取", color: C.media }],
    [{ label: "通路擴張", color: C.question },{ label: "通路保護", color: C.api }],
  ];

  const solutions = [
    { items: [
      { svc: "invos API",     color: C.api,     text: "跨通路會員整合：三大藥局婦嬰通路掃碼發票自動整合為一份顧客資料，跨通路購買行為統一管理，建立私域資產" },
      { svc: "invos Insight", color: C.insight, text: "實體通路月報：每月追蹤大樹/卡多摩/丁丁的銷售額與人流變化，量化通路衰退速度，提前預警與調整策略" },
      { svc: "invos Media",   color: C.media,   text: "通路舊客再行銷：圈選「特定通路購美強生×近期未回購」受眾，在促銷前後精準投放補貨提醒廣告" },
    ]},
    { items: [
      { svc: "invos Media",   color: C.media,   text: "跨品類精準新客：圈選「購買尿布/奶瓶/月子產品但未購奶粉」的潛在新手父母，在 Meta/Google 投放美強生首購廣告" },
      { svc: "invos Insight", color: C.insight, text: "品類新客洞察：追蹤嬰幼兒奶粉品類每月新客人數及首選品牌分佈，量化卡洛塔妮搶占速度，作為預算調配依據" },
      { svc: "invos API",     color: C.api,     text: "首購後私域整合：新客首購上傳發票加入美強生會員，啟動嬰兒成長陪伴序列（1階→2階→3階），自動留住升級流量" },
    ]},
    { items: [
      { svc: "invos API",     color: C.api,     text: "升級旅程自動化：優生第3次購買 → 鉑睿升級推播；親舒依消耗週期 → 補貨提醒，全程個人化私域經營" },
      { svc: "invos Media",   color: C.media,   text: "精準升級廣告：圈選「優生高頻買家×未購鉑睿」，投放鉑睿品質差異化廣告，量化廣告帶動升級的實際 ROAS" },
      { svc: "invos Insight", color: C.insight, text: "升級漏斗分析：追蹤優生→鉑睿轉化率，識別升級流失點，持續優化升級觸發機制與時機" },
    ]},
    { items: [
      { svc: "invos Insight", color: C.insight, text: "競品流失追蹤：量化卡洛塔妮買家中曾購美強生的比例，識別可轉換客規模，即時掌握競品擴張速度" },
      { svc: "invos Media",   color: C.media,   text: "競品受眾精準截獲：圈選「購卡洛塔妮×同期購優生」交叉買家，定向投放鉑睿轉換廣告，優先鎖定最高轉換潛力群" },
      { svc: "invos API",     color: C.api,     text: "轉換後私域留存：競品轉換首購後加入美強生私域，啟動鉑睿使用陪伴序列，建立新品牌黏性、阻止回流競品" },
    ]},
    { items: [
      { svc: "invos Insight", color: C.insight, text: "電商品牌月報：追蹤美強生在 momo/蝦皮/酷澎的銷售額、人流與競品電商滲透速度，動態調整電商定價策略" },
      { svc: "invos Media",   color: C.media,   text: "實體舊客×電商再行銷：圈選「藥局婦嬰通路購美強生但從未電商購買」族群，推播 momo 首購優惠，擴大跨通路觸及" },
      { svc: "invos API",     color: C.api,     text: "非授權通路偵測：監測酷澎等平台美強生交易量，追蹤灰色市場規模，量化水貨對官方定價與通路利益的侵蝕" },
    ]},
  ];

  let rowY = Y0 + 0.54;

  sl.addShape("rect", { x: X0, y: rowY, w: CW * 5, h: 0.27, fill: { color: "1A100A" }, line: { color: C.accent, width: 0.5 } });
  sl.addText("通路佈局行動方案", { x: X0 + 0.1, y: rowY, w: CW * 5 - 0.2, h: 0.27, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: C.accent, valign: "middle" });
  rowY += 0.27;

  const actionH = 2.08;
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW;
    sl.addShape("rect", { x: cx, y: rowY, w: CW - 0.04, h: actionH, line: { color: C.dim, width: 0.5 } });
    const bArr = actions[c].map((t, i) => ({
      text: t, options: { bullet: true, fontSize: 8, color: C.white, breakLine: i < actions[c].length - 1, paraSpaceAfter: 3 }
    }));
    sl.addText(bArr, { x: cx + 0.08, y: rowY + 0.08, w: CW - 0.2, h: actionH - 0.38, fontFace: FONT, valign: "top", lineSpacingMultiple: BODY_LINE_SP });
    const tags = actionTags[c];
    for (let t = 0; t < tags.length; t++) {
      sl.addShape("rect", { x: cx + 0.08 + t * 1.2, y: rowY + actionH - 0.26, w: 1.1, h: 0.2, fill: { color: tags[t].color }, line: { color: tags[t].color, width: 0 } });
      sl.addText(tags[t].label, { x: cx + 0.08 + t * 1.2, y: rowY + actionH - 0.26, w: 1.1, h: 0.2, margin: 0, align: "center", valign: "middle", fontFace: FONT, fontSize: 8, bold: true, color: C.white });
    }
  }
  rowY += actionH + 0.04;

  sl.addShape("rect", { x: X0, y: rowY, w: CW * 5, h: 0.27, fill: { color: "090F0D" }, line: { color: C.insight, width: 0.5 } });
  sl.addText("▶  invosData 幫你在每個通路找到對的人", { x: X0 + 0.1, y: rowY, w: CW * 5 - 0.2, h: 0.27, margin: 0, fontFace: FONT, fontSize: 10, bold: true, color: "4ADE80", valign: "middle" });
  rowY += 0.27;

  const svcH = 3.12;
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW;
    sl.addShape("rect", { x: cx, y: rowY, w: CW - 0.04, h: svcH, line: { color: C.dim, width: 0.5 } });
    let iy = rowY + 0.1;
    for (const item of solutions[c].items) {
      sl.addShape("rect", { x: cx + 0.07, y: iy, w: 1.55, h: 0.2, fill: { color: item.color }, line: { color: item.color, width: 0 } });
      sl.addText("▶ " + item.svc, { x: cx + 0.07, y: iy, w: 1.55, h: 0.2, margin: 0, align: "center", valign: "middle", fontFace: FONT, fontSize: 8, bold: true, color: C.white });
      sl.addText(item.text, { x: cx + 0.07, y: iy + 0.22, w: CW - 0.18, h: 0.68, margin: 0, fontFace: FONT, fontSize: 8, color: C.white, wrap: true, lineSpacingMultiple: BODY_LINE_SP });
      iy += 0.95;
    }
  }

  drawPageNum(sl, 4);
}

// ════════════════════════════════════════════════════════════
//  MAIN
// ════════════════════════════════════════════════════════════
async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = `${BRAND_NAME} BCG策略建議（通路佈局×客群配對版）`;
  pres.author = "invosData";

  slideCover(pres);
  slide2(pres);
  slide3(pres);
  slide4(pres);

  await pres.writeFile({ fileName: OUT });
  console.log("✅  Saved:", OUT);
}

main().catch(e => { console.error(e); process.exit(1); });
