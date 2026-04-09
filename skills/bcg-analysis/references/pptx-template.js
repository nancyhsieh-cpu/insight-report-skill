"use strict";
// ════════════════════════════════════════════════════════════
//  invosData BCG 策略矩陣 PPTX Template
//  使用方式：複製此檔案，修改下方 ── BRAND CONFIG ── 區塊
//  即可為任何品牌生成四頁 BCG 策略投影片
// ════════════════════════════════════════════════════════════
const pptxgen = require("pptxgenjs");
const path = require("path");

// ── BRAND CONFIG ─────────────────────────────────────────────
// 修改以下三個變數，其餘程式碼保持不變
const BRAND_NAME  = "美強生";          // 品牌名稱（用於標題與檔名）
const REPORT_DATE = "2602MAT";         // 報告期間標示
const OUT = path.join(__dirname, `../../mnt/outputs/${BRAND_NAME}_BCG策略建議_通路新舊客版.pptx`);

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
  teal:       "008786",
  red:        "C33B0E",
  warn:       "EF4444",
  white:      "FFFFFF",
  offWhite:   "E2E8F0",
  muted:      "94A3B8",
  dim:        "64748B",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.25 });

// ── 三色頂部色條（所有投影片共用） ─────────────────────────────
// 橘（左 60%）+ 青（中 20%）+ 紅（右 20%），總寬 13.3"
function drawTopBar(sl) {
  const W = 13.3;
  sl.addShape("rect", { x: 0,         y: 0, w: W * 0.6, h: 0.08, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  sl.addShape("rect", { x: W * 0.6,   y: 0, w: W * 0.2, h: 0.08, fill: { color: C.teal },   line: { color: C.teal, width: 0 } });
  sl.addShape("rect", { x: W * 0.8,   y: 0, w: W * 0.2, h: 0.08, fill: { color: C.red },    line: { color: C.red, width: 0 } });
}

// ── 投影片頁碼（左下角） ───────────────────────────────────────
function drawPageNum(sl, num) {
  sl.addText(String(num), {
    x: 0.15, y: 7.1, w: 0.4, h: 0.3, margin: 0,
    fontSize: 9, color: C.dim, align: "left", valign: "bottom"
  });
}

// ════════════════════════════════════════════════════════════
//  SLIDE 1 — 封面頁
// ════════════════════════════════════════════════════════════
function slideCover(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };

  drawTopBar(sl);

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
    fontSize: 14, color: C.white, align: "center", valign: "middle"
  });

  // ── 標題文字 ──
  sl.addText("行動建議", {
    x: 2.5, y: 3.65, w: 8.3, h: 1.2, margin: 0,
    fontSize: 36, bold: true, color: C.white, align: "center", valign: "middle"
  });

  drawPageNum(sl, 1);
}

// ════════════════════════════════════════════════════════════
//  SLIDE 2 — BCG 矩陣 + 策略定位說明（合併版）
// ════════════════════════════════════════════════════════════
function slide2(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };

  drawTopBar(sl);

  sl.addText(`${BRAND_NAME}嬰幼兒奶粉｜BCG 策略矩陣（通路 × 新舊客視角）`, {
    x: 0.3, y: 0.14, w: 8.8, h: 0.38, margin: 0,
    fontSize: 15, bold: true, color: C.white
  });
  sl.addText(`${REPORT_DATE}｜invosData 引客數據`, {
    x: 0.3, y: 0.5, w: 8.8, h: 0.24, margin: 0, fontSize: 9, color: C.muted
  });

  const LX = 0.2, LY = 0.78;
  const QW = 3.45, QH = 3.2, GAP = 0.1;

  sl.addText("← 掌握度低", { x: LX, y: LY - 0.22, w: QW, h: 0.2, margin: 0, align: "center", fontSize: 7.5, color: C.dim, italic: true });
  sl.addText("掌握度高 →", { x: LX + QW + GAP, y: LY - 0.22, w: QW, h: 0.2, margin: 0, align: "center", fontSize: 7.5, color: C.dim, italic: true });

  // Y-axis labels (left side, vertical text simulated)
  sl.addText("規模大但發展受限 ↑", { x: -0.1, y: LY + QH * 0.3, w: 0.35, h: 1.2, margin: 0, fontSize: 7, color: C.dim, italic: true, rotate: 270 });
  sl.addText("↑ 規模小但發展潛力高（進攻）", { x: -0.1, y: LY + QH + GAP + QH * 0.2, w: 0.35, h: 1.5, margin: 0, fontSize: 7, color: C.dim, italic: true, rotate: 270 });

  // Dogs (top-left)
  sl.addShape("rect", { x: LX, y: LY, w: QW, h: QH, fill: { color: "171C2A" }, line: { color: C.dog, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: LX, y: LY, w: QW, h: 0.35, fill: { color: C.dog }, line: { color: C.dog, width: 0 } });
  sl.addText("🐶  Dogs", { x: LX + 0.1, y: LY, w: QW - 0.2, h: 0.35, margin: 0, fontSize: 11, bold: true, color: C.white, valign: "middle" });
  sl.addText([
    { text: "鉑睿電商缺席", options: { bold: true, fontSize: 10, color: "D1D5DB", breakLine: true } },
    { text: "消費者電商選牌時找不到鉑睿，直接流向卡洛塔妮", options: { fontSize: 8.5, color: C.offWhite, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "實體通路雙重擠壓", options: { bold: true, fontSize: 10, color: "D1D5DB", breakLine: true } },
    { text: "大樹/卡多摩/丁丁 三大通路人流同步下滑，通路集中依賴風險高", options: { fontSize: 8.5, color: C.offWhite } },
  ], { x: LX + 0.12, y: LY + 0.42, w: QW - 0.22, h: QH - 0.5, valign: "top" });

  // Cash Cow (top-right)
  const CX = LX + QW + GAP;
  sl.addShape("rect", { x: CX, y: LY, w: QW, h: QH, fill: { color: "101825" }, line: { color: C.cashcow, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: CX, y: LY, w: QW, h: 0.35, fill: { color: C.cashcow }, line: { color: C.cashcow, width: 0 } });
  sl.addText("🐮  Cash Cow", { x: CX + 0.1, y: LY, w: QW - 0.2, h: 0.35, margin: 0, fontSize: 11, bold: true, color: C.white, valign: "middle" });
  sl.addText([
    { text: "藥局婦嬰既有客", options: { bold: true, fontSize: 10, color: "93C5FD", breakLine: true } },
    { text: "三大通路仍是最大人流基礎，需從「等客上門」轉為主動留存", options: { fontSize: 8.5, color: C.offWhite, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "優生一階穩定回購", options: { bold: true, fontSize: 10, color: "93C5FD", breakLine: true } },
    { text: "1,677人・購買次數與人數雙增，最穩定現金流，但停留在 NT$931 低客單", options: { fontSize: 8.5, color: C.offWhite } },
  ], { x: CX + 0.12, y: LY + 0.42, w: QW - 0.22, h: QH - 0.5, valign: "top" });

  // Question Marks (bottom-left)
  const QBY = LY + QH + GAP;
  sl.addShape("rect", { x: LX, y: QBY, w: QW, h: QH, fill: { color: "130E22" }, line: { color: C.question, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: LX, y: QBY, w: QW, h: 0.35, fill: { color: C.question }, line: { color: C.question, width: 0 } });
  sl.addText("❓  Question Marks", { x: LX + 0.1, y: QBY, w: QW - 0.2, h: 0.35, margin: 0, fontSize: 11, bold: true, color: C.white, valign: "middle" });
  sl.addText([
    { text: "品類新客攻防", options: { bold: true, fontSize: 10, color: "C4B5FD", breakLine: true } },
    { text: "卡洛塔妮 +71% 正搶佔新手父母，美強生無主動觸達新客的機制", options: { fontSize: 8.5, color: C.offWhite, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "電商第二通路", options: { bold: true, fontSize: 10, color: "C4B5FD", breakLine: true } },
    { text: "momo/蝦皮官方旗艦存在感低，選牌時消費者優先看到競品", options: { fontSize: 8.5, color: C.offWhite } },
  ], { x: LX + 0.12, y: QBY + 0.42, w: QW - 0.22, h: QH - 0.5, valign: "top" });

  // Star (bottom-right)
  sl.addShape("rect", { x: CX, y: QBY, w: QW, h: QH, fill: { color: "1A1208" }, line: { color: C.star, width: 1.5 }, shadow: makeShadow() });
  sl.addShape("rect", { x: CX, y: QBY, w: QW, h: 0.35, fill: { color: C.star }, line: { color: C.star, width: 0 } });
  sl.addText("⭐  Star", { x: CX + 0.1, y: QBY, w: QW - 0.2, h: 0.35, margin: 0, fontSize: 11, bold: true, color: "1A1A1A", valign: "middle" });
  sl.addText([
    { text: "品牌舊客升級路徑", options: { bold: true, fontSize: 10, color: "FDE68A", breakLine: true } },
    { text: "優生1,677人黏性已確認，升級機制一建立即可解鎖 LTV 3–4 倍提升", options: { fontSize: 8.5, color: C.offWhite, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "競品轉換客（交叉購買者）", options: { bold: true, fontSize: 10, color: "FDE68A", breakLine: true } },
    { text: "發票可識別「卡洛塔妮×優生同期購買者」，美強生獨有競爭情報優勢", options: { fontSize: 8.5, color: C.offWhite } },
  ], { x: CX + 0.12, y: QBY + 0.42, w: QW - 0.22, h: QH - 0.5, valign: "top" });

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
    fontSize: 9.5, bold: true, color: C.muted, italic: true
  });

  const HDR = 0.27, DESC_Y = 0.30, DESC_H = 0.33, DIV_Y = 0.65, LBL_Y = 0.68, BULL_Y = 0.87;

  const stratBoxes = [
    {
      color: C.star, titleColor: "1A1A1A",
      title: "⭐  STAR（立即可解鎖的成長動能）",
      desc:  "黏性客群與可識別競品交叉購買者已存在，差的只是主動介入機制。工具一到位即可短期變現潛在價值。",
      labelColor: C.star,
      bullets: [
        "品牌舊客升級：優生1,677人黏性確認，引導10%升鉑睿 → 每筆消費翻3–4倍",
        "競品轉換：發票可識別「卡洛塔妮×優生交叉購買者」，美強生獨有競爭情報",
      ]
    },
    {
      color: C.question, titleColor: C.white,
      title: "❓  Question Mark（主動介入才能解鎖）",
      desc:  "機會確實存在，但缺乏電商基礎與主動觸達機制。競品正快速填補空白，若不儘速建立將持續流失。",
      labelColor: C.question,
      bullets: [
        "品類新客：卡洛塔妮+71%搶佔新生兒父母，美強生無主動數位觸達新手父母的機制",
        "電商通路：三階市場美強生佔比<2%，momo/蝦皮旗艦存在感低，品牌缺席決策點",
      ]
    },
    {
      color: C.cashcow, titleColor: C.white,
      title: "🐮  Cash Cow（守住基本盤，驅動升級）",
      desc:  "三大藥局婦嬰通路仍是最大人流基礎，但客群正被動流失中。轉為主動留存並引導舊客升級是核心任務。",
      labelColor: "93C5FD",
      bullets: [
        "主動留客：購後無留存機制，建立掃碼/加會員流程即可截住每一筆購買行為",
        "客單升級：優生回購客引導升鉑睿，是成本最低、LTV 提升最快的槓桿路徑",
      ]
    },
    {
      color: C.dog, titleColor: C.white,
      title: "🐶  Dogs（止血為先，勿過度投入）",
      desc:  "電商缺席加上實體通路人流下滑，核心任務是防止流血擴大，而非逆勢加碼投資。",
      labelColor: "9CA3AF",
      bullets: [
        "電商補位：鉑睿在 momo/蝦皮幾乎缺席，補上官方旗艦基本存在感可截斷消費者外流",
        "通路防守：三大藥局婦嬰通路人流下滑中，趁客群尚在時主動「資產化」，阻止隱形流失",
      ]
    },
  ];

  stratBoxes.forEach((b, i) => {
    const BY = RY + i * (boxH4 + GAP);
    sl.addShape("rect", { x: RX, y: BY, w: RW, h: boxH4, fill: { color: C.bgCard }, line: { color: b.color, width: 2 }, shadow: makeShadow() });
    sl.addShape("rect", { x: RX, y: BY, w: RW, h: HDR, fill: { color: b.color }, line: { color: b.color, width: 0 } });
    sl.addText(b.title, { x: RX + 0.1, y: BY, w: RW - 0.15, h: HDR, margin: 0, valign: "middle", fontSize: 10.5, bold: true, color: b.titleColor });
    sl.addText(b.desc, { x: RX + 0.15, y: BY + DESC_Y, w: RW - 0.28, h: DESC_H, margin: 0, fontSize: 8.5, color: C.offWhite, wrap: true });
    sl.addShape("rect", { x: RX + 0.15, y: BY + DIV_Y, w: RW - 0.28, h: 0.022, fill: { color: b.color }, line: { color: b.color, width: 0 } });
    sl.addText("關鍵資訊", { x: RX + 0.15, y: BY + LBL_Y, w: RW - 0.28, h: 0.18, margin: 0, fontSize: 8.5, bold: true, color: b.labelColor });
    sl.addText(
      b.bullets.map((t, j) => ({ text: t, options: { bullet: true, fontSize: 8.5, color: C.offWhite, breakLine: j < b.bullets.length - 1, paraSpaceAfter: 3 } })),
      { x: RX + 0.15, y: BY + BULL_Y, w: RW - 0.28, h: boxH4 - BULL_Y - 0.04, valign: "top" }
    );
  });

  drawPageNum(sl, 2);
}

// ════════════════════════════════════════════════════════════
//  COLUMN CONFIG (shared by slides 3 & 4)
// ════════════════════════════════════════════════════════════
const COLS = [
  { label: "藥局婦嬰通路整體",   sub: "大樹 / 卡多摩 / 丁丁",        badge: "🐮 Cash Cow",       badgeColor: C.cashcow,  note: "最大客群 | 通路人流下滑中" },
  { label: "品類新客",         sub: "新生兒父母・首次選牌",           badge: "❓ Question Mark",  badgeColor: C.question, note: "卡洛塔妮+71%搶先" },
  { label: "品牌舊客 × 升級",  sub: "優生→鉑睿・留客深化消費",       badge: "⭐ Star",            badgeColor: C.star,     note: "升級機制缺失 | 潛力最大" },
  { label: "競品轉換客",       sub: "卡洛塔妮/雀巢→美強生",          badge: "⭐ Star",            badgeColor: C.star,     note: "可識別受眾 | 高轉換潛力" },
  { label: "電商第二通路",      sub: "momo / 蝦皮官方旗艦",           badge: "❓ Question Mark",  badgeColor: C.question, note: "高潛力 | 現況薄弱" },
];

function drawColHeaders(sl, X0, Y0, CW, full) {
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW, col = COLS[c];
    const hh = full ? 0.75 : 0.52;
    sl.addShape("rect", { x: cx, y: Y0, w: CW - 0.04, h: hh, fill: { color: C.bgPanel }, line: { color: col.badgeColor, width: 0.8 } });
    sl.addShape("rect", { x: cx, y: Y0, w: CW - 0.04, h: 0.05, fill: { color: col.badgeColor }, line: { color: col.badgeColor, width: 0 } });
    sl.addText(col.label, { x: cx + 0.06, y: Y0 + 0.06, w: CW - 0.16, h: 0.24, margin: 0, fontSize: 10.5, bold: true, color: C.white });
    sl.addText(col.sub,   { x: cx + 0.06, y: Y0 + 0.28, w: CW - 0.16, h: 0.18, margin: 0, fontSize: 8.5, color: C.muted });
    if (full) {
      sl.addShape("rect", { x: cx + 0.06, y: Y0 + 0.46, w: CW - 0.18, h: 0.02, fill: { color: col.badgeColor }, line: { color: col.badgeColor, width: 0 } });
      sl.addText(col.badge, { x: cx + 0.06, y: Y0 + 0.5, w: CW * 0.66, h: 0.2, margin: 0, fontSize: 8.5, bold: true, color: col.badgeColor });
      sl.addText(col.note,  { x: cx + 0.06, y: Y0 + 0.5, w: CW - 0.14, h: 0.2, margin: 0, align: "right", fontSize: 7.5, color: C.dim });
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
  sl.addText(label, { x: X0 + 0.1, y: rowY, w: TW - 0.2, h: 0.27, margin: 0, fontSize: 10, bold: true, color, valign: "middle" });
  const cY = rowY + 0.27;
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW;
    sl.addShape("rect", { x: cx, y: cY, w: CW - 0.04, h: rowH, fill: { color: C.bgCard }, line: { color: C.dim, width: 0.5 } });
    const bArr = [];
    data[c].forEach((item, i) => {
      const isLast = (i === data[c].length - 1);
      const [txt, ref] = Array.isArray(item) ? item : [item, null];
      if (ref) {
        // bullet text run (no breakLine yet — the ref run handles paragraph break)
        bArr.push({ text: txt + "  ", options: { bullet: true, fontSize: 9, color: C.offWhite } });
        // inline page ref — small, gray, italic; breakLine separates paragraphs
        bArr.push({ text: ref, options: {
          fontSize: 7, color: "64748B", italic: true,
          breakLine: !isLast,
          paraSpaceAfter: !isLast ? 3.5 : 0
        }});
      } else {
        bArr.push({ text: txt, options: {
          bullet: true, fontSize: 9, color: C.offWhite,
          breakLine: !isLast,
          paraSpaceAfter: !isLast ? 3.5 : 0
        }});
      }
    });
    sl.addText(bArr, { x: cx + 0.08, y: cY + 0.09, w: CW - 0.2, h: rowH - 0.14, valign: "top" });
  }
}

// ════════════════════════════════════════════════════════════
//  SLIDE 3 — 當前瓶頸 × 核心機會
// ════════════════════════════════════════════════════════════
function slide3(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  drawTopBar(sl);
  sl.addText(`${BRAND_NAME}嬰幼兒奶粉｜當前瓶頸 × 核心機會`, {
    x: 0.25, y: 0.1, w: 12.8, h: 0.36, margin: 0, fontSize: 13, bold: true, color: C.white
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
  drawRow(sl, "當前瓶頸", C.warn, "1F0909", bottlenecks, X0, rowY1, CW, rowH);
  drawRow(sl, "核心機會", "22C55E", "091F09", opps, X0, rowY1 + 0.27 + rowH + 0.05, CW, rowH);

  drawPageNum(sl, 3);
}

// ════════════════════════════════════════════════════════════
//  SLIDE 4 — 行動建議 × invosData 方案
// ════════════════════════════════════════════════════════════
function slide4(pres) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  drawTopBar(sl);
  sl.addText(`${BRAND_NAME}嬰幼兒奶粉｜行動建議 × invosData 數據解決方案`, {
    x: 0.25, y: 0.1, w: 12.8, h: 0.36, margin: 0, fontSize: 13, bold: true, color: C.white
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
  sl.addText("行動建議", { x: X0 + 0.1, y: rowY, w: CW * 5 - 0.2, h: 0.27, margin: 0, fontSize: 10, bold: true, color: C.accent, valign: "middle" });
  rowY += 0.27;

  const actionH = 2.08;
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW;
    sl.addShape("rect", { x: cx, y: rowY, w: CW - 0.04, h: actionH, fill: { color: C.bgCard }, line: { color: C.dim, width: 0.5 } });
    const bArr = actions[c].map((t, i) => ({
      text: t, options: { bullet: true, fontSize: 9, color: C.offWhite, breakLine: i < actions[c].length - 1, paraSpaceAfter: 3 }
    }));
    sl.addText(bArr, { x: cx + 0.08, y: rowY + 0.08, w: CW - 0.2, h: actionH - 0.38, valign: "top" });
    const tags = actionTags[c];
    for (let t = 0; t < tags.length; t++) {
      sl.addShape("rect", { x: cx + 0.08 + t * 1.2, y: rowY + actionH - 0.26, w: 1.1, h: 0.2, fill: { color: tags[t].color }, line: { color: tags[t].color, width: 0 } });
      sl.addText(tags[t].label, { x: cx + 0.08 + t * 1.2, y: rowY + actionH - 0.26, w: 1.1, h: 0.2, margin: 0, align: "center", valign: "middle", fontSize: 7.5, bold: true, color: C.white });
    }
  }
  rowY += actionH + 0.04;

  sl.addShape("rect", { x: X0, y: rowY, w: CW * 5, h: 0.27, fill: { color: "090F0D" }, line: { color: C.insight, width: 0.5 } });
  sl.addText("▶  invosData 數據解決方案", { x: X0 + 0.1, y: rowY, w: CW * 5 - 0.2, h: 0.27, margin: 0, fontSize: 10, bold: true, color: "4ADE80", valign: "middle" });
  rowY += 0.27;

  const svcH = 3.12;
  for (let c = 0; c < 5; c++) {
    const cx = X0 + c * CW;
    sl.addShape("rect", { x: cx, y: rowY, w: CW - 0.04, h: svcH, fill: { color: C.bgDark }, line: { color: C.dim, width: 0.5 } });
    let iy = rowY + 0.1;
    for (const item of solutions[c].items) {
      sl.addShape("rect", { x: cx + 0.07, y: iy, w: 1.55, h: 0.2, fill: { color: item.color }, line: { color: item.color, width: 0 } });
      sl.addText("▶ " + item.svc, { x: cx + 0.07, y: iy, w: 1.55, h: 0.2, margin: 0, align: "center", valign: "middle", fontSize: 7.5, bold: true, color: C.white });
      sl.addText(item.text, { x: cx + 0.07, y: iy + 0.22, w: CW - 0.18, h: 0.68, margin: 0, fontSize: 8.5, color: C.offWhite, wrap: true });
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
  pres.title = `${BRAND_NAME} BCG策略建議（通路×新舊客版）`;
  pres.author = "invosData";

  slideCover(pres);
  slide2(pres);
  slide3(pres);
  slide4(pres);

  await pres.writeFile({ fileName: OUT });
  console.log("✅  Saved:", OUT);
}

main().catch(e => { console.error(e); process.exit(1); });
