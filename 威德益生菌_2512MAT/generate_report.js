"use strict";
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, WidthType, BorderStyle,
  ShadingType, VerticalAlign
} = require("docx");
const fs = require("fs");
const path = require("path");

// ── REPORT CONFIG ─────────────────────────────────────────────
const BRAND      = "威德";
const CATEGORY   = "益生菌";
const PERIOD     = "2512MAT";
const DATE_RANGE = "2025/01 ～ 2025/12";
const DATE_COMP  = "2024/01 ～ 2024/12";
const SAMPLE_N   = "306,000";
const OUT = path.join(__dirname, `${BRAND}${CATEGORY}市場洞察分析報告摘要_invos.docx`);

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

// ── Source citation helper ─────────────────────────────────────
const srcRun = (src) => new TextRun({
  text: ` (${src})`, font: "Arial", size: 18, italics: true, color: "888888"
});
const PSRC = (text, src, opts = {}) => new Paragraph({
  children: [
    new TextRun({ text, font: "Arial", size: opts.size || 22, ...opts }),
    srcRun(src)
  ],
  ...sp(opts.before || 0, opts.after || 120)
});
const BULLETSRC = (text, src) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  children: [
    new TextRun({ text, font: "Arial", size: 22 }),
    srcRun(src)
  ],
  ...sp(0, 60)
});

// ── Data Table ──────────────────────────────────────────────
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

// ── Insight Box ─────────────────────────────────────────────
function insightBox(title, sections) {
  const rows = [];
  rows.push(new TableRow({
    children: [new TableCell({
      borders: borders(BLUE),
      shading: { fill: BLUE_L, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({
        children: [new TextRun({ text: title, font: "Arial", size: 24, bold: true, color: BLUE })],
        ...sp(0, 0)
      })]
    })]
  }));
  sections.forEach(({ label, items }) => {
    const children = [];
    children.push(new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: 22, bold: true, color: DARK })], ...sp(0, 60) }));
    items.forEach(({ text, src }) => {
      children.push(new Paragraph({
        children: [
          new TextRun({ text, font: "Arial", size: 22 }),
          ...(src ? [srcRun(src)] : [])
        ],
        ...sp(0, 60)
      }));
    });
    rows.push(new TableRow({
      children: [new TableCell({
        borders: borders(GRAY_B),
        shading: { fill: GRAY_L, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children
      })]
    }));
  });
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows });
}

// ── Service Table ───────────────────────────────────────────
function serviceTable(svc, items) {
  const leftW = 1600, rightW = 7426;
  const contentRows = items.map((item, i) => new TableRow({
    children: [
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
      new TableCell({
        borders: borders(),
        width: { size: rightW, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [
          new Paragraph({ children: [new TextRun({ text: `◆ ${item.title}`, font: "Arial", size: 22, bold: true, color: DARK })], ...sp(0, 60) }),
          new Paragraph({
            children: [
              new TextRun({ text: item.text, font: "Arial", size: 22 }),
              ...(item.src ? [srcRun(item.src)] : [])
            ],
            ...sp(0, 0)
          })
        ]
      })
    ]
  }));
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [leftW, rightW], rows: contentRows });
}

// ── Summary Table ──────────────────────────────────────────
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
//  報告內容
// ════════════════════════════════════════════════════════════

// ── 1.1 市場成長結構表格 ──
const mktTable = dataTable(
  ["指標", "2024 MAT", "2025 MAT", "YoY 成長率"],
  [
    ["整體銷售額", "—", "—", "+8%"],
    ["購買人數（不重複）", "—", "—", "+5%"],
    ["人均購買次數", "1.8 次", "1.8 次", "0%"],
    ["單次平均金額", "—", "NT$ 2,008", "+2%"],
  ],
  [2500, 2000, 2000, 2526]
);

// ── 1.2 通路表格 ──
const channelTable = dataTable(
  ["通路", "銷額佔比", "人流佔比", "銷額 YoY", "人流 YoY"],
  [
    ["好市多", "22%", "30%", "+12%", "+8%"],
    ["葡眾（直銷）", "11%", "1%", "-10%", "-6%"],
    ["momo", "8%", "7%", "+2%", "-7%"],
    ["蝦皮", "6%", "10%", "-3%", "0%"],
    ["營養師輕食", "5%", "1%", "-12%", "-21%"],
    ["大樹", "3%", "1%", "+10%", "+15%"],
    ["配方時代", "3%", "2%", "+226%", "+211%"],
    ["7-11", "2%", "5%", "+7%", "+6%"],
    ["康是美", "2%", "2%", "+4%", "+5%"],
  ],
  [1800, 1600, 1600, 1800, 2226]
);

// ── 1.3 品牌表格 ──
const brandTable = dataTable(
  ["品牌", "銷額佔比", "人流佔比", "銷額 YoY", "人流 YoY"],
  [
    ["威德", "16%", "24%", "-2%", "-2%"],
    ["葡眾", "12%", "2%", "-9%", "-2%"],
    ["營養師輕食", "6%", "2%", "-1%", "-2%"],
    ["善存", "4%", "6%", "+71%", "+38%"],
    ["BHK's", "3%", "5%", "0%", "+4%"],
    ["大正欣表飛鳴", "3%", "5%", "+10%", "+8%"],
    ["配方時代", "3%", "1%", "+214%", "+201%"],
    ["科克蘭", "2%", "2%", "+25%", "+10%"],
    ["葡萄王", "2%", "4%", "+18%", "+51%"],
  ],
  [1800, 1600, 1600, 1800, 2226]
);

// ── NES 表格 ──
const nesTable = dataTable(
  ["買者類型", "2024 人數", "2025 人數", "人流佔比", "人流 YoY"],
  [
    ["品類新進買者 (N)", "123,060", "117,987", "59%", "-4%"],
    ["品類既有買者 (E)", "65,771", "80,453", "41%", "+22%"],
    ["品類流失買者 (S)", "91,713", "108,378", "—", "+18%"],
  ],
  [2200, 1600, 1600, 1400, 2226]
);

// ── 威德 NES 表格 ──
const weiderNesTable = dataTable(
  ["買者類型", "人數佔比", "銷額佔比", "人數 YoY", "銷額 YoY"],
  [
    ["新客", "50%", "36%", "-16%", "-22%"],
    ["競品客", "9%", "7%", "-5%", "-16%"],
    ["既有買者", "41%", "57%", "+23%", "+19%"],
  ],
  [2000, 1600, 1600, 1800, 2026]
);

// ── 第二章：品牌專屬洞察 ──
const insights = [
  {
    title: '洞察 1 「單通路依賴症」—— 96% 銷額鎖死好市多，品牌無腹地可退',
    sections: [
      { label: "▶ 數據發現", items: [
        { text: "威德 2025 年 96% 銷額來自好市多，全家銷額 YoY 暴跌 -76%，其他綜合電商、藥妝幾乎無存在感。", src: "P4, P22, P33" },
        { text: "同期好市多內威德銷額市佔從 78% 下滑至 71%，善存銷額 YoY +107%，正快速擠壓威德份額。", src: "P4, P22" },
        { text: "市場其他品牌如配方時代（官網 +226%）、大正（康是美成長顯著）都在多通路佈局，威德幾乎是市場前十大品牌中通路最單一的。", src: "P21, P32" },
      ]},
      { label: "▶ 市場洞察", items: [
        { text: "威德的通路集中度已達危險水位。96% 銷額綁定單一通路意味著：好市多的任何政策變動（上架新競品、調整貨架位置、促銷檔期安排）都直接決定威德的生死。善存進入好市多一年就拿下 13% 滲透率，證明好市多的貨架空間並非威德獨享。", src: "P4, P55" },
        { text: "更關鍵的是，威德在好市多以外的通路正在「歸零」——全家 -76%、其他通路幾乎無量，品牌正在失去所有備用戰場。", src: "P22, P33 + 推論" },
      ]},
      { label: "▶ 突破方向", items: [
        { text: "威德需要建立「好市多以外的第二成長曲線」。綜合電商（momo/蝦皮）是益生菌市場成長中的通路，且益生菌買者在 momo 的併買行為以保健食品為主，消費者已有在電商購買保健品的習慣。建議優先佈局 momo 官方旗艦，以好市多同款產品的小包裝/試用組切入，降低電商首購門檻。", src: "P21, P24 + 推論" },
      ]},
    ]
  },
  {
    title: '洞察 2 「新客荒」—— 新客補充速度斷崖式下滑，成長引擎熄火',
    sections: [
      { label: "▶ 數據發現", items: [
        { text: "威德新客人數 YoY -16%，銷額 -22%，遠大於市場新客衰退幅度（-4%）。", src: "P38, P13" },
        { text: "好市多檔期首購新客佔比從 2024/05 的 44.7% 一路下滑至 2025/10 的 30.8%，實際首購人數也持續萎縮。", src: "P66" },
        { text: "新客人均次數 -3%、單次金額 -6%，不僅人數減少，新客的平均投入金額也在下滑。", src: "P38" },
        { text: "同期善存新客 +19%、配方時代新客 +197%、葡萄王新客 +55%，顯示市場並非整體缺新客，而是威德特別吸引不到。", src: "P39" },
      ]},
      { label: "▶ 市場洞察", items: [
        { text: "威德的新客危機不是市場環境問題，而是品牌自身的獲客能力問題。市場新客只衰退 -4%，威德卻衰退 -16%，代表威德正在額外流失原本可以獲取的新客。好市多檔期是威德唯一的大量獲客管道，但檔期吸引力正在遞減——首購佔比四個檔期連續下滑，檔期正在從「開源機器」變成「舊客補貨日」。", src: "P38, P66 + 推論" },
      ]},
      { label: "▶ 突破方向", items: [
        { text: "威德需要在好市多檔期之外建立持續性的新客獲取管道。考慮到益生菌市場 25-29 歲新客品類黏著度低（容易嘗試也容易離開），威德應鎖定 30-44 歲、已有益生菌購買習慣但尚未買過威德的「品類舊客 × 威德新客」族群，這群人轉換後的留存率遠高於品類新客。", src: "P15, P53 + 推論" },
      ]},
    ]
  },
  {
    title: '洞察 3 「善存侵蝕效應」—— 新競品開拓新客群，威德腹背受敵',
    sections: [
      { label: "▶ 數據發現", items: [
        { text: "善存 2024 年中進駐好市多，一年內銷額 YoY +107%，在好市多佔比從 6% 提升至 13%。", src: "P22, P54" },
        { text: "善存在好市多的新客中，86% 是好市多舊客但前期完全沒買益生菌的人。", src: "P56" },
        { text: "威德當期淨流出到善存 383 人為最大淨流出品牌，但流出至善存的 1,038 位買者中，53% 是前期才剛接觸威德的新客。", src: "P51, P52" },
        { text: "善存買者忠誠度僅 60%（4 成兼買競品），威德忠誠度維持 87.8%。", src: "P55" },
      ]},
      { label: "▶ 市場洞察", items: [
        { text: "善存的威脅不在於「搶走威德的鐵粉」，而在於「攔截威德的潛在新客」。善存 86% 新客是好市多舊客但之前不買益生菌的人——這群人正是威德最需要的增量來源。善存搶先一步把這些「潛在益生菌新客」轉換為善存用戶，等於在源頭掐住了威德的新客水源。", src: "P56, P52 + 推論" },
        { text: "另外，流出至善存的威德買者中 53% 是新客（前期才剛買威德），代表威德的新客留存機制有漏洞，新客在第一次購買後很容易被善存吸走。", src: "P52" },
      ]},
      { label: "▶ 突破方向", items: [
        { text: "威德不能只防守「現有買者不被搶走」（忠誠度 87.8% 其實很穩），更要搶先觸及「好市多裡還沒買過益生菌的人」。建議透過 invos 數據識別好市多會員中有保健食品購買習慣但尚未購買益生菌的族群，在善存之前建立品牌認知。同時強化新客首購後的留存機制，避免新客在第一次嘗試後就轉向善存。", src: "P55, P56, 策略推論" },
      ]},
    ]
  },
  {
    title: '洞察 4 「舊客撐盤假象」—— 舊客成長掩蓋結構性衰退訊號',
    sections: [
      { label: "▶ 數據發現", items: [
        { text: "威德既有買者人數 YoY +23%，銷額 +19%，是品牌唯一成長的客群，貢獻 57% 銷額。", src: "P38" },
        { text: "但舊客行為正在質變：人均次數 -6%（1.5 次）、重度買者人均次 -12%（3.9 次），購買頻率全面下滑。", src: "P4, P60" },
        { text: "威德中/重度買者人數增加，但輕度買者銷額 YoY -5%，正在流失。", src: "P60" },
        { text: "威德買者離開品類的人數是流失至競品的 5 倍，顯示更大的問題不是被搶走，而是消費者直接不買了。", src: "P49" },
      ]},
      { label: "▶ 市場洞察", items: [
        { text: "威德的「舊客成長」其實是倖存者偏差——留下來的人確實在買，但他們的購買頻率在下降，輕度用戶正在靜默流失。品牌銷額之所以只微跌 -2%，是因為舊客的「集中購買」（單次金 +2%）暫時撐住了場面。但一旦這批忠誠舊客的購買頻率繼續下滑，或開始嘗試善存等新選擇，品牌將沒有後備力量。", src: "P38, P60 + 推論" },
        { text: "更值得警惕的是：買者離開品類（直接不買益生菌）的比例遠高於流失至競品，意味著威德的產品可能沒有成功建立「持續補充」的消費習慣。", src: "P49 + 推論" },
      ]},
      { label: "▶ 突破方向", items: [
        { text: "威德需要從「一次性檔期大量購買」轉向「定期補充」的消費模式設計。現行模式下，消費者在 5 月/10 月檔期囤貨，中間半年幾乎不購買（非檔期月回購率僅 1-3%），這種消費節奏天然不利於建立品牌黏著度。建議建立補貨提醒機制，在消費者上次購買後 3-4 個月主動觸達，將「囤貨型購買」轉化為「週期性補充」。", src: "P67, P80, 策略推論" },
      ]},
    ]
  },
  {
    title: '洞察 5 「檔期天花板」—— 促銷已無法拉動增長，需突破結構瓶頸',
    sections: [
      { label: "▶ 數據發現", items: [
        { text: "2025/10 檔期放寬限購盒數（6→12 盒），但人均盒數 2.78 盒，反而低於 2025/05 的 2.90 盒。", src: "P64, P65" },
        { text: "購買行為兩極化：買 1-5 盒的比例從 82.4% 反彈至 87.5%，但少數重度買者購買 12 盒以上的比例從 1.1% 升至 2.6%。", src: "P65" },
        { text: "2025/10 檔期總買者數較 2025/05 下降 11.5%，銷售額下降 15.3%。", src: "P64" },
        { text: "2025/05 檔期買者在 2025/10 的新客回購率僅 18.9%（vs 歷史 23-25%），舊客回購率 42.6% 尚可但也略降。", src: "P81" },
      ]},
      { label: "▶ 市場洞察", items: [
        { text: "放寬限購盒數的策略假設是「消費者想買更多但被限制住了」，但數據證明這個假設是錯的——7 成消費者每檔期購買不超過 3 盒，這是消費需求的自然上限，不是限購造成的。放寬限購只讓少數重度買者多囤，對大多數消費者無感。", src: "P65 + 推論" },
        { text: "檔期銷額下滑的根本原因不是「買的量不夠」，而是「來的人不夠」。首購新客佔比從 44.7% 降至 30.8%，檔期正在從「開源機器」退化為「舊客補貨日」。", src: "P64, P66 + 推論" },
      ]},
      { label: "▶ 突破方向", items: [
        { text: "檔期策略需要從「拉深度」（讓人買更多盒）轉向「拉廣度」（讓更多新人來買）。建議將檔期資源的一部分轉移到「新客首購激勵」：例如首購專屬小包裝組合、好市多 APP 推播鎖定「有買保健食品但沒買過益生菌」的會員。同時可考慮在兩大檔期之間增設一個小型「新客體驗檔」，不依賴折扣而是以試用/體驗切入，擴大接觸面。", src: "策略推論" },
      ]},
    ]
  },
];

// ── 第三章：invos 三大服務建議 ──
const insightSvc = {
  name: "Insight",
  color: GREEN,
  tagline: "賦能智慧決策",
  items: [
    { title: "威德「新客健康指數」監測儀表板", text: "建立追蹤「新客獲取速度 vs 市場新客衰退速度 vs 競品（善存/配方時代）新客成長速度」的三維品牌健康指數模型。每季提供結構性預警報告，讓威德管理層能在銷售額數字之外，清楚看見品牌的新客獲取品質，提前制定防禦策略。", src: "【invos 服務】" },
    { title: "善存進入好市多後的競爭影響深度解碼", text: "持續追蹤善存在好市多的滲透率、新客來源結構（品類新客 vs 競品轉換）、忠誠度變化。每季比對威德與善存在好市多的買者重疊率、流動方向、購買頻次差異，精準評估善存對威德的實際侵蝕程度 vs 市場擴張貢獻。", src: "【invos 服務】" },
    { title: "好市多檔期效率追蹤報告", text: "每檔期結束後提供「首購新客數、新客回購率、舊客喚回率、人均盒數分佈」四維檔期效率報告，取代「看總銷額增減」的粗放評估方式。同步追蹤檔期間的非檔期月購買行為，識別「哪些消費者開始在非檔期也買」的高價值訊號。", src: "【invos 服務】" },
  ]
};

const mediaSvc = {
  name: "Media",
  color: TEAL,
  tagline: "高效鎖客擴散",
  items: [
    { title: "好市多「益生菌潛在新客」精準觸達計畫", text: "以 invos 發票數據識別好市多會員中「有購買保健食品但尚未買過益生菌」的族群（善存 86% 新客正是此類人），在好市多檔期前 2-4 週精準投放威德品牌教育廣告，搶先於善存建立「第一個益生菌品牌」的心智佔據。", src: "【invos 服務】" },
    { title: "電商通路「威德小包裝」種草計畫", text: "鎖定 momo/蝦皮上有益生菌瀏覽或購買紀錄但未買過威德的消費者，推廣威德電商專屬小包裝試用組（降低首購門檻），搭配好市多同款大包裝的換購優惠碼，建立「電商試用→好市多回購」的跨通路消費路徑。", src: "【策略方向】" },
    { title: "威德新客「首購後 90 天」留存廣告", text: "鎖定威德首購新客，在購買後 30/60/90 天分階段投放留存廣告：第 30 天推送使用提醒與效果科普，第 60 天推送回購優惠，第 90 天（接近下次檔期前）推送檔期預告。目標是將新客首購→回購轉換率從目前的 18.9% 提升至對標舊客水平。", src: "【invos 服務】" },
  ]
};

const apiSvc = {
  name: "API",
  color: PURPLE,
  tagline: "私域深度經營",
  items: [
    { title: "威德全通路買者私域識別池", text: "整合好市多、momo、蝦皮、全家等通路的威德購買記錄，建立可識別、可觸達的「威德買者私域名單」。目前威德缺乏好市多以外的消費者觸達能力，私域名單是所有後續「非檔期觸達」行動的數據基礎。", src: "【invos 服務】" },
    { title: "「檔期間補貨提醒」自動化觸達", text: "以消費者上次購買時間為基準，在購買後 3-4 個月（預估食用完畢時間）自動推送補貨提醒。目標是將目前「5 月囤 → 10 月補」的半年購買週期，縮短為每 3-4 個月一次的穩定週期，提升人均年購次數（目前僅 1.5 次）。", src: "【invos 服務】" },
    { title: "威德流失買者「品類離開預警」系統", text: "威德買者離開品類的人數是流失至競品的 5 倍（P49），代表最大威脅不是善存搶客，而是消費者直接不買益生菌了。透過 invos API 識別「已 6 個月未購買且無競品購買紀錄」的威德流失買者，在完全離開品類前精準推送召回方案。", src: "P49, 【invos 服務】" },
  ]
};

// ── 第四章：策略重點一覽 ──
const summaryData = [
  { name: "Insight", color: GREEN, items: [
    "新客健康指數監測儀表板 — 追蹤新客獲取 vs 競品搶客三維指標（洞察#2）",
    "善存競爭影響深度解碼 — 季度追蹤善存對威德的實際侵蝕程度（洞察#3）",
    "好市多檔期效率追蹤報告 — 四維評估取代看銷額（洞察#5）",
  ]},
  { name: "Media", color: TEAL, items: [
    "好市多益生菌潛在新客觸達 — 搶先善存建立心智佔據（洞察#3）",
    "電商小包裝種草計畫 — 建立好市多以外的第二成長通路（洞察#1）",
    "新客首購後 90 天留存廣告 — 提升 18.9% 回購率至舊客水平（洞察#2）",
  ]},
  { name: "API", color: PURPLE, items: [
    "全通路買者私域識別池 — 非檔期觸達的數據基礎（洞察#1）",
    "檔期間補貨提醒自動化 — 縮短半年購買週期至 3-4 個月（洞察#4）",
    "品類離開預警系統 — 攔截「直接不買」的最大流失源（洞察#4）",
  ]},
];

// ════════════════════════════════════════════════════════════
//  Document Assembly
// ════════════════════════════════════════════════════════════
async function main() {
  const children = [
    // ── 封面 ──
    BLANK(),
    new Paragraph({ children: [new TextRun({ text: "【引客數據 invos】", font: "Arial", size: 28, bold: true, color: BLUE })], alignment: AlignmentType.CENTER, ...sp(0, 80) }),
    new Paragraph({ children: [new TextRun({ text: `${PERIOD} ${CATEGORY}市場洞察分析報告`, font: "Arial", size: 36, bold: true })], alignment: AlignmentType.CENTER, ...sp(80, 80) }),
    new Paragraph({ children: [new TextRun({ text: `${BRAND}品牌專屬市場洞察與 invos 服務策略建議`, font: "Arial", size: 24, italics: true, color: "555555" })], alignment: AlignmentType.CENTER, ...sp(0, 80) }),
    new Paragraph({ children: [new TextRun({ text: `數據觀測期：${DATE_RANGE} vs ${DATE_COMP}　｜　樣本：約 ${SAMPLE_N} 人`, font: "Arial", size: 20, color: "888888" })], alignment: AlignmentType.CENTER, ...sp(0, 480) }),

    // ══════════════════════════════════════════
    // 第一章：益生菌市場整體現況
    // ══════════════════════════════════════════
    H1("一、益生菌市場整體現況"),

    // 1.1
    H2("1.1 市場成長結構"),
    PSRC("益生菌市場 2025 年銷售額 YoY +8%，主要受購買人數成長 +5% 所驅動，單次金額微幅成長 +2%（NT$ 2,008），人均次數則維持約 1.8 次不變。市場成長結構以「更多人來買」為主要動能，而非既有買者「買更多次」。", "P3, P10"),
    BLANK(),
    mktTable,
    BLANK(),

    // 指標市場意義解讀
    P("▎指標市場意義解讀", { size: 22, bold: true, before: 120 }),
    PSRC("1. 銷售額（Sales Amount）— 市場整體的健康指標：+8% 的成長率顯示益生菌市場仍處於穩健成長期，消費需求持續擴張，尚未進入飽和。", "P10"),
    PSRC("2. 購買人數（Number of Buyers）— 市場滲透率與廣度：+5% 的人數成長是市場擴張的核心訊號，代表益生菌品類仍在持續滲透新消費者。但成長率低於銷售額（+8%），意味著部分成長也來自既有消費者的貢獻提升。", "P10"),
    PSRC("3. 人均次數（Purchase Frequency）— 品牌黏著度與習慣：維持 1.8 次不變，代表益生菌仍屬於「低頻補充型」品類，多數消費者一年購買不到 2 次。這與好市多檔期驅動的囤貨型消費模式高度相關。", "P10"),
    PSRC("4. 單次金額（Transaction Value）— 消費升級與單價：+2% 的微幅上升可能反映 (1) 高單價品牌（善存 NT$1,830、配方時代等）進入市場 (2) 消費者傾向購買較多數量 (3) 好市多量販包裝單次購買金額天然較高。", "P10, P11"),
    new Paragraph({
      children: [new TextRun({ text: "綜合分析：益生菌市場正處於「量增價穩」的健康成長期。成長主要靠新進消費者帶動（人數 +5%），而非既有消費者買更多次。但市場高度依賴好市多通路（銷額佔比 22%），且購買頻率偏低（1.8 次/年），意味著品牌間的競爭本質是「搶人」而非「搶頻次」——誰能先觸及並留住新進消費者，誰就掌握增量。", font: "Arial", size: 22, bold: true }),
      srcRun("P3, P10, P21 + 推論")],
      ...sp(60, 120)
    }),
    BLANK(),

    // 1.2 通路
    H2("1.2 通路分佈與競爭格局"),
    PSRC("好市多為益生菌市場 Top 1 通路，銷額佔比 22%，人流持續成長 +8%。綜合電商（momo + 蝦皮合計約 14%）銷額微幅成長，但蝦皮人流下滑。直銷通路中葡眾位居市場第二（11%），但整體表現下滑。配方時代在品牌官網經營表現突出（+226%），是當期黑馬。", "P21"),
    BLANK(),
    channelTable,
    BLANK(),

    // 通路 × 品牌競爭定位
    P("▎通路 × 品牌競爭定位", { size: 22, bold: true, before: 120 }),
    PSRC("威德主要銷售管道為好市多（96% 銷額），在全家也有上架但銷額大幅衰退 -76%。好市多當期銷額成長 +12%，威德在其中成長 +2%，但善存成長 +147% 大幅擠壓威德份額（78%→71%）。", "P4, P22, P74"),
    PSRC("好市多內主要競品：善存（佔比 6%→13%）為威德最大威脅，其次為樂益活、大正欣表飛鳴。", "P22"),
    PSRC("momo 通路以營養師輕食、BHK's、大研生醫為主要品牌，威德在 momo 存在感極低。", "P24"),
    PSRC("康是美以 BHK's、倍適為主，蝦皮以 MIHONG 等平價品牌為主，威德在藥妝/電商通路幾乎無佈局。", "P11, P22"),
    BLANK(),

    // 1.3 品牌競爭
    H2("1.3 品牌競爭態勢"),
    PSRC("市場品牌多元，前三大品牌（威德 16%、葡眾 12%、營養師輕食 6%）成長率微幅下滑或持平，而市佔 2-5% 的中型品牌多數成長率超過 10%，正在快速搶佔市場增量。", "P32"),
    BLANK(),
    brandTable,
    BLANK(),
    BULLETSRC("快速成長品牌：善存（+71%，好市多新進品牌）、配方時代（+214%，官網直營爆發）、葡萄王（+18%，7-11 通路佈局成效）", "P32"),
    BULLETSRC("穩定品牌：BHK's（0%）、大正（+10%），在藥妝/電商通路穩定經營", "P32"),
    BULLETSRC("衰退品牌：葡眾（-9%，直銷通路下滑）、營養師輕食（-1%）", "P32"),
    BLANK(),

    // 1.4 NES
    H2("1.4 新舊客買者狀態（NES 分析）"),
    P("益生菌市場已進入「存量競爭」階段，經營重點正從「積極獲新」轉向「精準留存」。", { size: 22, bold: true }),
    BLANK(),
    nesTable,
    BLANK(),
    PSRC("New（品類新進買者）：123,060→117,987 人，YoY -4%。雖然新客仍佔多數（59%），但人數已轉為負成長，代表市場吸引新消費者的整體力道正在減弱，新客獲取成本可能正在墊高。", "P13"),
    PSRC("Existing（品類既有買者）：65,771→80,453 人，YoY +22%。核心消費者群體持續擴大，是市場的「壓艙石」。既有買者佔比從 35% 提升至 41%，貢獻 61% 銷額，人均貢獻度遠高於新客。", "P13, P14"),
    PSRC("Sleeping（品類流失買者）：91,713→108,378 人，YoY +18%。流失人數增幅高，且絕對數量已超過新客（108,378 vs 117,987），市場正逼近「淨流出」臨界點，需密切留意。", "P13"),
    new Paragraph({
      children: [new TextRun({ text: "市場意義總結：新客微降、舊客成長、流失增加 → 典型的存量競爭格局。品牌間的競爭本質已從「誰吸引更多新人」轉為「誰能讓人持續回購」。", font: "Arial", size: 22, bold: true }),
      srcRun("P13, P14 + 推論")],
      ...sp(60, 120)
    }),
    BULLETSRC("新/流失買者中 25-29 歲佔比高於既有客群，顯示年輕族群品類黏著度較低，購買行為較易轉變", "P15"),
    BULLETSRC("女性佔整體市場約 69%，重度買者中女性 35-44 歲（42%）為最高貢獻族群", "P3, P19"),
    BLANK(),

    // 1.5 HML
    H2("1.5 輕中重度買者狀態（HML 分析）"),
    PSRC("重度買者（前 10%）貢獻整體 50% 營收，人流成長 +13%，是市場銷額成長的核心引擎。中度買者人數 +11%，帶動銷額 +11%。輕度買者成長有限（+2%），主因為新客數減少。", "P17, P18"),
    BULLETSRC("重度買者：人均年購 4.5 次，年消費 > NT$7,504，人流 +13% 但人均次 -4%、單次金 -5%，顯示購買行為趨向「集中少次」", "P17, P18"),
    BULLETSRC("中度買者：年消費 NT$3,036-7,504，人數 +11% 是市場擴張的中堅力量", "P17, P18"),
    BULLETSRC("輕度買者：佔 70% 人數但僅貢獻 23% 銷額，成長有限，多為一次性購買的新客", "P17, P18"),
    BULLETSRC("通路差異：直銷/品牌通路以中重度買者為主，好市多/電商/藥妝面向輕中度買者", "P30"),
    BLANK(),

    // ══════════════════════════════════════════
    // 第二章：威德品牌專屬市場洞察
    // ══════════════════════════════════════════
    H1("二、威德品牌專屬市場洞察"),
    P("以下五大洞察均基於威德在數據中的獨特表現，具高度針對性。每則洞察若「換成其他益生菌品牌」則不適用。", { size: 22 }),
    BLANK(),

    // 威德品牌總覽
    P("▎威德品牌總覽", { size: 22, bold: true }),
    PSRC("威德整體銷額 YoY -2%，購買人數 -2%，人均次數 1.5 次（-6%），單次金額 NT$1,634（+2%）。銷額微幅衰退，受買者數及人均次減少影響。", "P4"),
    BLANK(),
    weiderNesTable,
    BLANK(),

    ...insights.flatMap(ins => [insightBox(ins.title, ins.sections), BLANK()]),

    // ══════════════════════════════════════════
    // 第三章：invos 三大服務策略建議
    // ══════════════════════════════════════════
    H1("三、invos 三大服務策略建議"),
    P("以下策略建議基於上述五大威德專屬洞察，精準扣合引客數據三大服務功能，提供可落地執行的行動方向。", { size: 22 }),
    BLANK(),

    H2("3.1 invos Insight｜市場機會洞察 × 賦能威德智慧決策"),
    new Paragraph({ children: [new TextRun({ text: "對應洞察：#2 新客荒 ／ #3 善存侵蝕效應 ／ #5 檔期天花板", font: "Arial", size: 20, italics: true, color: "666666" })], ...sp(0, 120) }),
    serviceTable(insightSvc, insightSvc.items),
    BLANK(),

    H2("3.2 invos Media｜精準獲客轉換 × 高效拓展威德新客"),
    new Paragraph({ children: [new TextRun({ text: "對應洞察：#1 單通路依賴症 ／ #2 新客荒 ／ #3 善存侵蝕效應", font: "Arial", size: 20, italics: true, color: "666666" })], ...sp(0, 120) }),
    serviceTable(mediaSvc, mediaSvc.items),
    BLANK(),

    H2("3.3 invos API｜全通路會員整合 × 建立威德私域護城河"),
    new Paragraph({ children: [new TextRun({ text: "對應洞察：#1 單通路依賴症 ／ #4 舊客撐盤假象 ／ #5 檔期天花板", font: "Arial", size: 20, italics: true, color: "666666" })], ...sp(0, 120) }),
    serviceTable(apiSvc, apiSvc.items),
    BLANK(),

    // ══════════════════════════════════════════
    // 第四章：策略重點一覽
    // ══════════════════════════════════════════
    H1("四、策略重點一覽"),
    summaryTable(summaryData),
    BLANK(),

    // ── 免責說明 ──
    new Paragraph({
      children: [new TextRun({ text: `本報告基於引客數據 invos 電子發票消費追蹤系統，觀測期 ${DATE_RANGE}（${PERIOD}） vs ${DATE_COMP}，樣本約 ${SAMPLE_N} 人。所有洞察與策略建議均根據真實消費行為數據推導，具${BRAND}品牌專屬針對性，僅供品牌內部策略參考使用。`, font: "Arial", size: 18, italics: true, color: "888888" })],
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
