/**
 * InvosData 閱讀邏輯文件生成腳本
 *
 * 用法：
 *   1. 修改下方「填入新簡報資料」區段（BRAND、PRODUCT、COLS 等）
 *   2. node skills/reading-guide/references/generate_guide.js
 *
 * 依賴：npm install（在 repo 根目錄）
 * 輸出：output/reading-guide/<BRAND>_<PRODUCT_SHORT>_閱讀邏輯.docx
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, Header, Footer, PageNumber,
} = docxLib;
const fs = require('fs');
const path = require('path');

// ════════════════════════════════════════════════════════════════
//  ★  填入新簡報資料（每次更換簡報只需修改這個區塊）
// ════════════════════════════════════════════════════════════════

const BRAND        = '威德';             // 品牌名稱
const PRODUCT      = '睡眠益生菌';       // 產品名稱
const PRODUCT_SHORT = '睡眠益生菌';      // 用於檔名（不含空格）
const CHANNELS     = '好市多 × momo';    // 通路範圍
const DATE         = '2026.03';          // 簡報日期
const CLIENT_TYPE  = '品牌業主';         // 客戶類型描述

// Dashboard 五欄資料
// 每欄包含：title、bcg（'STAR' 或 'QMARK'）、metric（關鍵指標）
// 以及四個內容陣列：bn（瓶頸）、op（機會）、ac（行動）、iv（InvosData方案）
const COLS = [
  {
    title: '好市多通路整體',
    bcg: 'STAR',
    metric: '95% 營收集中  ／  份額54.7%  ／  銷售+63%、人流+50%',
    bn: [
      '新普利2025年進場，通路集中風險正在升高',
      '人均次數成長0%，業績完全靠新人流而非回購，結構脆弱',
      '無機制取得好市多客戶聯繫方式，每賣一盒等於流失一個客戶',
    ],
    op: [
      '份額54.7% 主導好市多，遠超老協珍28%，品牌地位穩固',
      '溢價24% 仍吸引人流持續流入（+50%），品質感知超出定價預期',
    ],
    ac: [
      '外盒加LINE QR Code，「28天睡眠挑戰」建立好市多購買→品牌會員路徑',
      '強化momo備援，降低95%集中風險，提前布局應對新普利競爭',
    ],
    iv: [
      'Insight：月追蹤威德/老協珍/新普利好市多份額與人流',
      'API：好市多發票匿名交易轉為品牌可追蹤CRM名單',
    ],
    key_argument:  '份額第一，但95%集中在單一通路——新競品進場就是結構性風險',
    key_number:    '每賣出一盒，威德就流失一個客戶（無任何後續接觸機制）',
    invos_angle:   'invos API：好市多匿名發票 → 品牌 CRM 名單，從「流失」變「可追蹤」',
    qa: {
      q: '95%集中好市多不是很正常嗎？',
      a: '正常，但脆弱。新普利2025年已進場——如果好市多給他們好位置，份額就會被吃。建議現在用momo做備援，不是因為好市多不好，是因為95%在任何一個通路都是高風險。',
    },
  },
  {
    title: '品類新客',
    bcg: 'STAR',
    metric: '92% 為品類新客  ／  新客銷售+37%、人流+35%',
    bn: [
      '92% 新客首購後無接觸機制（無LINE、無教育內容）',
      '35–39歲主力客群需求語境尚未驗證，廣告溝通切角未對應',
    ],
    op: [
      '好市多為主要品類新客入口（95%營收×92%新客），先發優勢顯著',
      '威德新客客單NT$2,207 > 老協珍NT$1,803，對品質期望更高',
    ],
    ac: [
      '外盒強化「益生菌×睡眠」科學差異化說明，讓NT$2,360溢價有明確依據',
      '問卷圈選25–39歲受眾驗證睡眠需求語境，再設計對應廣告素材投放',
    ],
    iv: [
      'Media：圈選「好市多保健品買者但未買睡眠益生菌」精準受眾投廣',
      'API + Meta：購買後自動回傳發票訊號，量化廣告帶動好市多實際新購ROAS',
    ],
    key_argument:  '92%新客是機會，但首購後沒有任何機制接觸他們——Costco客戶而非品牌客戶',
    key_number:    '威德新客客單NT$2,207 > 老協珍NT$1,803，選威德的人對品質期望更高',
    invos_angle:   'invos Media：圈選「買過保健品但從未買睡眠益生菌」的好市多買者，精準獲取高意圖新客',
    qa: {
      q: '92% 新客代表品牌越來越大，不好嗎？',
      a: '新客多是好事，但重度買者只有6%（益喜氏21%）。如果新客買了一次、兩次就不見了，成長只是短暫的。真正健康的品牌，新客和回頭客要同步成長。',
    },
  },
  {
    title: '品牌舊客 × 消費升級',
    bcg: 'QMARK',
    metric: '重度買者僅6%（vs. 益喜氏21%）  ／  留客機制缺失',
    bn: [
      '無任何回購提醒機制，完全靠消費者自發記憶，天然流失率高',
      '重度買者6%正在衰退（vs. 益喜氏21%），重度銷額貢獻僅19%',
      '輕度72%無升級路徑設計，中度22%升級動力不足',
    ],
    op: [
      '既有客頻率1.47次 vs 新客1.11次，加觸發機制可顯著提升回購',
      '若重度比例6%→12%，總銷額可新增15–20%（益喜氏模型驗證）',
    ],
    ac: [
      'LINE第21天推補貨提醒；檔期前1週推好市多優惠；空窗期推momo/官網',
      '推出雙盒組合包獨立規格，直接觸發重度消費行為；附睡眠評測問卷推動升級',
    ],
    iv: [
      'API：上傳發票解鎖回饋，自動21天推播回購，匿名交易轉為品牌CRM名單',
      'Media：圈選上一檔期好市多買者，在下一檔期前1週精準投廣',
    ],
    key_argument:  '重度買者6% vs. 益喜氏21%——威德同時缺重度買者的「人數」和「消費深度誘因」',
    key_number:    '若重度6%→12%，總銷額可新增15–20%（依益喜氏現有結構反推）',
    invos_angle:   'invos API：自動21天回購推播，把「靠記憶回購」變成「機制驅動的回購」',
    qa: {
      q: '我們已經有老客了，為什麼還需要做留客機制？',
      a: '有舊客不等於有留客機制——現在的舊客完全靠自發記憶回來買，這是最脆弱的留存方式。invos API的21天推播可以把自然流失率降低，同時圈出有回購潛力的買者讓廣告更精準。',
    },
  },
  {
    title: '競品轉換客',
    bcg: 'STAR',
    metric: '老協珍→威德淨流入+111人  ／  流出僅14人',
    bn: [
      '新普利2025年進場，三品牌同台競爭，品類新客選擇增加',
      '威德與老協珍均以「保健品」定位呈現，成份差異不夠清晰，買者容易混淆',
    ],
    op: [
      '好市多競品轉換勝率高（淨流入+111人，流出14人），現有策略具吸引力',
      '老協珍→威德轉換後消費金額大幅升級，轉換客願意為益生菌配方支付更高費用',
    ],
    ac: [
      '強化好市多貨架說明牌（shelf talker），建立「益生菌 vs 芝麻素/GABA」心智區隔',
      '老協珍促銷週期結束後優先進場試吃/堆頭，搶在舊客未續購前提供體驗',
    ],
    iv: [
      'Media：圈選「購買老協珍/新普利但從未購買威德」競品客群，在採購週期前精準投廣',
      'API：離線轉換追蹤，驗證競品截流廣告對好市多實際轉換的貢獻度',
    ],
    key_argument:  '競品轉換勝率高，但新普利進場讓三方競爭格局改變——優勢要主動鞏固',
    key_number:    '老協珍→威德淨流入+111人，流出僅14人——說明益生菌配方的吸引力已超過芝麻素',
    invos_angle:   'invos Media：圈選「買過老協珍但從未買威德」的買者，在下次採購週期前精準截流',
    qa: {
      q: '新普利才剛進場，威脅還沒出現，現在需要擔心嗎？',
      a: '好市多的貨架位置是零和的——新普利拿到好位置的那天，才是影響開始發酵的時候。我們建議現在就用 invos Insight 追蹤，把反應時間從「察覺後才行動」縮短到「數字出現就部署」。',
    },
  },
  {
    title: 'momo 第二通路',
    bcg: 'QMARK',
    metric: '市場份額22%（好市多8%的2.75倍）  ／  客單NT$2,650  ／  WTP NT$3,027',
    bn: [
      '95%集中好市多，momo未系統布局，品牌旗艦頁薄弱，電商搜尋能見度低',
      '好市多空窗期（約8–9個月）舊客補貨無通路承接，缺差異化規格',
    ],
    op: [
      'momo份額22%是好市多8%的2.75倍，屬潛在增量市場',
      '消費者WTP NT$3,027 > 客單NT$2,650，仍有溢價空間可挖',
    ],
    ac: [
      '建立momo品牌旗艦店，完善商品頁與「益生菌×睡眠」關鍵字佈局',
      '推出momo獨家90包加量規格，定價NT$2,899，對應WTP NT$3,027，與好市多差異化',
      '好市多空窗期以invos Media圈選上一檔期買者推播momo補貨廣告',
    ],
    iv: [
      'Insight：追蹤威德/益喜氏/新普利在momo份額與客單，動態調整定價策略',
      'API：串接好市多×momo發票，計算真實跨渠道回購率作為資源投入核心KPI',
    ],
    key_argument:  'momo份額22%是好市多8%的2.75倍——好市多空窗期的8–9個月，舊客補貨去哪裡？',
    key_number:    '消費者WTP NT$3,027 > momo客單NT$2,650——代表有定價空間尚未被挖掘',
    invos_angle:   'invos Media：圈選好市多上一檔期買者，在空窗期精準推播momo補貨廣告（轉換意願最高的受眾）',
    qa: {
      q: 'momo的22%份額是整體市場，不是我們品類，參考價值有限吧？',
      a: '你說得對，22%是睡眠保健品在momo的整體份額——是趨勢方向，不是威德自己的數字。威德在momo的現況份額可以另外出細分分析。但趨勢方向很明確：消費者去momo買睡眠保健品的比例遠高於好市多，這個空間是真實的。',
    },
  },
];

// 附錄數據（按「指標 → 數值」格式）
const APPENDIX_DATA = [
  ['威德品類新客人數佔比', '92%（市場77%、益喜氏76%）'],
  ['威德品類新客銷額貢獻', '87%（市場60%、益喜氏68%）'],
  ['威德重度買者人數佔比', '約6%（市場9%、益喜氏21%）'],
  ['威德重度買者銷額貢獻', '19%（市場41%、益喜氏50%）'],
  ['威德輕度買者人數佔比', '72%（市場75%、益喜氏39%）'],
  ['好市多營收佔比', '95%'],
  ['好市多通路份額', '54.7%（老協珍28%）'],
  ['好市多銷售/人流成長', '+63%、人流+50%'],
  ['老協珍→威德淨轉換', '+111人（流出僅14人）'],
  ['momo 市場份額', '22%（好市多8%的2.75倍）'],
  ['momo 客單 / WTP', 'NT$2,650；消費者WTP NT$3,027'],
  ['睡眠益生菌成份市場成長', '+6%（鎂+58%；芝麻素-14%、色胺酸-39%）'],
];

// 開場話術與結尾話術
const OPENING_SCRIPT = '「這份報告的三頁，第一頁是今天的主題，第二頁是我們用來看問題的視角，第三頁是結論——所有你需要知道的事情都在第三頁。我們今天花最多時間在第三頁。」';
const CLOSING_SCRIPT = `「我們今天看到的五個維度，每一個都指向同一個核心問題：${BRAND}在${CHANNELS.split('×')[0].trim()}賣得很好，但沒有把這些買家變成可以繼續對話的品牌資產。最小的第一步，是讓每一盒都有辦法認識買家——這就是 invos API 在做的事。我們可以先從這裡談起。」`;

// ════════════════════════════════════════════════════════════════
//  以下為文件生成程式碼，一般不需要修改
// ════════════════════════════════════════════════════════════════

const FONT       = 'Microsoft JhengHei';
const CONTENT_W  = 9360;

const border  = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };

// Helper: heading
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: '1E3A5F' })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: '1E293B' })],
  });
}
function h3(text, color = '334155') {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color })],
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 40, after: 80, line: 320 },
    children: [new TextRun({
      text, font: FONT,
      size: opts.size || 22,
      color: opts.color || '374151',
      bold: opts.bold || false,
      italics: opts.italic || false,
    })],
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 40, after: 60, line: 300 },
    children: [new TextRun({ text, font: FONT, size: 21, color: '374151' })],
  });
}
function tip(text) {
  return new Paragraph({
    spacing: { before: 40, after: 80, line: 300 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: '\uD83D\uDCA1 ', font: FONT, size: 21 }),
      new TextRun({ text, font: FONT, size: 21, color: '374151', italics: true }),
    ],
  });
}
function divider() {
  return new Paragraph({
    spacing: { before: 180, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: 'E2E8F0', space: 1 } },
    children: [],
  });
}
function spacer(before = 120) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [] });
}

function callout(label, text, bg = 'EFF6FF', accent = '1E3A5F') {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 4, color: accent },
        bottom: { style: BorderStyle.NONE },
        left:   { style: BorderStyle.THICK, size: 16, color: accent },
        right:  { style: BorderStyle.NONE },
      },
      shading: { fill: bg, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 200, right: 120 },
      width: { size: CONTENT_W, type: WidthType.DXA },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: label, font: FONT, size: 20, bold: true, color: accent })],
        }),
        new Paragraph({
          spacing: { before: 0, after: 0, line: 300 },
          children: [new TextRun({ text, font: FONT, size: 20, color: '374151' })],
        }),
      ],
    })]})],
  });
}

function twoCol(rows, colW = [2160, 7200]) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colW,
    rows: rows.map(([left, right, isHeader]) =>
      new TableRow({ children: [
        new TableCell({
          borders, width: { size: colW[0], type: WidthType.DXA },
          shading: { fill: isHeader ? '1E293B' : 'F1F5F9', type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 120 },
          verticalAlign: VerticalAlign.TOP,
          children: [new Paragraph({
            spacing: { before: 0, after: 0, line: 280 },
            children: [new TextRun({ text: left, font: FONT, size: 20, bold: true, color: isHeader ? 'F1F5F9' : '1E3A5F' })],
          })],
        }),
        new TableCell({
          borders, width: { size: colW[1], type: WidthType.DXA },
          shading: { fill: isHeader ? '1E293B' : 'FFFFFF', type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 120 },
          verticalAlign: VerticalAlign.TOP,
          children: [new Paragraph({
            spacing: { before: 0, after: 0, line: 280 },
            children: [new TextRun({ text: right, font: FONT, size: 20, color: isHeader ? 'F1F5F9' : '374151', bold: isHeader })],
          })],
        }),
      ]})
    ),
  });
}

// ── Build document ────────────────────────────────────────────────────────

const children = [];

// Cover
children.push(spacer(0));
children.push(new Paragraph({
  spacing: { before: 0, after: 80 },
  children: [new TextRun({ text: `${BRAND} × InvosData`, font: FONT, size: 22, color: '64748B' })],
}));
children.push(new Paragraph({
  spacing: { before: 0, after: 200 },
  children: [new TextRun({ text: '簡報閱讀邏輯', font: FONT, size: 44, bold: true, color: '1E3A5F' })],
}));
children.push(new Paragraph({
  spacing: { before: 0, after: 60 },
  children: [new TextRun({ text: `${CHANNELS} 行銷落地策略  ／  報告者參考手冊`, font: FONT, size: 24, color: '475569' })],
}));
children.push(new Paragraph({
  spacing: { before: 0, after: 480 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E3A5F', space: 1 } },
  children: [new TextRun({ text: `${DATE}  ／  InvosData`, font: FONT, size: 20, color: '94A3B8' })],
}));

// 一、文件用途
children.push(h1('一、文件用途'));
children.push(body(`本文件供 InvosData 業務人員或簡報人在向${BRAND}客戶進行提案前閱讀，說明整份三頁簡報的敘事邏輯、各頁核心訊息、銜接話術，以及常見的客戶問題與回應建議。`));
children.push(spacer(80));
children.push(callout('簡報架構一覽', 'P.1 封面  →  P.2 策略框架說明（BCG矩陣）  →  P.3 五大客群行銷 Dashboard', 'EFF6FF', '1E3A5F'));
children.push(spacer(160));

// 二、整體敘事架構
children.push(h1('二、整體敘事架構'));
children.push(body('整份簡報遵循「診斷 → 框架 → 行動」三段式結構，引導客戶從資料洞察出發，接受 InvosData 解決方案的必要性。'));
children.push(spacer(120));
children.push(twoCol([
  ['頁碼', '核心任務', true],
  ['P.1  封面',      `建立信任感與專業定位——讓客戶在開口之前就感受到「這是為${BRAND}量身打造的分析」`],
  ['P.2  BCG框架',   '建立共同語言——讓客戶理解 STAR / Question Mark 的判斷邏輯，避免後續逐欄說明時反覆解釋術語'],
  ['P.3  Dashboard', '交付核心價值——用一頁總攬五大客群的問題、機會與建議行動，並自然帶出 InvosData 各產品的使用場景'],
], [1440, 7920]));
children.push(spacer(180));
children.push(body('三個頁面之間的銜接邏輯如下：', { bold: true }));
children.push(spacer(60));
children.push(bullet('封面奠定「數據驅動、針對性」的基調，不是通用型簡報'));
children.push(bullet('BCG說明頁讓客戶「先懂框架再看數字」，降低 Dashboard 頁的認知負擔'));
children.push(bullet('Dashboard 頁用視覺密度傳達「我們幫你想清楚了」的訊息，而不是把分析報告丟回給客戶'));
children.push(spacer(80));
children.push(tip('建議先確認客戶對 BCG 矩陣的熟悉程度。若客戶有品牌管理背景可快速帶過 P.2，若為中小品牌業主則 P.2 需花較多時間說明。'));
children.push(divider());

// 三、逐頁閱讀邏輯
children.push(h1('三、逐頁閱讀邏輯'));

// P.1
children.push(h2('P.1  封面'));
children.push(h3('頁面目標', '1E3A5F'));
children.push(body(`快速建立「這份分析是為${BRAND}量身訂製」的信任感，同時讓客戶知道資料來源是可信的。`));
children.push(spacer(100));
children.push(h3('說話重點', '1E3A5F'));
children.push(bullet(`主動點出標題中的「${CHANNELS}」——這是${BRAND}最在乎的通路，點出來就能讓對方瞬間感受到針對性`));
children.push(bullet('提及「發票實購消費者分析」時，強調這是真實購買行為數據，而非問卷或市調，藉此與一般顧問報告做區隔'));
children.push(bullet(`可加一句破冰語：「我們這次想用數字告訴你，${BRAND}現在最應該把資源放在哪裡」`));
children.push(spacer(100));
children.push(callout('建議開場話術（參考）', OPENING_SCRIPT, 'F0FFF4', '16A34A'));
children.push(spacer(160));

// P.2
children.push(h2('P.2  策略定位說明（BCG矩陣框架）'));
children.push(h3('頁面目標', '1E3A5F'));
children.push(body('讓客戶在看 Dashboard 之前，理解每個客群被標記為 STAR 或 Question Mark 的判斷邏輯，以及這代表什麼樣的資源優先順序。'));
children.push(spacer(100));
children.push(h3('兩個標籤的解釋重點', '1E3A5F'));
children.push(spacer(80));
children.push(twoCol([
  ['標籤', '解釋重點', true],
  ['★ STAR（明星）',       '「表現已強，但不是可以放著不管的意思——優勢要主動鞏固，否則競品跟進後會失去。」'],
  ['?  Question Mark\n（待解鎖）', '「機會存在，但目前尚未有效掌握——不是不重要，而是現在的做法還沒有把這塊做起來。」'],
], [2160, 7200]));
children.push(spacer(120));
children.push(h3('常見卡點', '1E3A5F'));
children.push(bullet('客戶可能對「STAR 還需要投資」感到困惑——可類比足球：攻勢好的球隊也必須守住陣地'));
children.push(bullet('客戶可能問「為什麼 momo 是 Question Mark」——要解釋整體市場 momo 份額高達 22%，機會存在，只是目前還沒布局'));
children.push(spacer(80));
children.push(tip('這頁通常不超過 3 分鐘。目的是建立語言，不是深入討論 BCG 理論。'));
children.push(divider());

// P.3
children.push(h2('P.3  行銷策略 Dashboard（核心頁）'));
children.push(h3('頁面目標', '1E3A5F'));
children.push(body('用一頁總攬五大客群的「當前瓶頸 → 核心機會 → 行動建議 → InvosData 方案」，讓客戶感受到「問題被精準辨識、解法已經準備好了」。'));
children.push(spacer(100));
children.push(h3('閱讀順序建議', '1E3A5F'));
children.push(body('Dashboard 資訊密度高，不建議逐欄朗讀。推薦以下兩種方式：'));
children.push(spacer(80));
children.push(callout(
  '方式 A：以「問題導向」引導（建議首次提案使用）',
  '先問客戶：「你現在最擔心哪一塊？」——依客戶回答跳到對應欄位，把 Dashboard 變成對話工具而非單向播報。',
  'FEF9EC', 'D97706'
));
children.push(spacer(120));
children.push(callout(
  '方式 B：以「輕重緩急」帶出順序（建議複訪或深談時使用）',
  '由左至右代表優先順序。前三欄是近期重點，momo 屬於中長期佈局。',
  'EFF6FF', '1E3A5F'
));
children.push(spacer(160));

children.push(h3('各欄閱讀要點', '1E3A5F'));
children.push(spacer(80));

for (const col of COLS) {
  const bcgColor = col.bcg === 'STAR' ? '60A5FA' : 'FB923C';
  const bcgLabel = col.bcg === 'STAR' ? '★ STAR' : '?  Question Mark';
  children.push(new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [new TextRun({ text: `${bcgLabel}  ${col.title}`, font: FONT, size: 22, bold: true, color: bcgColor })],
  }));
  children.push(bullet(`核心衝突：${col.key_argument}`));
  children.push(bullet(`最有力數字：${col.key_number}`));
  children.push(bullet(`InvosData 切入點：${col.invos_angle}`));
  children.push(spacer(60));
}

children.push(divider());

// 四、InvosData 解決方案
children.push(h1('四、InvosData 解決方案說明邏輯'));
children.push(body('Dashboard 最下方的 invosData 行是本次提案的商業轉換關鍵。以下是三個產品線應對應哪些客戶痛點：'));
children.push(spacer(120));
children.push(twoCol([
  ['產品', '對應的核心痛點 → 話術方向', true],
  ['invos\nInsight', '「你現在知道你的份額，但不知道競品在做什麼、什麼時候會追上來」\n→ 每月競爭監測，讓品牌從被動應對轉為主動偵測'],
  ['invos\nMedia',   `「你的廣告在打誰？有沒有打到真正會買的人？」\n→ 發票受眾包讓廣告對象從「猜測」變成「有實際購買記錄的真實消費者」`],
  ['invos\nAPI',     `「每賣一盒等於流失一個客戶——因為你不知道他是誰」\n→ 好市多匿名發票轉 CRM 名單，建立品牌可直接觸達的會員資產`],
], [1440, 7920]));
children.push(spacer(120));
children.push(callout(
  '升級路徑建議',
  `${BRAND}目前使用 invos Insight。建議本次提案目標：短期導入 invos Media（新客獲取），3–6 個月後導入 invos API（留客深耕）。不建議三個同時推，會讓客戶覺得超出預算規模。`,
  'FEF9EC', 'D97706'
));
children.push(divider());

// 五、常見問答
children.push(h1('五、常見客戶問題與建議回應'));
children.push(spacer(80));
children.push(twoCol([
  ['問題', '建議回應方向', true],
  ...COLS.map(col => [col.qa.q, col.qa.a]),
  [
    'Q: 這些數字來自哪裡，可信嗎？',
    '「所有數字來自 InvosData 彙整的台灣實購發票數據，覆蓋約 6,000+ 萬張年度發票，是真實的消費行為，不是問卷推估。」',
  ],
], [2520, 6840]));
children.push(divider());

// 六、結尾話術
children.push(h1('六、建議結尾話術'));
children.push(body('提案尾聲建議用「下一步很小、風險很低」的框架關閉對話：'));
children.push(spacer(80));
children.push(callout('建議結語（參考）', CLOSING_SCRIPT, 'F0FFF4', '16A34A'));
children.push(spacer(160));

// 附錄
children.push(h1('附錄：關鍵數據速查表'));
children.push(body('以下數字請以 Moya 或業務確認版為準（如有更新請覆蓋此表）：', { italic: true }));
children.push(spacer(80));
children.push(twoCol([
  ['指標', '數值', true],
  ...APPENDIX_DATA,
], [3240, 6120]));

// ── Build & save ──────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 240 } } } },
      ],
    }],
  },
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: '374151' } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: FONT },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: FONT },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: 'E2E8F0', space: 4 } },
        spacing: { before: 0, after: 160 },
        children: [new TextRun({ text: `${BRAND} ${PRODUCT}｜簡報閱讀邏輯`, font: FONT, size: 18, color: '94A3B8' })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { top: { style: BorderStyle.SINGLE, size: 3, color: 'E2E8F0', space: 4 } },
        spacing: { before: 160, after: 0 },
        children: [
          new TextRun({ text: 'InvosData 內部使用  ／  ', font: FONT, size: 18, color: '94A3B8' }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: '64748B' }),
        ],
      })] }),
    },
    children,
  }],
});

const outName = `${BRAND}_${PRODUCT_SHORT}_閱讀邏輯.docx`;
const outDir  = path.join(__dirname, '../../../output/reading-guide');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, outName);

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log(`✅  Saved: ${outPath}`);
}).catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
