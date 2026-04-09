# Insight Report Skill

invosData AI Skills 集合 — 可重複使用的 prompt + 程式模板，搭配任何 AI 工具生成市場洞察報告。

## Skills 總覽

| Skill | 說明 | 輸出 | 使用順序 |
|-------|------|------|---------|
| [insight-report-summary](skills/insight-report-summary/) | 產出市場洞察分析報告摘要（四章 Word 文件） | .docx | 1. 先用這個 |
| [bcg-analysis](skills/bcg-analysis/) | 產出 BCG 策略矩陣分析（通路 x 新舊客視角） | 三頁 PPTX | 2. 再用這個 |
| [reading-guide](skills/reading-guide/) | 產出簡報閱讀邏輯（內部業務提案參考手冊） | .docx | 3. 最後用這個 |

## 輸出範例

每個 skill 的實際產出範例可在 [`examples/`](examples/) 資料夾下載。以下說明各 skill 會生成什麼內容：

### ① insight-report-summary — 市場洞察分析報告摘要（Word）

> 範例檔案：[`台塑生醫MD醫之方益生菌市場洞察分析報告摘要_invos.docx`](examples/台塑生醫MD醫之方益生菌市場洞察分析報告摘要_invos.docx)

四章結構的 Word 文件，約 10–15 頁：

| 章節 | 內容 |
|------|------|
| **第一章：市場整體現況** | 市場成長結構（銷售額/購買人數/人均次數/單次金 + YoY）、各階表現、通路分佈與競爭格局、品牌競爭態勢、新舊客 NES 分析 |
| **第二章：品牌專屬市場洞察** | 5 則針對目標品牌的獨特洞察，每則含「數據發現 → 市場洞察 → 突破方向」 |
| **第三章：invos 三大服務策略建議** | invos Insight / Media / API 各 3–4 個具體方案，含受眾圈選、投放時機、預期效果 |
| **第四章：策略重點一覽** | 三大服務 × 各 3 條核心行動的彙總表格 |

---

### ② bcg-analysis — BCG 策略矩陣分析（PPTX）

> 範例檔案：[`台塑生醫MD醫之方_BCG策略建議_通路新舊客版.pptx`](examples/台塑生醫MD醫之方_BCG策略建議_通路新舊客版.pptx)

三頁深色主題投影片：

| 頁面 | 內容 |
|------|------|
| **Slide 1：BCG 矩陣圖** | 以「通路 × 新舊客」為軸的 BCG 四象限定位圖 + 右側策略定位文字說明 |
| **Slide 2：當前瓶頸 × 核心機會** | 5 個客群欄位，每欄列出 BCG 標籤、關鍵指標、當前瓶頸與核心機會 |
| **Slide 3：行動建議 × invosData 方案** | 5 個客群欄位，每欄列出具體行動建議與對應的 invosData 服務方案 |

---

### ③ reading-guide — 簡報閱讀邏輯（Word）

> 範例檔案：[`威德_睡眠益生菌_閱讀邏輯.docx`](examples/威德_睡眠益生菌_閱讀邏輯.docx)

六章結構的 Word 文件，供內部業務人員提案前閱讀：

| 章節 | 內容 |
|------|------|
| **第一章：文件用途** | 說明這份手冊的目的與使用情境 |
| **第二章：整體敘事架構** | 三頁簡報的故事線與核心訊息串接邏輯 |
| **第三章：逐頁閱讀邏輯** | 每頁每欄的解讀方式、重點數字、銜接話術 |
| **第四章：invosData 解決方案話術** | Insight / Media / API 各服務的推薦說法 |
| **第五章：常見客戶問答** | 預期客戶會問的問題與建議回應 |
| **第六章：結尾話術 + 數據速查表** | 結尾收束語 + 附錄關鍵數據快速參考 |

---

## 快速開始

### 方式一：作為 AI Prompt 使用（推薦）

1. 打開任何支援程式碼執行的 AI 工具（Claude、ChatGPT、Cursor 等）
2. 將目標 skill 的 `SKILL.md` 內容貼入對話
3. 上傳你的市場數據報告，並說明報告期間、品類、目標品牌，例如：
   > 此為 2025/12 MAT 益生菌市場洞察分析報告，請針對品牌「台塑生醫MD醫之方」做分析
4. AI 會依照步驟自動分析並生成報告

**建議流程：** `insight-report-summary`（Word 洞察報告）→ `bcg-analysis`（BCG 策略 PPTX）→ `reading-guide`（內部提案手冊）

### 方式二：手動修改 template 執行

```bash
# 安裝依賴
npm install

# 以 bcg-analysis 為例：修改 references/pptx-template.js 中的品牌資料，然後執行
node skills/bcg-analysis/references/pptx-template.js

# 輸出位置：output/bcg-analysis/<品牌名>.pptx
```

## Repo 結構

```
insight-report-skill/
├── README.md                              # 本文件
├── package.json                           # 共用依賴
├── .gitignore
├── examples/                              # 各 skill 的實際產出範例
│   ├── 台塑生醫MD醫之方益生菌市場洞察分析報告摘要_invos.docx
│   ├── 台塑生醫MD醫之方_BCG策略建議_通路新舊客版.pptx
│   └── 威德_睡眠益生菌_閱讀邏輯.docx
├── skills/
│   ├── insight-report-summary/            # ① 市場洞察分析報告摘要（Word）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── docx-template.js
│   ├── bcg-analysis/                      # ② BCG 策略矩陣分析（PPTX）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── pptx-template.js
│   ├── reading-guide/                     # ③ 簡報閱讀邏輯（內部提案手冊）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── generate_guide.js
│   └── <your-next-skill>/                 # 未來新增的 skill
│       ├── SKILL.md
│       └── references/
└── output/                                # 生成的報告（gitignored）
    ├── insight-report-summary/
    ├── bcg-analysis/
    └── reading-guide/
```

## 新增 Skill 的方式

1. 在 `skills/` 下建立新資料夾（例如 `skills/competitor-report/`）
2. 新增 `SKILL.md`：AI 指令本體，描述分析流程和輸出格式
3. 新增 `references/`：放程式碼模板、設計素材等參考檔案
4. 更新本 README 的 Skills 總覽表格
