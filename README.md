# Insight Report Skill

invosData AI Skills 集合 — 可重複使用的 prompt + 程式模板，搭配任何 AI 工具生成市場洞察報告。

> **給 AI 工具 / 同事使用 AI 工具的提醒**：當有人把本 repo 連結貼給 AI 工具時，第一步先**問使用者意圖**，不要直接假設。先丟二擇一問題：
>
> 1. **跑報告**（下游應用）— 產出 Word 洞察報告摘要 / 一頁式行動建議 / 閱讀邏輯手冊 / BCG PPTX（舊版迭代）。
>    接著依序問：①要跑哪一個 skill（**建議主力使用 `action-recommendation`（第三代、最新版）**；`bcg-analysis`（第一代）/ `bcg-channel-people`（第二代）是舊版迭代，仍可用，適合需要完整四頁簡報的場景）、②市場洞察分析報告原檔（PDF/Word/試算表）、③目標品牌、④品類與報告期間；若跑 `action-recommendation` 或 `reading-guide`，再問是否已跑過舊版 BCG，沿用欄位切法不要重切。
> 2. **改 skill / 修 bug** — 讀 `skills/<name>/SKILL.md` 與 `references/*.js`、改邏輯、提 PR。
>    接著依序問：①要動哪一個 skill、②是 prompt 邏輯、template 產檔錯誤還是新增功能、③具體範例或錯誤訊息。若是新增 skill，請依下方「新增 Skill 的方式」建檔並更新總覽表格。
>
> 兩種情境的下一步完全不同，先問清楚可避免做白工。注意：本 repo 是 prompt + pptxgenjs/docx 模板，不是 Google Apps Script 專案，AI 可直接代為執行 `node` 產檔。

> ⚠️ **PPTX 排版硬性規則（所有 PPTX skill 必遵守）：所有標題列一律限縮在一行內。**
> 包含欄位 header、象限卡、策略定位框的 title/subtitle/problem/goal，以及各 bullet 的粗體小標。
> 欄寬 2.54"、fontSize 8 時，中文 ≤ 15 字、英數混合 ≤ 20 字為安全線。超過就**簡化用字**，
> 不要縮字型、不要加大卡片、不要改行距硬塞。產檔後務必打開 PPTX 目測確認沒有任何一列折行。

> ⚠️ **推論內容硬性規則（所有 skill 必遵守）：超出原始數據的推論／延伸想像，產檔前必須先詢問使用者。**
> 洞察、行動建議、目標數字中，只要有「原始報告數據沒有直接支持」的段落，AI 必須在寫進檔案前主動標示：
> (1) 這段是推論不是數據 (2) 從什麼數據跳到什麼結論的分析邏輯 (3) 請使用者確認是否採用。
> **不可**把推論包裝成有數據支持的樣子直接寫進產出檔。
> 目標數字若非原始 benchmark，寧可只寫方向（「對標通路均值」「貼近市場水平」）也不要編造具體數字。
> 違反這條會讓產出檔對客戶失去可信度——客戶追問「這個數字從哪來」時答不出來。

## Skills 總覽

| Skill | 說明 | 輸出 | 使用順序 |
|-------|------|------|---------|
| [insight-report-summary](skills/insight-report-summary/) | 產出市場洞察分析報告摘要（四章 Word 文件） | .docx | 1. 先用這個 |
| [action-recommendation](skills/action-recommendation/) | 一頁式行動建議（第三代 BCG，最新版、推薦主力） | 兩頁 PPTX | 2. ⭐ 推薦用這個 |
| [bcg-analysis](skills/bcg-analysis/) | BCG 策略矩陣 — 品牌成長策略版（第一代，舊版迭代、仍可用） | 四頁 PPTX | 2a. 需完整四頁時用 |
| [bcg-channel-people](skills/bcg-channel-people/) | BCG 策略矩陣 — 通路佈局找對的人版（第二代，舊版迭代、仍可用） | 四頁 PPTX | 2b. 需完整四頁時用 |
| [reading-guide](skills/reading-guide/) | 產出簡報閱讀邏輯（內部業務提案參考手冊） | .docx | 3. 最後用這個 |

> **迭代說明**：BCG 系列已迭代三代 — `bcg-analysis`（第一代）→ `bcg-channel-people`（第二代）→ `action-recommendation`（第三代、最新版、推薦主力）。舊版仍保留可用，適合需要完整四頁簡報的場景。

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

### ② bcg-analysis — BCG 策略矩陣：品牌成長策略版（PPTX）

> 範例檔案：[`台塑生醫MD醫之方_BCG策略建議_通路新舊客版.pptx`](examples/台塑生醫MD醫之方_BCG策略建議_通路新舊客版.pptx)

四頁深色主題投影片，**讀者：品牌經理/行銷主管/決策層**：

| 頁面 | 內容 |
|------|------|
| **Slide 1：封面頁** | 深色背景 + 中央圖示 + 「行動建議」標題 |
| **Slide 2：BCG 矩陣圖** | 以「品牌掌握度 × 客群規模」為軸的 BCG 四象限 + 右側策略定位說明 |
| **Slide 3：當前瓶頸 × 核心機會** | 5 個客群欄位，每欄列出瓶頸與機會（各附報告頁碼） |
| **Slide 4：行動建議 × invosData 方案** | 5 個客群欄位，每欄列出行動建議與 invosData 服務方案 |

---

### ②b bcg-channel-people — BCG 策略矩陣：通路佈局找對的人版（PPTX）

> 範例檔案：[`台塑生醫MD醫之方_BCG策略建議_通路找人版.pptx`](examples/台塑生醫MD醫之方_BCG策略建議_通路找人版.pptx)

四頁深色主題投影片，**讀者：通路業務/通路行銷/電商營運/事業部主管**：

| 頁面 | 內容 |
|------|------|
| **Slide 1：封面頁** | 深色背景 + 中央圖示 + 「行動建議」標題 |
| **Slide 2：通路×客群 BCG 矩陣圖** | 以「客群掌握度 × 通路規模與潛力」為軸，象限用直覺語言（如「最知道誰在買，但人正在走」） |
| **Slide 3：為什麼找不到對的人 × 如何用數據找到** | 5 欄（momo/康是美/寶雅/舊客/競品客），瓶頸指向「人找不到」，機會指向「數據怎麼找」 |
| **Slide 4：通路佈局行動方案 × invosData 方案** | 產品×通路配對行動 + 「幫你在每個通路找到對的人」服務方案 |

**兩版差異：** 品牌成長策略版用在前期提案/高階會議（由上往下看全局）；通路佈局版用在執行團隊工作坊（由下往上看落地）

---

### ③ action-recommendation — 一頁式行動建議（PPTX）

> 範例檔案：[`台塑生醫MD醫之方_行動建議.pptx`](examples/台塑生醫MD醫之方_行動建議.pptx)

三頁深色主題投影片，**讀者：客戶決策者／採購窗口**，把 BCG 簡報的「當前瓶頸 × 核心機會 × invosData 解決方案」濃縮到單頁：

| 頁面 | 內容 |
|------|------|
| **Slide 1：封面頁** | 深色背景 + 中央圖示 + 「行動建議」標題 |
| **Slide 2：一頁式整合頁** | 5 欄 × 3 列：Header 目標卡（動詞+通路/客群+對標市場的目標）／關鍵資訊（突破性針對性洞察）／行動建議（invosData 可做的事）＋底部類別 tag |
| **Slide 3：附數據來源版** | 與 Slide 2 內容 100% 一致，每段末尾加 `(P##)` 頁碼或 `【策略推論】`／`【invos 服務】`／`【策略方向】` 標籤（淡灰色），讓讀者能逐段追溯出處 |

**三大核心要素（每欄都要齊）：**
1. 突破性與針對性的關鍵資訊（非通論、對該客群獨有）
2. 對標整體市場或通路的關鍵目標（每個數字都有 benchmark）
3. 針對關鍵目標的 invosData 行動建議（可執行、可歸因）

**Slide 3 實作機制：** template 定義 `COLUMNS`（5 欄內容）與 `SOURCES`（5 欄對應來源）兩個陣列，AI 撰寫 COLUMNS 時需同步填 SOURCES，產檔時自動綁定渲染。

**使用時機：** 本 skill 為**第三代 BCG、最新版、推薦主力**。可直接獨立使用（skill 會依報告原檔內建切 5 欄客群/通路）；若先跑過舊版 bcg-analysis / bcg-channel-people，欄位切法沿用 bcg 不重切，避免客戶看到兩份不一致。

---

### ④ reading-guide — 簡報閱讀邏輯（Word）

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

**推薦流程（最新版）：** `insight-report-summary`（Word 洞察報告）→ `action-recommendation`（一頁式行動建議，第三代）→ `reading-guide`（內部提案手冊）

**舊版迭代流程（仍可用，需要完整四頁簡報時）：** `insight-report-summary` → `bcg-analysis`（第一代）或 `bcg-channel-people`（第二代）→ `action-recommendation`（沿用 bcg 切法濃縮單頁）→ `reading-guide`

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
│   ├── 台塑生醫MD醫之方_BCG策略建議_通路找人版.pptx
│   ├── 台塑生醫MD醫之方_行動建議.pptx
│   └── 威德_睡眠益生菌_閱讀邏輯.docx
├── skills/
│   ├── insight-report-summary/            # ① 市場洞察分析報告摘要（Word）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── docx-template.js
│   ├── bcg-analysis/                      # ② BCG 策略矩陣：品牌成長策略版（PPTX）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── pptx-template.js
│   ├── bcg-channel-people/                # ②b BCG 策略矩陣：通路佈局找對的人版（PPTX）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── pptx-template.js
│   ├── action-recommendation/             # ③ 一頁式行動建議（PPTX，濃縮 BCG 給客戶看）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── pptx-template.js
│   ├── reading-guide/                     # ④ 簡報閱讀邏輯（內部提案手冊）
│   │   ├── SKILL.md
│   │   └── references/
│   │       └── generate_guide.js
│   └── <your-next-skill>/                 # 未來新增的 skill
│       ├── SKILL.md
│       └── references/
└── output/                                # 生成的報告（gitignored）
    ├── insight-report-summary/
    ├── bcg-analysis/
    ├── bcg-channel-people/
    ├── action-recommendation/
    └── reading-guide/
```

## 新增 Skill 的方式

1. 在 `skills/` 下建立新資料夾（例如 `skills/competitor-report/`）
2. 新增 `SKILL.md`：AI 指令本體，描述分析流程和輸出格式
3. 新增 `references/`：放程式碼模板、設計素材等參考檔案
4. 更新本 README 的 Skills 總覽表格
