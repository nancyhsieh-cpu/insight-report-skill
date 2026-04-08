# Insight Report Skill

invosData AI Skills 集合 — 可重複使用的 prompt + 程式模板，搭配任何 AI 工具生成市場洞察報告。

## Skills 總覽

| Skill | 說明 | 輸出 |
|-------|------|------|
| [bcg-analysis](skills/bcg-analysis/) | BCG 策略矩陣分析（通路 x 新舊客視角） | 三頁 PPTX |

## 快速開始

### 方式一：作為 AI Prompt 使用（推薦）

1. 打開任何支援程式碼執行的 AI 工具（Claude、ChatGPT、Cursor 等）
2. 將目標 skill 的 `SKILL.md` 內容貼入對話
3. 上傳你的市場數據報告，並說明報告期間、品類、目標品牌，例如：
   > 此為 2025/12 MAT 益生菌市場洞察分析報告，請針對品牌「台塑生醫MD醫之方」做 BCG 分析
4. AI 會依照步驟自動分析並生成報告

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
├── skills/
│   ├── bcg-analysis/                      # BCG 策略矩陣分析
│   │   ├── SKILL.md                       #   AI 指令（核心 prompt）
│   │   └── references/                    #   程式碼模板 & 參考資料
│   │       └── pptx-template.js
│   └── <your-next-skill>/                 # 未來新增的 skill
│       ├── SKILL.md
│       └── references/
└── output/                                # 生成的報告（gitignored）
    └── bcg-analysis/
```

## 新增 Skill 的方式

1. 在 `skills/` 下建立新資料夾（例如 `skills/competitor-report/`）
2. 新增 `SKILL.md`：AI 指令本體，描述分析流程和輸出格式
3. 新增 `references/`：放程式碼模板、設計素材等參考檔案
4. 更新本 README 的 Skills 總覽表格
