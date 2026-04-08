---
name: bcg-analysis
description: >
  invosData 品牌策略 BCG 矩陣分析工具。當使用者上傳任何嬰幼兒奶粉（或快消品）市場數據報告 PPTX，
  並希望生成 BCG 策略分析簡報時，立即啟動此 skill。觸發語包含：
  「幫我做 BCG 分析」、「生成策略簡報」、「分析這份報告」、「做成像上次那樣的 PPT」、
  「上傳新報告」、「新的 MAT 報告」、「分析市場數據」、「做引客數據提案簡報」，
  或使用者直接上傳 .pptx 報告並提及品牌名稱時。
  輸出為三頁 PPTX：(1) BCG 矩陣圖 (2) 當前瓶頸×核心機會 (3) 行動建議×invosData 方案。
---

# invosData BCG 策略矩陣分析 Skill

## 你的任務

使用者上傳了一份市場數據報告 PPTX，要求你生成一份針對特定品牌的 BCG 策略分析三頁投影片。

**固定輸出：三頁深色主題 PPTX（用 pptxgenjs 生成）**
- Slide 1：BCG 矩陣圖 + 右側策略定位說明
- Slide 2：當前瓶頸 × 核心機會（5 欄 × 2 列）
- Slide 3：行動建議 × invosData 數據解決方案（5 欄 × 2 列）

---

## Step 1：讀取並分析報告

用 `python3` + `zipfile` 提取報告 PPTX 所有頁面的文字：

```bash
python3 -c "
import zipfile, re, sys
path = '<uploaded_pptx_path>'
with zipfile.ZipFile(path) as z:
    slides = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)])
    for i, s in enumerate(slides, 1):
        xml = z.read(s).decode('utf-8', errors='ignore')
        text = re.sub(r'<[^>]+>', ' ', xml)
        text = re.sub(r'\s+', ' ', text).strip()
        print(f'=== P{i} ===')
        print(text[:1200])
        print()
"
```

閱讀全部頁面後，整理出以下資訊（記錄每個數據的頁碼，之後寫進投影片用）：

| 分析維度 | 要找的資訊 |
|---------|-----------|
| 整體市場 | 市場規模、成長率、新舊客人流比例 |
| 通路分佈 | 各通路（藥局婦嬰/電商/藥妝）銷售額與人流 |
| 品牌競爭 | 目標品牌 vs 主要競品市佔、成長趨勢 |
| 新客表現 | 品類新客數量、哪個品牌截獲最多新客 |
| 舊客表現 | 回購率、購買次數、客單價（單次金） |
| 產品線 | 各 SKU/系列的買家數、銷售額、升級路徑 |
| 電商 | 各平台滲透率、非授權通路（如酷澎）問題 |
| 競品交叉 | 有沒有跨品牌購買（交叉購買）資料 |

---

## Step 2：建立品牌洞察框架

以「**通路 × 新舊客視角**」定義 BCG 的五個欄位（column）：

| 欄 | 標題 | 副標 | BCG 象限 |
|----|------|------|---------|
| 1 | 藥局婦嬰通路整體 | 大樹 / 卡多摩 / 丁丁（或當地主力通路） | Cash Cow |
| 2 | 品類新客 | 新生兒父母・首次選牌 | Question Mark |
| 3 | 品牌舊客 × 升級 | [主力SKU]->[高端SKU]・留客深化消費 | Star |
| 4 | 競品轉換客 | [主要競品]->[目標品牌] | Star |
| 5 | 電商第二通路 | momo / 蝦皮官方旗艦 | Question Mark |

> 若報告顯示通路結構不同（如無電商數據），可調整欄位定義，但保持 5 欄結構。

**BCG 矩陣軸線定義（此框架特有，非傳統定義）：**
- **X 軸（左→右）**：品牌對該客群的掌握度（低→高）
- **Y 軸（上→下）**：現有規模大但發展受限 → 規模小但成長潛力高（進攻）

象限位置：
```
Dogs         | Cash Cow
（掌握低、守）  （掌握高、守）
─────────────┼─────────────
Question Mark | Star
（掌握低、攻）  （掌握高、攻）
```

---

## Step 3：填寫三頁投影片的內容

在生成程式碼之前，先依資料確定每個 cell 的文字內容。

### Slide 1：BCG 四象限 + 右側說明框

**四個象限各寫 2 個要點（粗體標題 + 說明句）**，從報告中抓具體數字：

| 象限 | 要點 1 | 要點 2 |
|------|--------|--------|
| Dogs | 通路缺席/衰退事實 | 實體通路集中風險 |
| Cash Cow | 最大人流通路現況 | 穩定回購 SKU + 客單 |
| Question Marks | 競品截獲新客速度 | 電商能見度不足 |
| Stars | 舊客升級 LTV 潛力 | 競品交叉購買識別 |

**右側四個策略說明框（每框：標題 + 一句描述 + 2 條關鍵資訊）**

### Slide 2：當前瓶頸 × 核心機會

每個欄 × 每個列 填入 **3 條 bullet**（各附頁碼引用 `報告 p.X`）：

**當前瓶頸**：用數據說明「現在有什麼問題」
**核心機會**：用數據說明「哪裡有機會可以解鎖」

每條 bullet 都要：
1. 有具體數字或事實（不能只是通論）
2. 對目標品牌有針對性（不能直接套用給其他品牌）
3. 附上頁碼：`["句子內容", "報告 p.X"]`

### Slide 3：行動建議 × invosData 方案

**行動建議**（每欄 3 條）：具體的執行步驟，從「分析」到「投放」到「留存」

**invosData 解決方案**（每欄 2-3 個服務模組）：

invosData 三大服務固定定義：
- **invos Insight**（綠色 `16A34A`）：市場機會洞察 — 賦能品牌智慧決策，建立可落地執行策略
- **invos Media**（青色 `0891B2`）：精準獲客轉換 — 高效鎖客，快速驗證，驅動全場景行銷擴散
- **invos API**（紫色 `7C3AED`）：全通路會員整合 — 將單次消費顧客加入私域流量池，個人化經營

每個解決方案 bullet 格式：`服務名 + 具體執行內容（50-80 字）`

---

## Step 4：生成 PPTX

讀取本 skill 目錄下的 `references/pptx-template.js` 作為程式碼基礎。

```bash
# 在 repo 根目錄安裝依賴
npm install

# 複製 template 並修改品牌資料後執行
node skills/bcg-analysis/references/pptx-template.js
```

**修改 template 時必須保持的設計規範：**

| 項目 | 規格 |
|------|------|
| 投影片尺寸 | LAYOUT_WIDE（13.3" x 7.5"）|
| 背景色 | `13182A`（深藍黑）|
| 主要文字色 | `E2E8F0`（偏白）|
| 頂部色條 | 橘（`F97316`）+ 金（`F59E0B`）|
| 欄數 | 固定 5 欄，X0=0.2, CW=(13.3-0.4)/5 |
| 每條 bullet 頁碼 | inline 小字灰斜體，格式 `[句子+"  ", "報告 p.X"]` tuple |
| 輸出路徑 | repo 根目錄下 `output/bcg-analysis/<品牌>_BCG策略建議_<date>.pptx` |

---

## Step 5：驗證輸出

```python
import zipfile, re
path = '<output_pptx_path>'
with zipfile.ZipFile(path) as z:
    slides = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)]
    print(f'Slides: {len(slides)}')  # 應為 3
    xml = z.read('ppt/slides/slide2.xml').decode('utf-8', errors='ignore')
    refs = re.findall(r'報告 p\.[^<\"]{1,20}', xml)
    print(f'Page refs in slide2: {len(refs)}')  # 應為 30
```

確認：
- [ ] 三頁投影片存在
- [ ] Slide 2 共 30 條頁碼引用（瓶頸 15 + 機會 15）
- [ ] 品牌名稱正確（非上一個品牌殘留）
- [ ] 輸出存在 `output/` 路徑下

---

## 輸出最後步驟

將生成的 PPTX 檔案提供給使用者，並附上簡短摘要：
- 三頁分別分析了什麼
- Slide 2 有多少條引用了哪幾頁
- 是否有任何數據信心度較低的部分需要使用者確認
