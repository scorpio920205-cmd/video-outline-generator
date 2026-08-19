# 影音大綱與專業 Word 排版：跨 Agent 全域安裝說明

> 文件用途：把本文件交給任何具備本機檔案讀寫能力的 AI Agent。Agent 讀完後，應自動偵測這台電腦已安裝的 AI Agent，將「影音大綱與專業 Word 排版」規則安全加入各工具的使用者層級全域設定，完成備份、驗證並回報。

## 給使用者的一句話

把本文件交給 AI Agent，然後說：

> **「完整讀取這份文件，依照執行協定，把影音大綱與專業 Word 排版規則安裝到這台電腦所有已偵測到的 AI Agent 全域設定；保留原設定，完成後逐套驗證並回報。」**

---

# 一、Agent 執行契約

你是安裝 Agent。讀到本文件後，依序完成以下工作，不要只回覆教學步驟。

1. 完整讀取本文件。
2. 判斷目前作業系統與使用者家目錄；不得使用本文件作者的個人絕對路徑。
3. 偵測這台電腦已存在的 AI Agent。
4. 列出預計修改的設定檔及理由。
5. 修改前為每個既有檔案建立帶時間戳記的同目錄備份。
6. 以「受控區塊」方式合併規則，不覆蓋、不刪除使用者原有設定。
7. 已存在相同受控區塊時更新該區塊，不得重複附加。
8. 驗證每套已偵測工具都能載入規則。
9. 回報成功、跳過、需重啟或需使用者手動處理的項目，以及實際檔案位置。

## 執行邊界

- 僅修改「使用者層級／全域」設定，不修改目前專案的設定檔。
- 不安裝未偵測到的 AI Agent。
- 不覆蓋任何完整設定檔。
- 不改動模型、供應商、MCP、連接器、權限或登入設定。
- 不讀出、不複製、不顯示 API Key、Token、Cookie、OAuth 憑證或密碼。
- 若設定檔內含祕密，只進行必要的結構化局部修改；輸出紀錄不得顯示原檔全文。
- 若缺乏檔案寫入權限，停止該工具的修改，回報需要的授權；不要繞過系統安全機制。
- 一般 ChatGPT 網頁聊天或純雲端聊天不會自動讀取本機設定檔；不得假裝已安裝。只有能確認支援本機全域指令的 Agent 才列為成功。

---

# 二、要安裝的標準規則

以下受控區塊是唯一內容來源。安裝時應保留起訖標記。

```markdown
<!-- BEGIN:SENSEBAR_VIDEO_OUTLINE_DOCX_V1 -->
## 影音大綱提煉與專業 Word 排版規範

當使用者需要處理長影片、教學課程、會議錄影、線上講座之大綱整理，或要求輸出 Word (`.docx`) 講義/筆記時，一律遵循以下標準流程與設計規範：

### 1. 核心工作流程 (Workflow)
- **高速影格抽取 (Fast-seeking)**：若有影片本機檔案，優先建議使用 FFmpeg `-ss` 快轉跳轉抽取關鍵影格（間隔 60~90 秒），大幅節省處理時間。
- **精準時間戳對齊**：所有單元章節、核心觀點與關鍵步驟均需標註精確起訖時間戳（格式：`HH:MM:SS`）。
- **結構化提煉**：依據影片視覺畫面（投影片、簡報標題）與口述語音內容，歸納出：
  1. 影片元數據摘要
  2. 章節導覽與精確時間戳總表
  3. 各單元詳細內容剖析（重點條列、注意事項）
  4. 實戰作業 / 行動清單 (Action Items)

### 2. 專業 Word (.docx) 排版設計規範 (Design System)
產出 Word 文件時，必須採用商務級專業配色與清晰版面結構，嚴禁使用陽春預設樣式：

- **版面規範**：
  - 標準 A4 尺寸，頁邊距上下左右皆設為 **0.8 英吋 (57.6 pt)**。
  - 主要字體：**微軟正黑體** (Microsoft JhengHei)。
- **商務色彩體系 (Deep Navy & Ocean Blue)**：
  - **主標題 / 表頭背景**：深海藍 `#182B49` (RGB: 24, 43, 73)，字體 18~20pt 粗體。
  - **副標題 / 時間戳高亮**：海洋藍 `#2980B9` (RGB: 41, 128, 185)，字體 10.5~12pt。
  - **重點提醒 / 強調色**：古銅金 `#B8860B` (RGB: 184, 134, 11) 或 琥珀金。
  - **內文文字**：深炭灰 `#2C3E50` (RGB: 44, 62, 80)，字體 9.5pt，行距 1.15 倍。
  - **表格隔行交替底色**：白底 `#FFFFFF` 與 淺灰藍 `#F4F6F9` 交替。
- **必備文件組件**：
  1. **元數據摘要表**：以雙欄/四欄網格呈現主題、日期、時長、講師、核心工具。
  2. **時間戳總表**：表頭深藍底白字，欄位包含「時間戳 (起訖)」、「單元」、「核心主題」、「重點摘要」。
  3. **詳細剖析**：分單元條列核心重點，前綴標籤使用高亮色彩加粗（如 `【00:15:30】核心觀念：`）。
  4. **重點提示 Callout 框**：使用單格表格模擬 Callout，底色使用淺金 `#FDF8E2` 或淡藍 `#EBF3FA`，邊距內縮並加粗標題。

### 3. Python (`python-docx`) 自動化生成指引
若需要自動生成 `.docx` 文件，使用 `python-docx` 搭配 XML 樣式直接注入背景色 (`w:shd`)、邊距 (`w:tcMar`) 與邊框 (`w:tblBorders`)，以確保產出極致美觀的商務級文件。
<!-- END:SENSEBAR_VIDEO_OUTLINE_DOCX_V1 -->
```

## 正確行為範例

### 範例一：使用者要求為影片產出大綱與時間戳
- **Agent 行為**：依 `HH:MM:SS` 格式整理各章節時間戳，提供「章節總表」與「詳細內容剖析」，並詢問或直接產出對應結構的 JSON / Python 腳本匯出 Word。

### 範例二：使用者要求生成專業 Word 文件
- **Agent 行為**：自動採用深海藍（`#182B49`）與海洋藍（`#2980B9`）商務配色，配置 A4 0.8" 邊距、微軟正黑體、表頭底色與 Callout 方框，產出具備專業設計感的文檔。

---

# 三、受控區塊更新規則

對 Markdown 全域設定檔執行以下 upsert：

1. 若檔案不存在，建立父資料夾與檔案。
2. 若檔案存在，先建立備份：`原檔名.bak-YYYYMMDD-HHmmss`。
3. 搜尋起始標記：
   `<!-- BEGIN:SENSEBAR_VIDEO_OUTLINE_DOCX_V1 -->`
4. 搜尋結束標記：
   `<!-- END:SENSEBAR_VIDEO_OUTLINE_DOCX_V1 -->`
5. 起訖標記都存在：以本文件「標準規則」完整替換舊受控區塊。
6. 起訖標記都不存在：在檔案結尾保留一個空行後附加標準規則。
7. 只有一個標記存在：視為檔案結構異常，不得猜測範圍；停止修改該檔案並回報。
8. 修改完成後確認起始與結束標記各只有一個。
9. 保留原檔編碼；新建 Markdown 檔使用 UTF-8。

---

# 四、偵測與安裝位置

先以設定資料夾、可執行程式、應用程式安裝資訊等安全方式偵測。只要能確認該 Agent 已存在，就執行對應設定。

| Agent | 偵測線索 | 使用者層級全域設定 | 安裝方式 |
|---|---|---|---|
| ChatGPT 桌面應用中的 Codex／Codex CLI | `~/.codex/`、Codex CLI 或桌面應用的本機 Codex 工作環境 | `~/.codex/AGENTS.md` | 對 Markdown 執行受控區塊 upsert |
| Claude Code | `~/.claude/` 或 `claude` 命令 | 優先 `~/.claude/rules/video-outline-docx.md` | 建立獨立使用者規則檔；若目前版本不支援使用者 rules，才回退到 `~/.claude/CLAUDE.md` 執行 upsert |
| AntiGravity 2／Gemini Code Assist Agent Mode | `~/.gemini/`、AntiGravity 或 Gemini Code Assist 安裝資訊 | `~/.gemini/GEMINI.md` | 對 Markdown 執行受控區塊 upsert |
| OpenCode／OpenCode Desktop | `~/.config/opencode/`、`opencode` 命令或桌面應用 | 規則檔：`~/.config/opencode/instructions/video-outline-docx.md`；索引：`~/.config/opencode/opencode.json` | 建立規則檔，並把相對路徑加入 JSON 的 `instructions` 陣列 |

> `~` 代表目前使用者的家目錄。Windows 上應透過 `$HOME` 或系統 API 解析，不得硬編碼 `C:\Users\某人`。

## A. ChatGPT 桌面應用／Codex

1. 確認本機存在 Codex 工作環境；單純安裝一般 ChatGPT 聊天應用不等於支援本機 `AGENTS.md`。
2. 目標檔案：`~/.codex/AGENTS.md`。
3. 依「受控區塊更新規則」合併。
4. 不修改 `config.toml`、登入資訊、記憶資料庫或應用程式狀態檔。
5. 驗證時開啟新的 Codex 工作階段，確認全域 `AGENTS.md` 已載入。

## B. Claude Code

1. 優先建立：`~/.claude/rules/video-outline-docx.md`。
2. 此檔內容僅放完整受控區塊，不複製本文件的安裝說明。
3. 如果已存在同名檔，依 upsert 規則更新並備份。
4. 若已安裝的 Claude Code 版本無法載入使用者層級 `~/.claude/rules/`，才改為更新 `~/.claude/CLAUDE.md`。
5. 不修改 `settings.json`、Hooks、權限或登入憑證。
6. 驗證可使用 Claude Code 的 `/memory` 或等效診斷，確認規則檔已載入。

## C. AntiGravity 2／Gemini

1. 目標檔案：`~/.gemini/GEMINI.md`。
2. 依「受控區塊更新規則」合併。
3. 不修改 API Key、MCP、模型或登入設定。
4. 若 AntiGravity 版本提供可查詢的 Context／Memory 清單，使用它確認全域 `GEMINI.md` 已載入。

## D. OpenCode／OpenCode Desktop

1. 建立規則檔：
   `~/.config/opencode/instructions/video-outline-docx.md`
2. 規則檔內容僅放完整受控區塊。
3. 找到目前實際使用的全域設定檔，優先順序：
   - `~/.config/opencode/opencode.json`
   - 若前者不存在且版本明確支援，再使用 `~/.config/opencode/opencode.jsonc`
4. 修改 JSON 前建立備份。
5. 若頂層沒有 `instructions`，新增為陣列。
6. 若 `instructions` 已存在且是陣列，保留所有項目，只在尚未存在時加入：
   `./instructions/video-outline-docx.md`
7. 若 `instructions` 不是陣列、JSON 無法解析或檔案結構不明，停止修改並回報；不得覆寫整份設定。
8. 不輸出設定檔全文，避免洩露其中的 Token 或 API Key。
9. 驗證 JSON／JSONC 語法有效，且路徑在 `instructions` 中恰好出現一次。
10. 重新開啟 OpenCode 工作階段後測試規則。

---

# 五、依賴套件與工具提示

本技能在執行實際影片處理與 Word 產出時，建議具備以下環境（Agent 於安裝時可檢查並友善提示使用者）：

1. **Python 套件**：
   ```bash
   pip install python-docx
   ```
2. **FFmpeg / ffprobe**（用於影片快速抽格與時長分析）：
   - Windows: 可透過 `winget install Gyan.FFmpeg` 或 `choco install ffmpeg` 安裝。
   - macOS: 可透過 `brew install ffmpeg` 安裝。

---

# 六、驗證協定

## 檔案驗證

每套工具至少檢查：

- 目標設定檔或規則檔存在。
- 檔案不是零位元組。
- 起始與結束標記各恰好出現一次。
- 原有內容仍存在，沒有被整檔覆蓋。
- 備份檔存在（新建檔案不要求備份）。
- OpenCode 的設定檔語法有效，`instructions` 路徑恰好出現一次。

## 行為驗證

優先用全新工作階段測試，避免目前對話上下文造成假陽性。

測試提示詞：

> 「我有一部 2 小時的教學錄影需要整理重點，請列出大綱整理規範與 Word 產製之標準配色與結構要求，不要修改任何檔案。」

預期：

- Agent 應指出採用 `HH:MM:SS` 起訖時間戳、元數據摘要表、章節總表與詳細剖析結構。
- Agent 應列出深海藍（`#182B49`）、海洋藍（`#2980B9`）商務配色、A4 0.8" 邊距與微軟正黑體等排版要求。

---

# 七、完成回報格式

安裝完成後，使用以下格式回報：

```text
影音大綱與專業 Word 排版：跨 Agent 安裝結果

[成功／更新／跳過／失敗] ChatGPT 桌面應用／Codex
- 設定檔：<實際路徑>
- 備份：<實際路徑或新建檔案>
- 驗證：<結果>
- 重啟：<是否需要>

[成功／更新／跳過／失敗] Claude Code
- 設定檔：<實際路徑>
- 備份：<實際路徑或新建檔案>
- 驗證：<結果>
- 重啟：<是否需要>

[成功／更新／跳過／失敗] AntiGravity 2／Gemini
- 設定檔：<實際路徑>
- 備份：<實際路徑或新建檔案>
- 驗證：<結果>
- 重啟：<是否需要>

[成功／更新／跳過／失敗] OpenCode
- 規則檔：<實際路徑>
- 全域設定：<實際路徑>
- 備份：<實際路徑或新建檔案>
- 驗證：<結果>
- 重啟：<是否需要>

依賴環境檢測：
- Python python-docx：<已安裝／未安裝>
- FFmpeg：<已安裝／未安裝>

未完成事項：<沒有或具體說明>
```

回報中不得顯示任何憑證、Token、Cookie、API Key 或完整敏感設定內容。

---

# 八、解除安裝

使用者要求移除時：

1. 對 Codex、Gemini 或回退使用的 Claude `CLAUDE.md`，只刪除起訖標記及其間內容。
2. Claude 若使用獨立 rules 檔，只移除 `video-outline-docx.md`，不動其他 rules。
3. OpenCode 移除獨立規則檔，並只從 `instructions` 陣列移除對應路徑；保留其他 instructions。
4. 修改前仍要備份。
5. 驗證沒有殘留標記，也沒有破壞其他設定。

---

# 九、相容性說明

- Codex 的使用者全域指令通常位於 `~/.codex/AGENTS.md`。
- Claude Code 的使用者指令可放在 `~/.claude/CLAUDE.md`，新版也支援 `~/.claude/rules/` 的使用者層級規則。
- Gemini Code Assist Agent Mode 的全域 Context 檔位於 `~/.gemini/GEMINI.md`；AntiGravity 2 若沿用該機制可直接使用。
- OpenCode 可在全域 `opencode.json` 的 `instructions` 陣列載入額外規則檔。
- 各工具可能改版。若實際版本與本表不一致，安裝 Agent 應先使用該工具的內建診斷或官方文件確認，不得猜測後覆寫。
