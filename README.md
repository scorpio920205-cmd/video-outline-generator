# 🎬 影音大綱與專業排版 Word 生成技能 (Video Outline Generator)

[![Antigravity Skill](https://img.shields.io/badge/Antigravity-Skill-blue.svg)](https://github.com/scorpio920205-cmd/video-outline-generator)
[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

專為 **長影片、教學課程、線上講座與會議錄影** 設計的結構化大綱整理與商務級 Word (`.docx`) 自動排版生成技能（Skill）。

整合 **FFmpeg 高速影格擷取（Fast Seeking）**、時間戳對齊、結構化章節提煉，以及基於 `python-docx` 的精美深海藍商務風格 Word 排版引擎，一鍵生成媲美專業顧問等級的講義與筆記文檔。

---

## 🌟 核心特色 (Key Features)

- ⚡ **FFmpeg 高速影格擷取**：透過 `-ss` 快速跳轉技術，2~3 小時的高畫質長影片僅需 2~3 分鐘即可自動抽取關鍵影格，省去逐格解碼的等待。
- 🧭 **精確起訖時間戳對齊**：章節單元、核心觀點、操作步驟均標註 `HH:MM:SS` 時間戳，方便快速回放對照。
- 📄 **商務級專業 Word 排版**：
  - **色彩體系**：深海藍（`#182B49`）、海洋藍（`#2980B9`）、古銅金（`#B8860B`）、炭灰內文（`#2C3E50`）。
  - **版面規範**：標準 A4、微軟正黑體、0.8 英吋優化邊距。
  - **內建組件**：影片元數據摘要表、章節起訖總表（雙色隔行交替）、分單元深度條列解析、重點任務 Callout 提示框。
- 🤖 **跨 Agent 一鍵安裝**：隨附 `AGENT_SETUP_影音大綱與專業Word排版_跨Agent全域安裝.md`，可一鍵無損安裝至 Claude Code、ChatGPT/Codex、Antigravity、OpenCode 全域環境。

---

## 📁 專案目錄結構 (Project Structure)

```text
video-outline-generator/
├── SKILL.md                                           # AI Agent 技能核心規範與標準作業流程
├── AGENT_SETUP_影音大綱與專業Word排版_跨Agent全域安裝.md   # 跨 Agent 一鍵全域安裝文件（可直接分享給任何人）
├── README.md                                          # 專案詳細說明與快速上手文件
├── requirements.txt                                   # Python 依賴清單
├── LICENSE                                            # MIT 授權條款
├── .gitignore                                         # Git 忽略設定
├── scripts/
│   ├── extract_video_frames.py                        # FFmpeg 高速影格抽取腳本
│   └── generate_docx_outline.py                       # 專業 Word (.docx) 生成引擎
└── templates/
    ├── outline_data_schema.json                       # Word 生成器所使用的 JSON 資料結構範本
    └── outline_template.md                            # Markdown 格式大綱寫作模板
```

---

## 🎁 跨 Agent 一鍵全域安裝指南 (Cross-Agent Universal Setup)

如果您想將此技能規則安裝到電腦上的所有 AI Agent（例如 Claude Code、ChatGPT/Codex、AntiGravity、OpenCode），只需將 [`AGENT_SETUP_影音大綱與專業Word排版_跨Agent全域安裝.md`](./AGENT_SETUP_影音大綱與專業Word排版_跨Agent全域安裝.md) 檔案交給您的 AI Agent，並說：

> **「完整讀取這份文件，依照執行協定，把影音大綱與專業 Word 排版規則安裝到這台電腦所有已偵測到的 AI Agent 全域設定；保留原設定，完成後逐套驗證並回報。」**

Agent 將自動完成安全備份、受控區塊合併與設定載入驗證！

---

## 🚀 快速上手 (Quick Start)

### 1. 環境需求
- Python 3.9+
- [FFmpeg](https://ffmpeg.org/)（需加入系統 PATH 環境變數）

### 2. 安裝依賴
```bash
pip install -r requirements.txt
```

### 3. 步驟一：高速擷取影片關鍵影格
使用 Fast Seeking 技術每 60 秒擷取一張關鍵影格：
```bash
python scripts/extract_video_frames.py --video "path/to/your_video.mp4" --interval 60 --output "./frames"
```
參數說明：
- `--video`: 輸入影片檔案路徑（支援 MP4, MKV, MOV, AVI 等）。
- `--interval`: 擷取間隔秒數（預設 60 秒）。
- `--scale`: 輸出縮圖寬度（預設 1280px）。
- `--output`: 影格輸出資料夾（預設 `./frames`）。

### 4. 步驟二：整理大綱資料 (JSON)
參考 `templates/outline_data_schema.json` 建立您的大綱資料：

```json
{
  "title": "2026年2月 - 選擇權收入策略與量化交易實戰",
  "subtitle": "全片精確時間戳、四大模組核心精華、交易心法與實戰操作完整大綱",
  "metadata": {
    "課程主題": "選擇權收入策略 (Iron Condor) 與量化實戰",
    "影片總長": "03:00:38 (180 分鐘)",
    "主講導師": "修老師、Fiona Wei、Perry 老師、Vega Ko",
    "核心工具": "TradingView, IBKR TWS, OptionStrat, QuantConnect"
  },
  "chapters": [
    {
      "timestamp": "00:00:00 - 00:05:40",
      "unit": "單元一",
      "title": "開場引言與學習進度檢核",
      "summary": "說明團輔定位、檢核學習進度與社群互助。"
    }
  ],
  "sections": [
    {
      "heading": "第一部分：開場引言與量化思維",
      "bullets": [
        {
          "prefix": "【00:00:00 - 00:05:40】團輔定位：",
          "text": "強調每月團體輔導旨在拓展多元投資策略..."
        }
      ]
    }
  ]
}
```

### 5. 步驟三：自動匯出專業排版 Word 檔
```bash
python scripts/generate_docx_outline.py --input "outline_data.json" --output "課程大綱與時間戳.docx"
```

---

## 🎨 Word 排版設計規範 (Design System)

| 元素 | 色彩 / 規格 | 說明 |
| :--- | :--- | :--- |
| **主標題** | `#182B49` (Deep Navy), 20pt 粗體 | 醒目、穩重之商務主視覺 |
| **副標題** | `#2980B9` (Ocean Blue), 12pt | 次要說明與課程目標 |
| **表格表頭** | `#182B49` 深底白字, 9.5pt 粗體 | 清晰劃分資料欄位 |
| **表格交替色** | `#F4F6F9` (淺灰藍) / `#FFFFFF` | 提升長表格易讀性 |
| **高亮時間戳** | `#2980B9` (Ocean Blue), 粗體 | 關鍵起訖時間一目了然 |
| **內文主體** | `#2C3E50` (Charcoal), 9.5pt, 1.15 倍行距 | 高對比、舒適長時間閱讀 |

---

## 📄 授權條款 (License)

本專案採用 [MIT License](LICENSE) 授權。
