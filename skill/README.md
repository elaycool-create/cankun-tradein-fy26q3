# 📦 Trade-in 週報 Skill

一鍵把 Apple ITLS Trade-in 匯出 Excel → 互動式 HTML 報告。

## 📁 目錄結構

```
cankun-tradein-skill/
├── SKILL.md              ← 完整規格、欄位定義、使用情境
├── analyze_and_report.py ← 主程式（單一檔，自包含）
└── README.md             ← 本檔，快速上手
```

## ⚡ 30 秒快速使用

```bash
# 1. 安裝 pandas
pip3 install pandas openpyxl

# 2. 跑分析
python3 analyze_and_report.py \
  --input ~/Desktop/Trade-in原始檔.xlsx \
  --output ~/Desktop/週報.html

# 3. 用瀏覽器開報告
open ~/Desktop/週報.html
```

## 🔄 增量更新（合併多週）

```bash
python3 analyze_and_report.py \
  --input 舊檔W01-W06.xlsx 新檔W07.xlsx \
  --output 週報_W01-W07.html \
  --export-json    # 同時輸出 JSON 供下次用
```

## 🎨 各種使用情境

| 想做什麼 | 指令 |
|---------|------|
| 燦坤週報（預設） | `--input data.xlsx` |
| 改成分析德誼 | `--dealer 德誼` |
| 改成分析晶實 | `--dealer 晶實` |
| 從 6/7 開始算第一週 | `--start-date 2026-06-07` |
| 只看 4 週 | `--num-weeks 4` |
| 自動推算到資料最後一週 | （省略 `--num-weeks` 即可） |

## 🚀 推上 GitHub Pages（公開連結）

```bash
# 第一次：建 repo + 啟用 Pages
gh repo create my-tradein --public
cd my-tradein
cp ../週報.html ./index.html
git init && git add . && git commit -m "init" && git push -u origin main
gh api repos/:owner/my-tradein/pages -f source[branch]=main -f source[path]=/

# 之後每次更新
cp ../週報.html ./index.html
git add . && git commit -m "update" && git push
```

公開網址：`https://<github_username>.github.io/my-tradein/`

## 📚 詳細說明

- 完整指標定義 → 看 [SKILL.md](./SKILL.md)
- 程式邏輯 → 看 [analyze_and_report.py](./analyze_and_report.py) 註解

## 🤝 客製建議

需要新增欄位／調整顏色／加總圖表 → 編輯 `analyze_and_report.py` 中的 `HTML_TEMPLATE` 與 `render_html()` 即可，所有資料前後端串接都在這個檔案。
