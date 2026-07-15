# Lydia Toolbox — 專案說明

## 本機架構
這是一個 Flask 本機應用（port 5001），同時也 host 靜態 HTML 供 GitHub Pages 使用。

```
packing-app/
├── app.py                  Flask 主程式（API routes）
├── config.json             API key 設定（不可 push，已在 .gitignore）
├── templates/              Flask HTML 頁面
│   ├── cbp_v2.html         CBP 報關品名助手（主要版本）
│   ├── packing_list.html   Packing List 產生器（靜態前端，GitHub Pages 用）
│   ├── currency_v2.html    匯率換算
│   ├── orders.html         訂單篩選
│   └── home.html           首頁
├── scripts/
│   └── build_packing_list.py   Excel Packing List 產生腳本
├── data/
│   └── items.json          料號資料庫（169 筆 SCOSCHE 料號）
├── static/
│   └── nav.js              側邊導覽列
└── assets/
    └── SCOSCHE_PACKING_TEMPLATE.xlsx   Excel 模版
```

## GitHub repo
https://github.com/lydiachen000428/Lydia-toolbox
branch: main

## 同步規則
- 本機改完後，commit 並 push 到 GitHub main branch
- config.json 絕對不可以 push（public repo，含 API key）
- 每次 commit message 用中文簡短說明改了什麼

## 常見指令
```bash
# 查看改了哪些檔案
git status

# 把所有修改加入並 commit
git add .
git commit -m "說明改了什麼"
git push origin main

# 只 push 特定檔案
git add templates/cbp_v2.html
git commit -m "CBP 頁面：加入對話框功能"
git push origin main
```

## 注意事項
- `packing_list.html` 同時用於本機 Flask 和 GitHub Pages
- CBP 頁面的 AI 功能透過本機 `/api/cbp_analyze` proxy，GitHub Pages 靜態版無法使用 AI 功能
- Excel 產生功能（`build_packing_list.py`）只在本機 Flask 環境運作
- 含棧板的料號（has_pallet: true）會自動產生第二頁 WITH PLT 工作表
