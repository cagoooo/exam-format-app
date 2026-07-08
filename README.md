# 考卷格式自動校正系統

這是一個本機 Flask 網頁 APP，可上傳 `.docx` 或 `.doc` 考卷，重新套用固定頁面大小、雙欄、欄距、頁邊距、表頭、字型、行距與題號縮排，輸出穩定的 `.docx`。適用單位預設為「桃園市龍潭區石門國民小學」。

## 啟動

```powershell
python app.py
```

開啟：

```text
http://127.0.0.1:5127
```

## 目前支援

- GitHub Pages 線上版：可直接上傳 Word 考卷並呼叫 Cloud Run API 轉換。
- Cloud Run API：使用 LibreOffice headless 處理 `.doc/.docx` 與 PDF 輸出。
- `.docx`：直接讀取內容並輸出校正後 Word。
- `.doc`：若本機已安裝 Microsoft Word，會透過 Word COM 自動轉為 `.docx`；若沒有 Word，請先手動另存成 `.docx`。
- 格式 profile：B4 橫式雙欄、B4 直式雙欄、A4 橫式雙欄、A4 直式雙欄。
- 模板設定：格式已外部化至 `config/profiles.json`。
- 圖片保真：讀取 DOCX media 並重新插入輸出檔，超過欄寬時等比例安全縮放。
- 表格保真：優先保留來源表格 XML，包含合併儲存格、框線、底色、對齊與表格內圖片 relationship。
- 公式保真：偵測 Word OMML 公式段落並保留原始 XML，避免公式被純文字抽取破壞。
- 注音保真：偵測國語 ruby 注音段落並保留原始 XML，避免 `<w:ruby>` 標記流失。
- 超頁診斷：分析與轉換報告會提示長段落、寬表格、大圖片、公式或注音段落等高風險跑版來源。
- 兩頁鎖定第二版：依壓縮等級調整字級、行距、欄距、頁邊距、圖片最大寬度、表格內距與表格行距。
- 轉換前後差異摘要：報告會列出系統實際套用的格式調整、圖片策略與表格策略。
- 大題不斷頁：大題標題盡量與下一段同頁，短題目段落盡量不拆行，表格列盡量不跨頁切開。
- PDF 預覽與頁數檢查：若本機 Microsoft Word 可用，轉換後會同步輸出 PDF、計算頁數並產生第一頁預覽。
- 兩頁鎖定模式：可設定目標頁數，系統會嘗試調整字級、行距、欄距與邊界。
- 基礎保真：保留段落中的粗體、斜體與底線。
- 批次驗證：可一次檢查五份範例檔的頁面尺寸、邊界、欄數與欄距。
- 分享預覽：已設定 favicon、Apple touch icon、PWA manifest、LINE/Facebook Open Graph 與 Twitter Card meta。
- Service Worker 更新提示：偵測到新版時會主動顯示「重新整理」提示，載入最新版本。
- 線上檔案安全：每次轉換使用獨立 job id，下載與預覽連結約 30 分鐘內有效。
- 線上端到端測試：可用腳本確認 Cloud Run API 的分析、轉換、下載仍正常。

## 社群分享與 GitHub Pages

OG 圖片位於：

```text
static/assets/og-exam-format.png
```

favicon 與 app icon 位於：

```text
static/assets/
```

若部署到 GitHub Pages，請設定公開網址環境變數，讓 `og:image` 產生絕對網址：

```powershell
$env:PUBLIC_SITE_URL="https://cagoooo.github.io/<repo-name>/"
python app.py
```

若之後改成純靜態匯出，請保留 `static/assets/` 與 `static/manifest.webmanifest`，並確認 HTML 中 `og:image` 是完整的 `https://.../static/assets/og-exam-format.png?v=...`。

## Service Worker 版本更新

相關檔案：

```text
sw.js
version.json
static/sw-update.js
scripts/bump-version.ps1
```

更新版本號：

```powershell
.\scripts\bump-version.ps1 -Notes "描述這次更新"
```

機制：

- `sw.js` 使用 `BUILD_VERSION` 建立版本化 cache。
- `version.json` 提供前端輪詢比對。
- `static/sw-update.js` 監聽 `updatefound`、`controllerchange`、`SW_ACTIVATED`，並在偵測到新版時提示使用者重新整理。
- Flask 的 `/sw.js` 與 `/version.json` 會回傳 `no-store`，避免本機開發時被快取卡住。

## 線上網址

GitHub Pages：

```text
https://cagoooo.github.io/exam-format-app/
```

Cloud Run API：

```text
https://exam-format-api-142975838924.asia-east1.run.app
```

## 測試

```powershell
python -m py_compile app.py scripts\batch_verify_examples.py
python scripts\batch_verify_examples.py
python scripts\generate_brand_assets.py
node --check static\app.js
node --check static\sw-update.js
python scripts\e2e_online_api.py
```

若要輸出 JSON：

```powershell
python scripts\batch_verify_examples.py --json
```

## 進度與開發路線

詳細進度表、P0/P1 清單與未來功能建議請見：

```text
PROGRESS_AND_ROADMAP.md
```

## 後續可擴充

- 保留圖片、公式與注音 ruby 的進階轉換規則。
- 建立更多科目與年級專用模板。
- 加入多人登入、歷史紀錄與校內部署。
