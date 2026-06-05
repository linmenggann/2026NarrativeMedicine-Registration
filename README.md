# 2026 醫院敘事醫學系列競賽活動 — 線上報名投稿網頁

奇美醫療體系 教學部主辦《醫療的瞬間・永恆的感動》視覺化報名與投稿單頁網站。

## 內容
- [`index.html`](index.html)：單一檔案網頁（報名／投稿、作品上傳、上傳進度條）。
- [`config.js`](config.js)：前端設定檔，填入後端網址 `APPS_SCRIPT_URL`。
- [`apps-script/Code.gs`](apps-script/Code.gs)：Google Apps Script 後端，將報名資料寫入 Google Sheets、作品檔案存入 Google 雲端硬碟。
- [`apps-script/SETUP.md`](apps-script/SETUP.md)：表頭欄位與部署串接步驟。

## 串接後端（送出資料 → Google Sheets／雲端硬碟）
1. 依 [`apps-script/SETUP.md`](apps-script/SETUP.md) 將 `Code.gs` 部署為「網頁應用程式」。
2. 設定後端網址（擇一）：
   - 填入 [`config.js`](config.js) 的 `APPS_SCRIPT_URL`（建議），或
   - 開啟網頁時加上網址參數 `?api=你的/exec網址`（會記住於瀏覽器）。
3. 重新整理網頁即可送出報名與上傳作品，過程會顯示上傳進度條。

> 報名資料寫入試算表分頁「競賽活動報名資料」；作品檔案依類別存入「醫療人文短影音／攝影／敘事醫學徵文」雲端資料夾。
