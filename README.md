# 2026 醫院敘事醫學系列競賽活動 — 線上報名投稿網頁

奇美醫療體系 教學部主辦《醫療的瞬間・永恆的感動》視覺化報名與投稿單頁網站。

## 內容
- [`index.html`](index.html)：單一檔案網頁（報名／投稿、作品上傳）。
- [`apps-script/Code.gs`](apps-script/Code.gs)：Google Apps Script 後端，將報名資料寫入 Google Sheets、作品檔案存入 Google 雲端硬碟。
- [`apps-script/SETUP.md`](apps-script/SETUP.md)：表頭欄位與部署串接步驟。

## 串接後端（送出資料 → Google Sheets／雲端硬碟）
1. 依 [`apps-script/SETUP.md`](apps-script/SETUP.md) 將 `Code.gs` 部署為「網頁應用程式」。
2. 把部署網址填入 `index.html` 最上方 `<script>` 內的 `const APPS_SCRIPT_URL = '';`。
3. 重新整理網頁即可送出報名與上傳作品。

> 報名資料寫入試算表分頁「競賽活動報名資料」；作品檔案依類別存入「醫療人文短影音／攝影／敘事醫學徵文」雲端資料夾。
