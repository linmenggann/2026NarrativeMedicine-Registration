/**
 * 2026 醫院敘事醫學系列競賽 — 線上報名後端 (Google Apps Script)
 * ------------------------------------------------------------------
 * 功能：
 *   1. 接收網頁送出的報名資料，寫入 Google Sheets 分頁「競賽活動報名資料」。
 *   2. 接收各類別上傳的作品檔案，依類別存入對應的 Google 雲端硬碟資料夾。
 *
 * 部署方式請參閱同資料夾的 SETUP.md。
 */

// ===================== 設定區（已依您提供的網址填入 ID）=====================

// Google Sheets ID 與分頁名稱
const SHEET_ID   = '1Mh0p8pgsAtkRdTDO7MOjPmr0E3BBsGeb_Q3Jt3M7AKU';
const SHEET_NAME = '競賽活動報名資料';

// 各類別作品上傳的 Google 雲端硬碟資料夾 ID
const FOLDERS = {
  v: '1OlZVUttznJiJoiFunnXjRU4RCLYrXYXz', // 醫療人文短影音
  p: '1D97iXfjjxbJDdCzpzzk6uAMX3AxsU0HE', // 醫療人文攝影
  w: '1Nv5PhZeRiJxFqz1FY5erm8iV82e3NHn2'  // 醫療人文敘事醫學徵文
};

// 類別中文名稱
const CAT_NAME = {
  v: '醫療人文短影音',
  p: '醫療人文攝影',
  w: '醫療人文敘事醫學徵文'
};

// Google Sheets 表頭（第一列）— 如分頁為空會自動建立
const HEADERS = [
  '時間戳記',
  '投稿者姓名',
  '服務單位',
  '聯絡電話/分機',
  'E-mail',
  '參賽類別',
  '短影音-作品名稱',
  '短影音-文字說明',
  '短影音-AI協作',
  '短影音-AI程式/網站',
  '短影音-檔案連結',
  '攝影-作品名稱',
  '攝影-文字說明',
  '攝影-AI協作',
  '攝影-AI程式/網站',
  '攝影-檔案連結',
  '徵文-標題',
  '徵文-簡介',
  '徵文-檔案連結'
];

// ===================== 主要進入點 =====================

/** 接收網頁 POST 的報名資料與作品檔案 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const data  = JSON.parse(e.postData.contents);
    const works = data.works || {};
    const ts    = new Date();
    const links = { v: '', p: '', w: '' };

    // 1) 各類別作品檔案存入對應雲端硬碟資料夾
    ['v', 'p', 'w'].forEach(function (k) {
      const work = works[k];
      if (!work || !work.files || !work.files.length) return;
      const folder = DriveApp.getFolderById(FOLDERS[k]);
      const urls = [];
      work.files.forEach(function (file, idx) {
        const bytes = Utilities.base64Decode(file.data);
        const blob  = Utilities.newBlob(
          bytes,
          file.mimeType || 'application/octet-stream',
          buildFileName_(data, k, file, idx, ts)
        );
        const saved = folder.createFile(blob);
        urls.push(saved.getUrl());
      });
      links[k] = urls.join('\n');
    });

    // 2) 報名資料寫入 Google Sheets
    const sheet = getSheet_();
    sheet.appendRow([
      Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      data.name  || '',
      data.unit  || '',
      data.phone || '',
      data.email || '',
      (data.categories || []).join('、'),
      pick_(works, 'v', 'title'),
      pick_(works, 'v', 'desc'),
      (works.v && works.v.ai) ? '是' : '',
      pick_(works, 'v', 'aiName'),
      links.v,
      pick_(works, 'p', 'title'),
      pick_(works, 'p', 'desc'),
      (works.p && works.p.ai) ? '是' : '',
      pick_(works, 'p', 'aiName'),
      links.p,
      pick_(works, 'w', 'title'),
      pick_(works, 'w', 'desc'),
      links.w
    ]);

    return json_({ result: 'success', message: '報名成功' });
  } catch (err) {
    return json_({ result: 'error', message: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** 提供瀏覽器直接開啟測試（GET）*/
function doGet() {
  return json_({ result: 'ok', message: '2026 敘事醫學競賽報名 API 運作中' });
}

// ===================== 輔助函式 =====================

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function pick_(works, k, field) {
  return (works[k] && works[k][field]) ? works[k][field] : '';
}

/** 產生易辨識的檔名：日期時間_姓名_作品名稱(_序號).副檔名 */
function buildFileName_(data, k, file, idx, ts) {
  const stamp = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const clean = function (s) { return String(s || '').replace(/[\\/:*?"<>|]/g, '').trim(); };
  const name  = clean(data.name) || '匿名';
  const title = clean(pick_(data.works || {}, k, 'title')) || '作品';
  const dot   = (file.name && file.name.lastIndexOf('.') >= 0) ? file.name.substring(file.name.lastIndexOf('.')) : '';
  const seq   = idx > 0 ? ('_' + (idx + 1)) : '';
  return stamp + '_' + name + '_' + title + seq + dot;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
