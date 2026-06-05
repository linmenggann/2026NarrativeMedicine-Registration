/**
 * 2026 醫院敘事醫學系列競賽 — 線上報名後端 (Google Apps Script)
 * ------------------------------------------------------------------
 * 兩種請求（皆 POST，內容為 JSON 字串）：
 *   A. { action:'initUpload', category, fileName, mimeType }
 *      → 以主辦者身分向 Google Drive 申請「可續傳上傳網址」，回傳 { uploadUrl }。
 *        前端再把檔案直接 PUT 到該網址（瀏覽器→Google），可上傳大檔案（>50MB），
 *        繞過 Apps Script 單次請求約 50MB 的限制。
 *   B. { name, unit, phone, email, categories, works }  （無 action）
 *      → 報名資料寫入 Google Sheets（總表＋分類別分頁）。
 *        works[k].files 為已上傳完成的檔案清單 [{ name, id, url }]。
 *
 * 部署方式請參閱同資料夾的 SETUP.md。
 * 注意：本版新增了 UrlFetchApp 對外請求與 Drive 上傳，重新部署時會要求重新授權。
 */

// ===================== 設定區 =====================

const SHEET_ID = '1Mh0p8pgsAtkRdTDO7MOjPmr0E3BBsGeb_Q3Jt3M7AKU';
const MASTER_SHEET = '競賽活動報名資料';

const CAT_TABS = {
  v: '醫療人文短影音',
  p: '醫療人文攝影',
  w: '醫療人文敘事醫學徵文'
};

const FOLDERS = {
  v: '1OlZVUttznJiJoiFunnXjRU4RCLYrXYXz', // 醫療人文短影音
  p: '1D97iXfjjxbJDdCzpzzk6uAMX3AxsU0HE', // 醫療人文攝影
  w: '1Nv5PhZeRiJxFqz1FY5erm8iV82e3NHn2'  // 醫療人文敘事醫學徵文
};

const HEADERS = [
  '時間戳記', '投稿者姓名', '服務單位', '聯絡電話/分機', 'E-mail', '參賽類別',
  '短影音-作品名稱', '短影音-文字說明', '短影音-AI協作', '短影音-AI程式/網站', '短影音-檔案連結',
  '攝影-作品名稱', '攝影-文字說明', '攝影-AI協作', '攝影-AI程式/網站', '攝影-檔案連結',
  '徵文-標題', '徵文-簡介', '徵文-檔案連結'
];

const CAT_HEADERS = {
  v: ['時間戳記', '投稿者姓名', '服務單位', '聯絡電話/分機', 'E-mail', '作品名稱', '文字說明', 'AI協作', 'AI程式/網站', '檔案連結'],
  p: ['時間戳記', '投稿者姓名', '服務單位', '聯絡電話/分機', 'E-mail', '作品名稱', '文字說明', 'AI協作', 'AI程式/網站', '檔案連結'],
  w: ['時間戳記', '投稿者姓名', '服務單位', '聯絡電話/分機', 'E-mail', '標題', '簡介', '檔案連結']
};

// ===================== 進入點 =====================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'initUpload') return json_(initUpload_(data));
    if (data.action === 'resolveFile') return json_(resolveFile_(data));
    return json_(register_(data));
  } catch (err) {
    return json_({ result: 'error', message: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return json_({ result: 'ok', message: '2026 敘事醫學競賽報名 API 運作中' });
}

/**
 * 【部署後請先在編輯器執行這個函式一次】
 * 用途：觸發「對外連線（UrlFetchApp）」與「雲端硬碟（Drive）」權限的授權同意視窗。
 * 執行 → 點「檢閱權限」→ 選帳號 →（若提示未驗證：進階→前往專案）→ 允許。
 * 完成後正式網頁即可申請上傳網址、上傳大檔案，不需再重新部署。
 */
function authorize() {
  UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis', { muteHttpExceptions: true });
  DriveApp.getRootFolder();
  Logger.log('授權完成：已具備對外連線與 Drive 權限。');
  return 'OK';
}

// ===================== A. 申請可續傳上傳網址 =====================

function initUpload_(data) {
  const folderId = FOLDERS[data.category];
  if (!folderId) return { result: 'error', message: '未知的類別：' + data.category };

  // 驗證資料夾存在，同時確保已授權 Drive 範圍（讓 getOAuthToken 取得 Drive 權杖）
  DriveApp.getFolderById(folderId);

  const meta = { name: data.fileName || ('作品_' + Date.now()), parents: [folderId] };
  if (data.mimeType) meta.mimeType = data.mimeType;

  const res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'X-Upload-Content-Type': data.mimeType || 'application/octet-stream'
      },
      payload: JSON.stringify(meta),
      muteHttpExceptions: true
    }
  );

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    return { result: 'error', message: '建立上傳工作階段失敗（' + code + '）' };
  }
  const headers = res.getAllHeaders();
  const uploadUrl = headers['Location'] || headers['location'];
  if (!uploadUrl) return { result: 'error', message: '未取得上傳網址' };

  return { result: 'success', uploadUrl: uploadUrl };
}

// ===================== A-2. 上傳後依檔名查回檔案連結 =====================
// （Google 可續傳上傳的 PUT 回應不含 CORS 標頭，瀏覽器讀不到回傳的 file id，
//   故上傳完成後改由前端帶「檔名」來向後端查詢該檔案的連結。）

function resolveFile_(data) {
  const folderId = FOLDERS[data.category];
  if (!folderId) return { result: 'error', message: '未知的類別：' + data.category };
  const it = DriveApp.getFolderById(folderId).getFilesByName(data.fileName || '');
  if (it.hasNext()) {
    const f = it.next();
    return { result: 'success', id: f.getId(), url: f.getUrl() };
  }
  return { result: 'error', message: '找不到剛上傳的檔案：' + (data.fileName || '') };
}

// ===================== B. 寫入報名資料 =====================

function register_(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const works = data.works || {};
    const ts    = new Date();
    const tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // 由前端已上傳的檔案清單組出連結（每列一個）
    const links = { v: '', p: '', w: '' };
    ['v', 'p', 'w'].forEach(function (k) {
      const w = works[k];
      if (!w || !w.files || !w.files.length) return;
      links[k] = w.files.map(function (f) { return f && f.url ? f.url : ''; })
                        .filter(String).join('\n');
    });

    // 總表
    getSheet_(MASTER_SHEET, HEADERS).appendRow([
      tsStr, data.name || '', data.unit || '', data.phone || '', data.email || '',
      (data.categories || []).join('、'),
      pick_(works, 'v', 'title'), pick_(works, 'v', 'desc'),
      (works.v && works.v.ai) ? '是' : '', pick_(works, 'v', 'aiName'), links.v,
      pick_(works, 'p', 'title'), pick_(works, 'p', 'desc'),
      (works.p && works.p.ai) ? '是' : '', pick_(works, 'p', 'aiName'), links.p,
      pick_(works, 'w', 'title'), pick_(works, 'w', 'desc'), links.w
    ]);

    // 分類別分頁
    ['v', 'p', 'w'].forEach(function (k) {
      if (!works[k]) return;
      const sheet = getSheet_(CAT_TABS[k], CAT_HEADERS[k]);
      if (k === 'w') {
        sheet.appendRow([
          tsStr, data.name || '', data.unit || '', data.phone || '', data.email || '',
          pick_(works, 'w', 'title'), pick_(works, 'w', 'desc'), links.w
        ]);
      } else {
        sheet.appendRow([
          tsStr, data.name || '', data.unit || '', data.phone || '', data.email || '',
          pick_(works, k, 'title'), pick_(works, k, 'desc'),
          (works[k] && works[k].ai) ? '是' : '', pick_(works, k, 'aiName'), links[k]
        ]);
      }
    });

    return { result: 'success', message: '報名成功' };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ===================== 輔助函式 =====================

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function pick_(works, k, field) {
  return (works[k] && works[k][field]) ? works[k][field] : '';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
