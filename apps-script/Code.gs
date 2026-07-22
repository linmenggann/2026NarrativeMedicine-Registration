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
  '時間戳記', '投稿者姓名', '人事號', '院區', '服務單位', '聯絡電話/分機', 'E-mail', '參賽類別',
  '短影音-編號', '短影音-作品名稱', '短影音-文字說明', '短影音-AI協作', '短影音-AI程式/網站', '短影音-檔案連結',
  '攝影-編號', '攝影-作品名稱', '攝影-文字說明', '攝影-AI協作', '攝影-AI程式/網站', '攝影-檔案連結',
  '徵文-編號', '徵文-標題', '徵文-簡介', '徵文-檔案連結'
];

const CAT_HEADERS = {
  v: ['編號', '時間戳記', '投稿者姓名', '人事號', '院區', '服務單位', '聯絡電話/分機', 'E-mail', '作品名稱', '文字說明', 'AI協作', 'AI程式/網站', '檔案連結'],
  p: ['編號', '時間戳記', '投稿者姓名', '人事號', '院區', '服務單位', '聯絡電話/分機', 'E-mail', '作品名稱', '文字說明', 'AI協作', 'AI程式/網站', '檔案連結'],
  w: ['編號', '時間戳記', '投稿者姓名', '人事號', '院區', '服務單位', '聯絡電話/分機', 'E-mail', '標題', '簡介', '檔案連結']
};

// ===================== 進入點 =====================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'initUpload') return json_(initUpload_(data));
    if (data.action === 'resolveFile') return json_(resolveFile_(data));
    if (data.action === 'stats') return json_(stats_());
    return json_(register_(data));
  } catch (err) {
    return json_({ result: 'error', message: String(err && err.message ? err.message : err) });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'stats') return json_(stats_());
  return json_({ result: 'ok', message: '2026 敘事醫學競賽報名 API 運作中' });
}

// ===================== 報名現況統計（僅彙整數字，不含個資） =====================

function stats_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tz = Session.getScriptTimeZone();

  // 各類別件數與 AI 協作數
  const cats = {};
  ['v', 'p', 'w'].forEach(function (k) {
    const sheet = ss.getSheetByName(CAT_TABS[k]);
    let count = 0, ai = 0;
    if (sheet) {
      count = Math.max(0, sheet.getLastRow() - 1);
      if (k !== 'w' && count > 0) {
        const aiCol = CAT_HEADERS[k].indexOf('AI協作') + 1;
        if (aiCol > 0) {
          const vals = sheet.getRange(2, aiCol, count, 1).getValues();
          ai = vals.filter(function (r) { return String(r[0]).trim() === '是'; }).length;
        }
      }
    }
    cats[k] = { name: CAT_TABS[k], count: count, ai: ai };
  });

  // 總表彙整：總筆數、依單位、依日期（不含姓名/電話/E-mail）
  const master = ss.getSheetByName(MASTER_SHEET);
  let total = 0;
  const byUnitMap = {}, byDateMap = {}, byBranchMap = {};
  if (master) {
    total = Math.max(0, master.getLastRow() - 1);
    if (total > 0) {
      const tsCol = HEADERS.indexOf('時間戳記');
      const unitCol = HEADERS.indexOf('服務單位');
      const branchCol = HEADERS.indexOf('院區');
      const numCols = ['短影音-編號', '攝影-編號', '徵文-編號'].map(function (h) { return HEADERS.indexOf(h); });
      const data = master.getRange(2, 1, total, master.getLastColumn()).getValues();
      data.forEach(function (row) {
        // 以「件數」統計（一筆報名報多類＝多件）
        let n = 0;
        numCols.forEach(function (i) { if (i >= 0 && String(row[i] || '').trim() !== '') n++; });
        const unit = (String(row[unitCol] || '').trim()) || '未填';
        byUnitMap[unit] = (byUnitMap[unit] || 0) + n;
        if (branchCol >= 0) {
          const branch = (String(row[branchCol] || '').trim()) || '未填';
          byBranchMap[branch] = (byBranchMap[branch] || 0) + n;
        }
        let dateStr = '';
        const tsv = row[tsCol];
        if (tsv instanceof Date) dateStr = Utilities.formatDate(tsv, tz, 'yyyy-MM-dd');
        else dateStr = String(tsv || '').substring(0, 10);
        if (dateStr) byDateMap[dateStr] = (byDateMap[dateStr] || 0) + 1;
      });
    }
  }
  const byUnit = Object.keys(byUnitMap).map(function (u) { return { unit: u, count: byUnitMap[u] }; })
    .sort(function (a, b) { return b.count - a.count; });
  const byBranch = Object.keys(byBranchMap).map(function (u) { return { branch: u, count: byBranchMap[u] }; })
    .sort(function (a, b) { return b.count - a.count; });
  const byDate = Object.keys(byDateMap).map(function (d) { return { date: d, count: byDateMap[d] }; })
    .sort(function (a, b) { return a.date < b.date ? -1 : 1; });

  return {
    result: 'success',
    generatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
    deadline: '2026-07-31',
    total: total,
    categories: cats,
    byUnit: byUnit,
    byBranch: byBranch,
    byDate: byDate
  };
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

    // 先取得各類別分頁與其即將分配的編號（供總表一併標示，方便交叉對照）
    const catSheets = {}, nums = { v: '', p: '', w: '' };
    ['v', 'p', 'w'].forEach(function (k) {
      if (!works[k]) return;
      const sheet = getSheet_(CAT_TABS[k], CAT_HEADERS[k]);
      catSheets[k] = sheet;
      nums[k] = sheet.getLastRow(); // 表頭已存在 → 等於現有資料列數＋1，即本筆編號（從 1 起）
    });

    // 總表（含各類別編號）
    getSheet_(MASTER_SHEET, HEADERS).appendRow([
      tsStr, data.name || '', empNoText_(data.empNo), data.branch || '',
      data.unit || '', data.phone || '', data.email || '',
      (data.categories || []).join('、'),
      nums.v, pick_(works, 'v', 'title'), pick_(works, 'v', 'desc'),
      (works.v && works.v.ai) ? '是' : '', pick_(works, 'v', 'aiName'), links.v,
      nums.p, pick_(works, 'p', 'title'), pick_(works, 'p', 'desc'),
      (works.p && works.p.ai) ? '是' : '', pick_(works, 'p', 'aiName'), links.p,
      nums.w, pick_(works, 'w', 'title'), pick_(works, 'w', 'desc'), links.w
    ]);

    // 分類別分頁（含自動編號）＋ 依編號為雲端檔案改名
    ['v', 'p', 'w'].forEach(function (k) {
      if (!works[k]) return;
      const sheet = catSheets[k];
      const num = nums[k];
      if (k === 'w') {
        sheet.appendRow([
          num, tsStr, data.name || '', empNoText_(data.empNo), data.branch || '',
          data.unit || '', data.phone || '', data.email || '',
          pick_(works, 'w', 'title'), pick_(works, 'w', 'desc'), links.w
        ]);
      } else {
        sheet.appendRow([
          num, tsStr, data.name || '', empNoText_(data.empNo), data.branch || '',
          data.unit || '', data.phone || '', data.email || '',
          pick_(works, k, 'title'), pick_(works, k, 'desc'),
          (works[k] && works[k].ai) ? '是' : '', pick_(works, k, 'aiName'), links[k]
        ]);
      }
      // 編號-作品名稱(徵文用主題/標題)-投稿者姓名(-序號).副檔名
      renameFilesByNo_(works[k].files, num, pick_(works, k, 'title'), data.name);
    });

    return { result: 'success', message: '報名成功' };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ===================== 輔助函式 =====================

/**
 * 依編號為已上傳的雲端檔案改名（改名不會改變檔案連結/ID）。
 * 命名：編號-作品名稱-投稿者姓名(-序號).副檔名（同一筆有多檔時才加序號以避免同名）。
 * 例：1-醫者仁心-王小明.mp4
 */
function renameFilesByNo_(files, no, title, applicant) {
  if (!files || !files.length) return;
  const clean = function (s) { return String(s || '').replace(/[\\/:*?"<>|]/g, '').trim(); };
  const t = clean(title) || '作品';
  const a = clean(applicant) || '匿名';
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    if (!f || !f.id) continue;
    try {
      const name = f.name || '';
      const ext = (name.lastIndexOf('.') >= 0) ? name.substring(name.lastIndexOf('.')) : '';
      const seq = (files.length > 1) ? ('-' + (i + 1)) : '';
      DriveApp.getFileById(f.id).setName(no + '-' + t + '-' + a + seq + ext);
    } catch (err) { /* 單檔改名失敗則略過，不影響報名 */ }
  }
}

// ===================== 一次性工具：刪列後重新編號 =====================
// 分頁有刪除報名列時，在編輯器執行對應函式一次：
//   重排該分頁「編號」為 1..N（依現有列順序遞補）→ 依新編號重新命名雲端檔案
//   → 同步總表對應的「類別-編號」欄（以時間戳記＋投稿者姓名比對；
//     總表中已無對應分頁列者，其編號欄會清空）。
// 注意：刪列後若不重編，下一筆新報名的編號會與現有最後一筆重複。

function renumberV() { return renumberCategory_('v'); } // 醫療人文短影音
function renumberP() { return renumberCategory_('p'); } // 醫療人文攝影
function renumberW() { return renumberCategory_('w'); } // 醫療人文敘事醫學徵文

function renumberCategory_(k) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(CAT_TABS[k]);
  if (!sheet) return '找不到分頁：' + CAT_TABS[k];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '分頁無資料，毋須重編。';

  const headers = CAT_HEADERS[k];
  const iNo = headers.indexOf('編號');
  const iTs = headers.indexOf('時間戳記');
  const iName = headers.indexOf('投稿者姓名');
  const iTitle = headers.indexOf(k === 'w' ? '標題' : '作品名稱');
  const iLinks = headers.indexOf('檔案連結');
  const tz = Session.getScriptTimeZone();
  const tsKey = function (v) {
    return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss') : String(v || '').trim();
  };
  const clean = function (s) { return String(s || '').replace(/[\\/:*?"<>|]/g, '').trim(); };

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const keyToNewNo = {}; // 時間戳記|姓名 → 新編號
  let renamed = 0, renumbered = 0;

  data.forEach(function (row, idx) {
    const newNo = idx + 1;
    keyToNewNo[tsKey(row[iTs]) + '|' + String(row[iName] || '').trim()] = newNo;
    if (Number(row[iNo]) === newNo) return; // 編號已正確，略過
    renumbered++;
    sheet.getRange(idx + 2, iNo + 1).setValue(newNo);
    // 依新編號重新命名雲端檔案（連結/檔案 ID 不變，僅改顯示名稱）
    const links = String(row[iLinks] || '').split('\n').map(function (s) { return s.trim(); }).filter(String);
    const t = clean(row[iTitle]) || '作品';
    const a = clean(row[iName]) || '匿名';
    links.forEach(function (url, i) {
      const m = url.match(/\/d\/([-\w]+)/);
      if (!m) return;
      try {
        const f = DriveApp.getFileById(m[1]);
        const name = f.getName();
        const ext = (name.lastIndexOf('.') >= 0) ? name.substring(name.lastIndexOf('.')) : '';
        const seq = (links.length > 1) ? ('-' + (i + 1)) : '';
        f.setName(newNo + '-' + t + '-' + a + seq + ext);
        renamed++;
      } catch (err) { /* 檔案不存在或無權限則略過 */ }
    });
  });

  // 同步總表「類別-編號」欄
  const numHeader = { v: '短影音-編號', p: '攝影-編號', w: '徵文-編號' }[k];
  const master = ss.getSheetByName(MASTER_SHEET);
  let synced = 0, cleared = 0;
  if (master && master.getLastRow() > 1) {
    const mTs = HEADERS.indexOf('時間戳記');
    const mName = HEADERS.indexOf('投稿者姓名');
    const mNo = HEADERS.indexOf(numHeader);
    const mData = master.getRange(2, 1, master.getLastRow() - 1, HEADERS.length).getValues();
    mData.forEach(function (row, idx) {
      if (String(row[mNo] === 0 ? '0' : (row[mNo] || '')).trim() === '') return; // 未報此類別
      const key = tsKey(row[mTs]) + '|' + String(row[mName] || '').trim();
      if (key in keyToNewNo) {
        if (Number(row[mNo]) !== keyToNewNo[key]) { master.getRange(idx + 2, mNo + 1).setValue(keyToNewNo[key]); synced++; }
      } else {
        master.getRange(idx + 2, mNo + 1).setValue(''); cleared++; // 分頁已刪除該筆 → 清空總表編號
      }
    });
  }

  const msg = CAT_TABS[k] + '：重編 ' + renumbered + ' 列、改名 ' + renamed + ' 個檔案、總表同步 ' + synced + ' 筆、清空 ' + cleared + ' 筆（已刪列者）。';
  Logger.log(msg);
  return msg;
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    // 人事號整欄設為純文字，避免如 8607E7 被解讀成科學記號
    const empIdx = headers.indexOf('人事號');
    if (empIdx >= 0) sheet.getRange(1, empIdx + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  }
  return sheet;
}

/** 人事號以文字寫入（前置單引號，儲存格顯示不含引號），雙重保險防止科學記號轉換 */
function empNoText_(v) {
  const s = String(v || '').trim();
  return s ? ("'" + s) : '';
}

/**
 * 【一次性工具】修正既有分頁的人事號欄格式為純文字。
 * 若總表/分類別分頁在此版之前已建立，請在編輯器手動執行本函式一次，
 * 讓既有分頁的人事號整欄改為純文字（僅影響之後寫入；已被轉成科學記號的舊值需手動更正）。
 */
function fixEmpNoFormat() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const targets = [
    { name: MASTER_SHEET, headers: HEADERS },
    { name: CAT_TABS.v, headers: CAT_HEADERS.v },
    { name: CAT_TABS.p, headers: CAT_HEADERS.p },
    { name: CAT_TABS.w, headers: CAT_HEADERS.w }
  ];
  targets.forEach(function (t) {
    const sheet = ss.getSheetByName(t.name);
    if (!sheet) return;
    const empIdx = t.headers.indexOf('人事號');
    if (empIdx >= 0) sheet.getRange(1, empIdx + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  });
  Logger.log('人事號欄已設為純文字格式。');
  return 'OK';
}

function pick_(works, k, field) {
  return (works[k] && works[k][field]) ? works[k][field] : '';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
