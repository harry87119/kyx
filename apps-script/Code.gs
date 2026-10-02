/**
 * @OnlyCurrentDoc
 * KYX記帳：Google 試算表後端
 * （上面那行讓授權時只要求存取「這一份」試算表，不會要求存取你全部的試算表）
 *
 * 使用方式請看 repo 裡的〈給朋友的設定步驟.md〉。
 * - 通行碼存在「專案設定 → 指令碼屬性」，屬性名稱為 TOKEN，不要寫在程式碼裡。
 * - 部署為網頁應用程式：執行身分「我」，存取權「任何人」。
 */

var SHEET_RECORDS = '紀錄';
var SHEET_SETTINGS = '設定';
var HEADERS = ['id', 'date', 'type', 'amount', 'note', 'ts', 'fx'];
var TYPES = ['salary', 'personal', 'fixpay', 'q_prod', 'q_trans', 'prod', 'trans',
  'ret_prod', 'ret_trans', 'rb_prod', 'rb_trans',
  'quota', 'advance', 'return', 'reimb'];

/* ---------- 對外入口 ---------- */

function doGet(e) {
  return respond_(function () {
    checkToken_(e && e.parameter ? e.parameter.token : '');
    return readAll_();
  });
}

function doPost(e) {
  return respond_(function () {
    var body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      throw new AppError_('bad_request', '送來的資料格式不對');
    }
    checkToken_(body.token);

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'add') return addEntries_(body.entries || (body.entry ? [body.entry] : []));
      if (body.action === 'delete') return deleteEntry_(body.id);
      if (body.action === 'saveSettings') return saveSettings_(body.settings || {});
      throw new AppError_('bad_request', '不認得的動作：' + body.action);
    } finally {
      lock.releaseLock();
    }
  });
}

/* 在編輯器裡手動執行一次，可以先建立好兩個工作表（不執行也沒關係，第一次使用時會自動建立） */
function setup() {
  recordsSheet_();
  settingsSheet_();
}

/* ---------- 動作 ---------- */

function readAll_() {
  var sh = recordsSheet_();
  var entries = [];
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var item = {
        id: String(r[0]),
        date: toDateStr_(r[1]),
        type: String(r[2]),
        amount: toNumber_(r[3]),
        note: String(r[4] == null ? '' : r[4]),
        ts: toNumber_(r[5])
      };
      if (r[6]) item.fx = String(r[6]); // 浮動固定支出的實際金額：對應哪一項
      entries.push(item);
    }
  }
  var s = readSettings_();
  var fixed = [];
  try {
    fixed = JSON.parse(s.fixed || '[]');
  } catch (err) {
    fixed = [];
  }
  var out = {
    v: 2, // 有這個欄位的版本：設定可以只存一部分、會回傳每日基準和通知狀態
    entries: entries,
    opening: toNumber_(s.opening),
    until: s.until ? toDateStr_(s.until) : '',
    fixed: Array.isArray(fixed) ? fixed : [],
    pay: toNumber_(s.pay),
    payday: toNumber_(s.payday),
    pct: toNumber_(s.pct)
  };
  if (s.daily != null && s.daily !== '') out.daily = toNumber_(s.daily);
  out.notes = parseObj_(s.notes);
  out.prefs = parseObj_(s.prefs);
  return out;
}

function addEntries_(list) {
  if (!Array.isArray(list)) throw new AppError_('bad_request', '紀錄格式不對');
  var sh = recordsSheet_();
  var existing = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) existing[String(ids[i][0])] = true;
  }
  var rows = [];
  for (var j = 0; j < list.length; j++) {
    var e = validEntry_(list[j]);
    if (existing[e.id]) continue; // 重送時不會重複新增
    existing[e.id] = true;
    rows.push([e.id, e.date, e.type, e.amount, safeText_(e.note), e.ts, safeText_(e.fx)]);
  }
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return { added: rows.length };
}

function deleteEntry_(id) {
  if (!id) throw new AppError_('bad_request', '缺少要刪除的紀錄 id');
  var sh = recordsSheet_();
  var last = sh.getLastRow();
  var removed = 0;
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(id)) {
        sh.deleteRow(i + 2);
        removed++;
      }
    }
  }
  return { removed: removed }; // 已經刪掉的再刪一次也算成功
}

/* 只存有送來的欄位：通知狀態、偏好這類常常變的東西單獨送，不會蓋掉別台裝置剛改的固定支出 */
function saveSettings_(s) {
  var vals = {};
  if ('opening' in s) vals.opening = toNumber_(s.opening);
  if ('until' in s) vals.until = typeof s.until === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s.until) ? s.until : '';
  if ('fixed' in s) vals.fixed = JSON.stringify(Array.isArray(s.fixed) ? s.fixed : []);
  if ('pay' in s) vals.pay = toNumber_(s.pay) > 0 ? toNumber_(s.pay) : 0;
  if ('payday' in s) {
    var payday = Math.round(toNumber_(s.payday));
    vals.payday = payday >= 1 && payday <= 31 ? payday : 10;
  }
  if ('pct' in s) {
    var pct = Math.round(toNumber_(s.pct));
    vals.pct = pct >= 10 && pct <= 60 ? pct : 35;
  }
  if ('daily' in s) {
    var daily = Math.round(toNumber_(s.daily));
    vals.daily = daily >= 50 && daily <= 20000 ? daily : 500;
  }
  if ('notes' in s) vals.notes = JSON.stringify(s.notes && typeof s.notes === 'object' ? s.notes : {});
  if ('prefs' in s) vals.prefs = JSON.stringify(s.prefs && typeof s.prefs === 'object' ? s.prefs : {});
  writeSettings_(vals);
  return { saved: true };
}

/* ---------- 工作表 ---------- */

function spreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new AppError_('no_sheet', '這個指令碼沒有綁定試算表，請從試算表的「擴充功能 → Apps Script」建立');
  return ss;
}

function recordsSheet_() {
  var ss = spreadsheet_();
  var sh = ss.getSheetByName(SHEET_RECORDS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RECORDS);
    // id、date、type、note 設成純文字，避免日期被自動轉換；amount、ts 是數字
    sh.getRange('A:C').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  // 舊版建立的工作表少了後面的欄位（例如 fx），補上標題
  var head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(head[HEADERS.length - 1]) !== HEADERS[HEADERS.length - 1]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange('G:G').setNumberFormat('@');
  }
  return sh;
}

function settingsSheet_() {
  var ss = spreadsheet_();
  var sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange(1, 1, 7, 2).setValues([
      ['name', 'value'],
      ['opening', '0'],
      ['until', ''],
      ['fixed', '[]'],
      ['pay', '0'],
      ['payday', '10'],
      ['pct', '35']
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* 照名稱找到那一列再寫；舊版建立的工作表沒有的欄位（例如 daily、notes）加在最後 */
function writeSettings_(vals) {
  var sh = settingsSheet_();
  var last = sh.getLastRow();
  var names = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var row = {};
  for (var i = 0; i < names.length; i++) if (names[i][0]) row[String(names[i][0])] = i + 2;
  for (var k in vals) {
    if (row[k]) sh.getRange(row[k], 2, 1, 1).setValues([[vals[k]]]);
    else {
      last = Math.max(last, 1) + 1;
      sh.getRange(last, 1, 1, 2).setValues([[k, vals[k]]]);
      row[k] = last;
    }
  }
}

function readSettings_() {
  var sh = settingsSheet_();
  var out = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0]) out[String(rows[i][0])] = rows[i][1];
    }
  }
  return out;
}

/* ---------- 工具 ---------- */

function parseObj_(v) {
  if (!v) return {};
  try {
    var o = JSON.parse(v);
    return o && typeof o === 'object' && !Array.isArray(o) ? o : {};
  } catch (err) {
    return {};
  }
}

function AppError_(code, message) {
  this.code = code;
  this.message = message;
}

function respond_(fn) {
  var out;
  try {
    out = fn();
    out.ok = true;
  } catch (err) {
    out = err instanceof AppError_
      ? { ok: false, error: err.code, message: err.message }
      : { ok: false, error: 'server', message: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!expected) throw new AppError_('no_token', '還沒設定通行碼：請到「專案設定 → 指令碼屬性」新增 TOKEN');
  if (String(token || '') !== String(expected)) throw new AppError_('bad_token', '通行碼不對');
}

function validEntry_(e) {
  if (!e || typeof e !== 'object') throw new AppError_('bad_request', '紀錄格式不對');
  var id = String(e.id || '');
  var amount = Number(e.amount);
  var date = String(e.date || '');
  if (!id) throw new AppError_('bad_request', '紀錄缺少 id');
  if (TYPES.indexOf(e.type) < 0) throw new AppError_('bad_request', '不認得的類型：' + e.type);
  if (!(amount > 0)) throw new AppError_('bad_request', '金額要大於 0');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new AppError_('bad_request', '日期格式要是 YYYY-MM-DD');
  if (e.type === 'fixpay' && !e.fx) throw new AppError_('bad_request', '固定支出的實際金額要指定是哪一項');
  return {
    id: id,
    date: date,
    type: e.type,
    amount: amount,
    note: String(e.note || '').slice(0, 200),
    ts: Number(e.ts) || Date.now(),
    fx: String(e.fx || '').slice(0, 40)
  };
}

/* 備註開頭是 = + - @ 時，前面加單引號，避免被試算表當成公式 */
function safeText_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function toNumber_(v) {
  var n = Number(String(v == null ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}

function toDateStr_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v || '');
}
