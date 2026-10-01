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
var HEADERS = ['id', 'date', 'type', 'amount', 'note', 'ts'];
var TYPES = ['salary', 'personal', 'q_prod', 'q_trans', 'prod', 'trans',
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
      entries.push({
        id: String(r[0]),
        date: toDateStr_(r[1]),
        type: String(r[2]),
        amount: toNumber_(r[3]),
        note: String(r[4] == null ? '' : r[4]),
        ts: toNumber_(r[5])
      });
    }
  }
  var s = readSettings_();
  var fixed = [];
  try {
    fixed = JSON.parse(s.fixed || '[]');
  } catch (err) {
    fixed = [];
  }
  return {
    entries: entries,
    opening: toNumber_(s.opening),
    until: s.until ? toDateStr_(s.until) : '',
    fixed: Array.isArray(fixed) ? fixed : [],
    daily: toNumber_(s.daily) || 600
  };
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
    rows.push([e.id, e.date, e.type, e.amount, safeText_(e.note), e.ts]);
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

function saveSettings_(s) {
  var opening = toNumber_(s.opening);
  var until = typeof s.until === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s.until) ? s.until : '';
  var fixed = Array.isArray(s.fixed) ? s.fixed : [];
  var daily = toNumber_(s.daily) > 0 ? toNumber_(s.daily) : 600;
  var sh = settingsSheet_();
  sh.getRange(2, 1, 4, 2).setValues([
    ['opening', opening],
    ['until', until],
    ['fixed', JSON.stringify(fixed)],
    ['daily', daily]
  ]);
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
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function settingsSheet_() {
  var ss = spreadsheet_();
  var sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange(1, 1, 5, 2).setValues([
      ['name', 'value'],
      ['opening', '0'],
      ['until', ''],
      ['fixed', '[]'],
      ['daily', '600']
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
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
  return {
    id: id,
    date: date,
    type: e.type,
    amount: amount,
    note: String(e.note || '').slice(0, 200),
    ts: Number(e.ts) || Date.now()
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
