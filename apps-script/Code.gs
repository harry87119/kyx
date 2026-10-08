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
var SHEET_GUIDE = '說明';
// 最後兩欄：tag＝標籤（早餐、午餐…），類型＝type 的中文名稱（給人和 Claude 看的，網頁不讀）
var HEADERS = ['id', 'date', 'type', 'amount', 'note', 'ts', 'fx', 'tag', '類型'];
var TYPE_NAMES = {
  salary: '真實薪水', personal: '私人花費', fixpay: '固定支出・實際金額',
  q_prod: '收到製作費', q_trans: '收到交通費', prod: '用製作費', trans: '用交通費',
  ret_prod: '還公司・製作費', ret_trans: '還公司・交通費', rb_prod: '公司還我・製作費', rb_trans: '公司還我・交通費',
  quota: '收到製作費（舊版）', advance: '先墊・製作費（舊版）', 'return': '還公司・製作費（舊版）', reimb: '公司還我・製作費（舊版）'
};
function typeName_(t) { return TYPE_NAMES[t] || t; }
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
  guideSheet_();
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
      if (r[7]) item.tag = String(r[7]); // 標籤
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
  guideSheet_();
  var out = {
    v: 3, // 2：設定可以只存一部分、回傳每日基準和通知狀態；3：紀錄有標籤欄
    entries: entries,
    opening: toNumber_(s.opening),
    until: s.until ? toDateStr_(s.until) : '',
    fixed: Array.isArray(fixed) ? fixed : [],
    pay: toNumber_(s.pay),
    payday: toNumber_(s.payday),
    pct: toNumber_(s.pct)
  };
  if (s.daily != null && s.daily !== '') out.daily = toNumber_(s.daily);
  try { out.sheetUrl = spreadsheet_().getUrl(); } catch (err) {} // 網頁上「打開我的試算表」用
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
    rows.push([e.id, e.date, e.type, e.amount, safeText_(e.note), e.ts, safeText_(e.fx), safeText_(e.tag), typeName_(e.type)]);
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
    sh.getRange('G:I').setNumberFormat('@');
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  // 舊版建立的工作表少了後面的欄位（例如 fx、tag、類型），補上標題；舊紀錄的「類型」也一次補上中文名稱
  var head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(head[HEADERS.length - 1]) !== HEADERS[HEADERS.length - 1]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange('G:I').setNumberFormat('@');
    var last = sh.getLastRow();
    if (last >= 2) {
      var types = sh.getRange(2, 3, last - 1, 1).getValues();
      sh.getRange(2, HEADERS.length, last - 1, 1).setValues(types.map(function (t) { return [t[0] ? typeName_(String(t[0])) : '']; }));
    }
  }
  return sh;
}

/* 「說明」工作表：給人和 Claude（Google 試算表裡的 Claude 側欄）看的，寫清楚每一欄和計算規則。
   版本號（A1）不一樣才重寫，平常不會動 */
var GUIDE_VERSION = 'KYX記帳・試算表說明（v1）';
function guideSheet_() {
  var ss = spreadsheet_();
  var sh = ss.getSheetByName(SHEET_GUIDE);
  if (sh && String(sh.getRange(1, 1, 1, 1).getValues()[0][0]) === GUIDE_VERSION) return sh;
  if (!sh) sh = ss.insertSheet(SHEET_GUIDE);
  var rows = [
    [GUIDE_VERSION, ''],
    ['這份試算表是「KYX記帳」網頁的資料庫，網頁每次同步都會讀寫「紀錄」和「設定」這兩頁。', ''],
    ['重要：請不要修改、排序或刪除「紀錄」和「設定」這兩頁', '格式改錯，網頁會算錯或讀不到。記帳、改帳、刪帳請用網頁。要統計、畫圖、做報表，請另外開新的工作表（例如用公式讀「紀錄」）。'],
    ['', ''],
    ['「紀錄」每一欄', ''],
    ['id', '每一筆的編號（網頁產生，不能重複、不能改）'],
    ['date', '日期 YYYY-MM-DD'],
    ['type', '類型代號（見下面），網頁看這一欄'],
    ['amount', '金額（新台幣，一律是正數；收入或支出看類型）'],
    ['note', '備註'],
    ['ts', '記帳當下的時間（毫秒），同一天裡排先後用'],
    ['fx', '只有「固定支出・實際金額」才有：是哪一項浮動固定支出（對應「設定」fixed 裡的 sid 或 id）'],
    ['tag', '標籤：早餐、午餐、晚餐、飲料、點心、日用品、交通、娛樂、旅遊、美容美妝、服飾（只有支出才有，可以空白）'],
    ['類型', 'type 的中文名稱，只是方便看；網頁不讀這一欄'],
    ['', ''],
    ['類型代號', ''],
    ['salary', '真實薪水（收入）。金額至少是前一筆常規薪水一半的才算「常規薪水」，更小的是小筆收入'],
    ['personal', '私人花費（支出）'],
    ['fixpay', '固定支出・實際金額（支出）：浮動的固定支出（水電瓦斯等）繳了以後記的實際金額，會取代那幾期的預估'],
    ['q_prod / q_trans', '收到製作費／收到交通費：公司先撥的工作用錢，不是自己的錢'],
    ['prod / trans', '用製作費／用交通費：工作上的花費，先從公司撥的錢扣'],
    ['ret_prod / ret_trans', '還公司：把公司撥的、沒用完的錢還回去'],
    ['rb_prod / rb_trans', '公司還我：自己先墊的工作花費，公司還回來'],
    ['quota / advance / return / reimb', '舊版的類型，一律當成製作費那一組'],
    ['', ''],
    ['計算規則（網頁怎麼算）', ''],
    ['自己的錢', '薪水 − 私人花費 − 固定支出實際金額 − 已到期的定額固定支出 − 公司欠我的（自己先墊的工作花費）。公司撥的錢沒用完的部分是「我欠公司」，不算自己的錢'],
    ['定額固定支出', '每期金額一樣（房租等），到了繳費日自動扣，不用記'],
    ['浮動固定支出', '每期金額不一樣（水電瓦斯等），填預估。不會自動扣：要記了實際金額（fixpay）才扣；沒記之前先預留，到下次發薪都沒記，那期就不再預留'],
    ['每日基準', '每天預計花多少（設定 daily，預設 500）。「錢可以撐到哪一天」照這個金額一天一天扣'],
    ['下次發薪日', '設定 until；空白時照每月發薪日（payday）推算'],
    ['', ''],
    ['「設定」每一列', ''],
    ['opening', '起始結餘（固定 0）'],
    ['until', '下次發薪日（手動填的，空白＝照每月發薪日推算）'],
    ['pay', '手動改過的上次薪水（0＝自動抓最近一筆常規薪水）'],
    ['payday', '每月發薪日（1～31）'],
    ['pct', '生活費比例（%），用來算建議基準'],
    ['daily', '每日基準（每天預計花多少）'],
    ['fixed', '每月固定支出（JSON）：name 名稱、amount 金額、day 每月幾號、since 從哪個月開始、end 到哪個月、vary 浮動、skip 跳過的那幾期'],
    ['notes / prefs', '網頁的通知狀態和偏好（深淺色等），不用理它']
  ];
  while (rows.length < 60) rows.push(['', '']);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  try { sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 720); } catch (err) {}
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
    fx: String(e.fx || '').slice(0, 40),
    tag: String(e.tag || '').slice(0, 20)
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
