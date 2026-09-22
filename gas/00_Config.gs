// 社内タイミー - 00_Config.gs
// 定数・SS接続・ユーティリティ関数
// 法的根拠：労働基準法第108条（賃金台帳）、所得税法第225条（支払調書）

const SECRETS = {
  get SPREADSHEET_ID()       { return _prop('SPREADSHEET_ID'); },
  get CHANNEL_ACCESS_TOKEN() { return _prop('CHANNEL_ACCESS_TOKEN'); },
  get CHANNEL_SECRET()       { return _safeProp('CHANNEL_SECRET', ''); },
  get ADMIN_USER_ID()        { return _safeProp('ADMIN_USER_ID', ''); },      // 未設定可
  get COMPANY_NAME()         { return _safeProp('COMPANY_NAME', '地涌の温泉'); },
  get COMPANY_ADDRESS()      { return _safeProp('COMPANY_ADDRESS', ''); },    // 後から設定可
  get COMPANY_TEL()          { return _safeProp('COMPANY_TEL', ''); },        // 後から設定可
};

let __SPREADSHEET_CACHE__ = null;

// 必須プロパティ（未設定はエラー）
function _prop(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('[Config] ' + key + ' が Script Properties に未設定');
  return v;
}

// オプションプロパティ（未設定はデフォルト値を返す）
function _safeProp(key, defaultValue) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v || defaultValue || '';
}

const SHEET = {
  STAFF:      'Staff',
  APPLIED:    'Applied',
  WORKLOG:    'WorkLog',
  JOB:        'JobMaster',
  PAY_RECORD: 'PayRecord',
  PAY_SUMM:   'PaySummary',
  CANCEL:     'CancelLog',
  EVAL:       'Evaluation',
  DOC:        'DocLog',
  KNOWLEDGE:  'Knowledge',
  TAX:        'TaxTable',
  MSG_CONFIG: 'MessageConfig',
  CHAT_LOG:   'ChatLog',
};

// 応募フォーム（静的ページ・GitHub Pages）
// GAS HtmlServiceのiframe配信は実機で画面幅が正しく取れない問題があるため、
// UI自体は外部ホスティングに置き、doGet(page=apply_data)とdoPost(action=apply)を
// JSON APIとして呼び出す構成にした（2026-08-22, implementation-notes参照）
const STATIC_APPLY_URL = 'https://kiyo7777.github.io/shanai-timee-apply/';

// 募集メッセージのデフォルトテンプレート
// {{dates_list}} は1行1日付（曜日あり・例：8月25日（火））形式。そのままLINE返信としてコピーできるようにするため
const DEFAULT_MSG_TEMPLATE =
  '📣 お手伝いスタッフ募集\n\n' +
  '{{dates_list}}\n\n' +
  '⏰ 勤務時間：{{start_time}}〜{{end_time}}\n' +
  '💴 日当：{{daily_wage}}円　交通費：{{transport}}円（日払い）\n\n' +
  'フォームから希望日を選べます：\n{{apply_url}}\n\n' +
  'LINEで返信する場合は、上の日付をそのままコピーしてください。\n\n' +
  '初めての方はお名前も教えてください😊\n{{phone_rule}}';

function getSpreadsheet() {
  if (__SPREADSHEET_CACHE__) return __SPREADSHEET_CACHE__;
  try {
    __SPREADSHEET_CACHE__ = SpreadsheetApp.openById(SECRETS.SPREADSHEET_ID);
    return __SPREADSHEET_CACHE__;
  } catch(e) {
    throw new Error('[Config] SSを開けません: ' + e.message);
  }
}

function getSheet(name) {
  const sh = getSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('[Config] シート未存在: ' + name + ' → setup()を実行してください');
  return sh;
}

// ヘッダーマップ（列名→0始まりインデックス）
function buildHeaderMap(sheet) {
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  hdr.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

// シートの全データをオブジェクト配列で返す
function sheetToObjects(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(row => {
    const obj = {};
    hdr.forEach((h, i) => { if (h) obj[String(h).trim()] = row[i]; });
    return obj;
  });
}

// 日時ユーティリティ（全てJST基準）
// new Date() はUTCベース。formatDate/Time/DateTimeが内部で+9hしてUTCゲッターで読む方式。
function getJstNow() { return new Date(); }

function formatDate(d) {
  if (!(d instanceof Date)) return String(d || '').substring(0, 10);
  const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
  return jst.getUTCFullYear() + '-' +
    String(jst.getUTCMonth() + 1).padStart(2, '0') + '-' +
    String(jst.getUTCDate()).padStart(2, '0');
}

function formatTime(d) {
  if (!(d instanceof Date)) return String(d || '');
  const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
  return String(jst.getUTCHours()).padStart(2, '0') + ':' +
    String(jst.getUTCMinutes()).padStart(2, '0');
}

// スプレッドシートの時刻セルはDateオブジェクトで返ってくる場合があるため安全に変換
function _safeTime_(t) {
  if (!t) return '';
  if (t instanceof Date) return formatTime(t);
  return String(t);
}

function formatDateTime(d) {
  if (!(d instanceof Date)) return String(d || '');
  const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
  return jst.getUTCFullYear() + '-' +
    String(jst.getUTCMonth() + 1).padStart(2, '0') + '-' +
    String(jst.getUTCDate()).padStart(2, '0') + ' ' +
    String(jst.getUTCHours()).padStart(2, '0') + ':' +
    String(jst.getUTCMinutes()).padStart(2, '0') + ':' +
    String(jst.getUTCSeconds()).padStart(2, '0');
}

// どんな入力でも YYYY-MM-DD 文字列に変換（JST基準）
function safeDate(val) {
  if (!val) return '';
  if (val instanceof Date) return formatDate(val);
  const s = String(val);
  if (s.includes('T') || s.includes('Z')) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? s.substring(0, 10) : formatDate(d);
  }
  return s.substring(0, 10);
}

function toEpochMs_(val) {
  if (!val) return 0;
  if (val instanceof Date) return val.getTime();
  const s = String(val).trim();
  if (!s) return 0;
  const d = new Date(s.replace(' ', 'T'));
  const ms = d.getTime();
  return isNaN(ms) ? 0 : ms;
}

function isNewerRow_(row, existing) {
  if (!existing) return true;
  const a = toEpochMs_(row.timestamp);
  const b = toEpochMs_(existing.timestamp);
  if (a !== b) return a > b;
  return String(row.dedupe_key || row.log_id || '') > String(existing.dedupe_key || existing.log_id || '');
}

function isPaid_(val) {
  return val === true || String(val).toLowerCase() === 'true';
}

function formatDateJa(dateStr) {
  const DOW = ['日', '月', '火', '水', '木', '金', '土'];
  const s = safeDate(dateStr);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return String(dateStr || ''); // 不正日付はそのまま返す（NaN表示防止）
  const [y, m, d] = s.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  return m + '/' + d + '（' + DOW[dt.getDay()] + '）';
}

function formatRecruitDateLine(dateStr) {
  const DOW = ['日', '月', '火', '水', '木', '金', '土'];
  const s = safeDate(dateStr);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return String(dateStr || '');
  const [y, m, d] = s.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  return m + '月' + d + '日（' + DOW[dt.getDay()] + '）';
}

function formatRecruitReplyExample(dateStr) {
  const s = safeDate(dateStr);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return String(dateStr || '');
  const parts = s.split('-').map(Number);
  return parts[1] + '月' + parts[2] + '日';
}

function ensureRecruitMessageParts(msg, applyUrl, replyExamples) {
  let out = String(msg || '');
  if (applyUrl && out.indexOf(applyUrl) === -1) {
    out += '\n\nフォームから応募：\n' + applyUrl;
  }
  if (replyExamples && out.indexOf(replyExamples) === -1) {
    out += '\n\nLINE返信用（日付だけコピー）：\n' + replyExamples;
  }
  return out;
}

// UUID生成
function uuid() { return Utilities.getUuid(); }

// dedupe_keyチェック
function isDuplicate(sheetName, key) {
  const sh = getSheet(sheetName);
  if (sh.getLastRow() < 2) return false;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const ki = hdr.indexOf('dedupe_key');
  if (ki < 0) return false;
  const col = sh.getRange(2, ki + 1, sh.getLastRow() - 1, 1).getValues();
  return col.some(r => r[0] === key);
}

// 列順序通りに行データを配列化するヘルパー
function objToRow(headers, obj) {
  return headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
}

// 日当計算（時給 × 勤務時間 + 交通費）
// 例: calcDailyWage(1050, '10:00', '13:00', 150) → { hours:3, dailyWage:3150, transport:150, total:3300 }
function calcDailyWage(hourlyWage, startTime, endTime, transport) {
  const st = _safeTime_(startTime) || '00:00';
  const et = _safeTime_(endTime)   || '00:00';
  const [sh, sm] = st.split(':').map(Number);
  const [eh, em] = et.split(':').map(Number);
  const hours     = Math.max(0, (eh * 60 + em - sh * 60 - sm) / 60);
  const dailyWage = Math.round((Number(hourlyWage) || 0) * hours);
  const trans     = Number(transport) || 0;
  return { hours, dailyWage, transport: trans, total: dailyWage + trans };
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { safeDate, calcDailyWage, formatDate, formatTime, _safeTime_, toEpochMs_, isPaid_ };
}
