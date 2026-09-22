// 社内タイミー - 01_Setup.gs
// setup() / resetAndSetup() / setupTriggers()
// ※ GASエディタから手動実行してください


function setup() {
  Logger.log('=== setup() 開始 ===');
  const ss = getSpreadsheet();
  Logger.log('SS: ' + ss.getId());

  _createSheet(ss, SHEET.STAFF, [
    'line_user_id', 'display_name', 'name', 'phone', 'tax_type',
    'bank_info', 'same_day_cancel', 'total_cancel', 'caution_flag',
    'eval_avg', 'eval_level', 'eval_count', 'memo',
    'registered_at', 'updated_at', 'job_id',
    'zip_code', 'address', 'birth_date', 'furigana', 'my_number'
  ]);
  _createSheet(ss, SHEET.APPLIED, [
    'timestamp', 'dedupe_key', 'event', 'user_id',
    'date', 'job_id', 'source', 'note'
  ]);
  _createSheet(ss, SHEET.WORKLOG, [
    'timestamp', 'log_id', 'user_id', 'date', 'job_id',
    'action', 'start_time', 'end_time', 'actual_hours',
    'wage', 'transport', 'withholding', 'net_pay',
    'receipt_url', 'doc_log_id', 'note', 'corrects'
  ]);
  _createSheet(ss, SHEET.JOB, [
    'job_id', 'title', 'start_time', 'end_time',
    'wage', 'transport', 'active', 'note', 'created_at'
  ]);
  _createSheet(ss, SHEET.PAY_RECORD, [
    'pay_id', 'user_id', 'date', 'job_id',
    'hourly_wage', 'actual_hours', 'wage', 'transport', 'withholding', 'net_pay',
    'paid', 'paid_at', 'payment_method', 'receipt_url', 'note', 'created_at'
  ]);
  _createSheet(ss, SHEET.PAY_SUMM, [
    'user_id', 'year', 'total_wage', 'total_transport',
    'total_withholding', 'total_net_pay', 'work_days',
    'needs_report', 'report_submitted', 'updated_at'
  ]);
  _createSheet(ss, SHEET.CANCEL, [
    'timestamp', 'user_id', 'date', 'job_id',
    'cancel_type', 'reason', 'notified_at'
  ]);
  _createSheet(ss, SHEET.EVAL, [
    'timestamp', 'user_id', 'work_date',
    'score_punctuality', 'score_attitude', 'score_instruction',
    'score_accuracy', 'score_speed', 'score_report',
    'score_teamwork', 'score_rehire', 'score_avg',
    'level', 'strong_point', 'caution', 'next_policy'
  ]);
  _createSheet(ss, SHEET.DOC, [
    'doc_id', 'user_id', 'doc_type', 'target_period',
    'file_url', 'created_at', 'created_by'
  ]);
  _createSheet(ss, SHEET.KNOWLEDGE, [
    'id', 'category', 'question', 'answer', 'keywords', 'updated_at'
  ]);
  _createSheet(ss, SHEET.CHAT_LOG, [
    'timestamp', 'user_id', 'direction', 'message'
  ]);
  _createSheet(ss, SHEET.TAX, [
    'wage_from', 'wage_to', 'withholding'
  ]);

  _setupJobMaster(ss);
  _setupKnowledge(ss);
  _setupTaxTable(ss);
  _setupMessageConfig(ss);
  _setupDriveFolders();
  _formatStaffTextColumns_();
  setupTriggers();

  Logger.log('=== setup() 完了 ===');
  return 'setup()完了';
}

function _createSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    Logger.log('シート作成: ' + name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  // 既存シートに不足列をバッチ追加
  const cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const missing = headers.filter(h => !cur.includes(h));
  if (missing.length > 0) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
    Logger.log('列追加: ' + name + ' → ' + missing.join(', '));
  }
}

// JobMaster 初期データ（時給・職種名は実運用に合わせること）
function _setupJobMaster(ss) {
  const sh = ss.getSheetByName(SHEET.JOB);
  const data = sheetToObjects(sh);
  const ids = data.map(r => r.job_id);
  const now = formatDateTime(getJstNow());
  const jobs = [
    { job_id: 'JOB001', title: '温泉施設清掃', start_time: '10:00', end_time: '13:00', wage: 1050, transport: 150, active: 'true', note: '', created_at: now },
  ];
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  jobs.forEach(job => {
    if (ids.includes(job.job_id)) return;
    sh.appendRow(objToRow(hdr, job));
    Logger.log('JobMaster挿入: ' + job.job_id);
  });
}

function _setupKnowledge(ss) {
  const sh = ss.getSheetByName(SHEET.KNOWLEDGE);
  if (sh.getLastRow() > 1) return;
  const now = formatDateTime(getJstNow());
  const items = [
    [1, '勤務時間', '何時から何時まで働けますか？', '勤務時間は10:00〜13:00です。', '時間,勤務', now],
    [2, '日当', '日当はいくらですか？', '時給1,050円×3時間＝3,150円＋交通費150円です。', '給料,日当,金額', now],
    [3, 'キャンセル', '急に行けなくなった場合は？', 'できるだけ早めにご連絡ください。当日キャンセルはご遠慮ください。', 'キャンセル,欠席', now],
    [4, '持ち物', '当日の持ち物は？', '動きやすい服装でお越しください。必要な道具はこちらで用意します。', '持ち物,準備', now],
    [5, '服装', '服装に決まりはありますか？', '清潔感のある動きやすい服装をお願いします。', '服装,格好', now],
  ];
  sh.getRange(2, 1, items.length, items[0].length).setValues(items);
  Logger.log('Knowledge初期データ挿入');
}

function _setupTaxTable(ss) {
  const sh = ss.getSheetByName(SHEET.TAX);
  if (sh.getLastRow() > 1) return;
  const table = [
    [0,     9299,  0  ],
    [9300,  9999,  200],
    [10000, 10999, 300],
    [11000, 11999, 500],
    [12000, 12999, 610],
    [13000, 13999, 720],
  ];
  sh.getRange(2, 1, table.length, 3).setValues(table);
  Logger.log('TaxTable初期データ挿入');
}

function _setupMessageConfig(ss) {
  _createSheet(ss, SHEET.MSG_CONFIG, ['key', 'value', 'note']);
  const sh = ss.getSheetByName(SHEET.MSG_CONFIG);
  const existing = sheetToObjects(sh).map(r => r.key);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const items = [
    { key: 'MSG_TEMPLATE',  value: DEFAULT_MSG_TEMPLATE, note: '募集メッセージテンプレート（{{dates_list}}等を使用）' },
    { key: 'BREAK_INFO',    value: '休憩：11時頃・12時頃（各5分）', note: '休憩案内' },
    { key: 'PHONE_RULE',    value: '就業中のスマホ使用は禁止です', note: 'スマホルール' },
    { key: 'DEADLINE_NOTE', value: 'ご連絡期限までにご返信ください', note: '期限注意' },
  ];
  items.forEach(item => {
    if (!existing.includes(item.key)) {
      sh.appendRow(objToRow(hdr, item));
      Logger.log('MessageConfig挿入: ' + item.key);
    }
  });
}

// MSG_TEMPLATEをDEFAULT_MSG_TEMPLATEで上書きする（GASエディタから1回実行）
function resetMsgTemplate() {
  const sh = getSheet(SHEET.MSG_CONFIG);
  const data = sheetToObjects(sh);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const valIdx = hdr.indexOf('value');
  for (let i = 0; i < data.length; i++) {
    if (data[i].key === 'MSG_TEMPLATE') {
      sh.getRange(i + 2, valIdx + 1).setValue(DEFAULT_MSG_TEMPLATE);
      Logger.log('MSG_TEMPLATE 更新完了');
      return;
    }
  }
  Logger.log('MSG_TEMPLATE が見つかりません');
}

function _setupDriveFolders() {
  const rootName = '社内タイミー';
  const subFolders = ['賃金台帳', '領収書', '出勤簿', '支払調書'];
  const roots = DriveApp.getRootFolder().getFoldersByName(rootName);
  const rootFolder = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
  subFolders.forEach(name => {
    const f = rootFolder.getFoldersByName(name);
    if (!f.hasNext()) {
      rootFolder.createFolder(name);
      Logger.log('Driveフォルダ作成: ' + name);
    }
  });
}

function setupTriggers() {
  const managed = [
    'archiveLastMonth',
    'sendWeeklyRecruit',
    'sendNextDayReminder',
    'sendWeeklyApplyRequest_Tauchi',
    'checkAndGenerateMonthlyPayslip',
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (managed.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  // 毎週月曜17時：全スタッフ支払完了チェック→賃金台帳PDF自動生成
  // GASトリガーに「第n月曜日」の直接指定がないため毎週月曜に発火させ、
  // 関数内（60_Docs.gs checkAndGenerateMonthlyPayslip 先頭）で実行日が1〜7日の
  // 範囲内（＝その月の最初の月曜）かをガードして「第一月曜日のみ実処理」を実現する
  ScriptApp.newTrigger('checkAndGenerateMonthlyPayslip')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(17).create();
  // 毎月1日 3時：アーカイブ
  ScriptApp.newTrigger('archiveLastMonth')
    .timeBased().onMonthDay(1).atHour(3).create();
  // 毎週月曜 8時：週次募集送信
  ScriptApp.newTrigger('sendWeeklyRecruit')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  // 毎日 18時：前日リマインド
  ScriptApp.newTrigger('sendNextDayReminder')
    .timeBased().everyDays(1).atHour(18).create();
  // 毎週月曜 19時：田内あきこへ翌週火・土の応募依頼
  ScriptApp.newTrigger('sendWeeklyApplyRequest_Tauchi')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(19).create();
  Logger.log('トリガー設定完了');
}

// Staff・JobMaster・ChatLog を保持したまま運用シートをリセット
// GASエディタから手動実行
function partialReset() {
  // 削除前にバックアップを強制（誤爆時のリカバリ用）
  const backupUrl = backupAllData();
  if (!backupUrl) throw new Error('partialReset: バックアップURLが取得できませんでした。削除を中止します。');
  Logger.log('バックアップ完了（partialReset前）: ' + backupUrl);

  const ss = getSpreadsheet();
  const reset = [
    SHEET.APPLIED, SHEET.WORKLOG, SHEET.PAY_RECORD, SHEET.PAY_SUMM,
    SHEET.CANCEL, SHEET.EVAL, SHEET.DOC,
  ];
  reset.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      ss.deleteSheet(sh);
      Logger.log('削除: ' + name);
    }
  });
  setup();
  Logger.log('=== partialReset 完了（Staff・JobMaster・ChatLog 保持）===');
}

// 全シートを別スプレッドシートにコピーして保存する（resetAndSetup誤爆時のリカバリ用）
// 失敗時は例外を投げる（呼び出し元でガードすること）
function backupAllData() {
  const ss = getSpreadsheet();
  const name = '社内タイミー_バックアップ_' + formatDate(getJstNow());
  let copy;
  try {
    copy = ss.copy(name);
  } catch (e) {
    const msg = 'backupAllData 失敗: ' + e.message;
    Logger.log(msg);
    throw new Error(msg);
  }
  const url = copy.getUrl();
  if (!url) throw new Error('backupAllData: コピー成功したが URL が取得できませんでした');
  Logger.log('バックアップ完了: ' + url);
  return url;
}

function resetAndSetup() {
  // 全削除前にバックアップを強制（誤爆時のリカバリ用）
  // backupAllData() が失敗した場合は例外を投げてここで停止する
  const backupUrl = backupAllData();
  if (!backupUrl) throw new Error('resetAndSetup: バックアップURLが取得できませんでした。削除を中止します。');
  Logger.log('バックアップ完了（resetAndSetup前）: ' + backupUrl);

  const ss = getSpreadsheet();
  const tmp = ss.getSheetByName('__tmp__') || ss.insertSheet('__tmp__');
  ss.getSheets()
    .filter(s => s.getName() !== '__tmp__')
    .forEach(s => ss.deleteSheet(s));
  setup();
  const t = ss.getSheetByName('__tmp__');
  if (t) ss.deleteSheet(t);
  Logger.log('リセット&セットアップ完了');
}

// 全シートを作り直し（GASエディタから1回実行）
function fullReset() {
  resetAndSetup();
  resetMsgTemplate();
  Logger.log('=== fullReset 完了 ===');
}

// ChatLog の in/out 両方から日付を抽出し Applied に補完（冪等・重複スキップ）
// in：ユーザー発言から M/D・M月D日 形式を抽出
// out：Bot返信「受け付けました」から ・M/D 形式を抽出
// 対象：今日〜30日以内
// GASエディタから手動実行 または 管理画面ボタンから実行
function backfillFromChatLog() {
  const chatSh = getSpreadsheet().getSheetByName(SHEET.CHAT_LOG);
  if (!chatSh || chatSh.getLastRow() < 2) { Logger.log('ChatLog: データなし'); return 0; }

  const hdr  = chatSh.getRange(1, 1, 1, chatSh.getLastColumn()).getValues()[0];
  const uIdx = hdr.indexOf('user_id');
  const dIdx = hdr.indexOf('direction');
  const mIdx = hdr.indexOf('message');
  const rows = chatSh.getRange(2, 1, chatSh.getLastRow() - 1, chatSh.getLastColumn()).getValues();

  const jobId  = 'JOB001';
  const now    = getJstNow();
  const today  = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const limit  = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);
  let added    = 0;

  function tryAdd(userId, month, day) {
    if (month < 1 || month > 12 || day < 1 || day > 31) return;
    const dateObj = new Date(today.getFullYear(), month - 1, day);
    if (dateObj < today || dateObj > limit) return;
    const dateStr = dateObj.getFullYear() + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    const ok = appendApplied('applied', userId, dateStr, jobId, 'chatlog', 'ChatLogから自動補完');
    if (ok) { added++; Logger.log('追記: ' + userId + ' ' + dateStr); }
    else     { Logger.log('スキップ: ' + userId + ' ' + dateStr); }
  }

  rows.forEach(function(row) {
    const dir    = String(row[dIdx] || '');
    const msg    = String(row[mIdx] || '');
    const userId = String(row[uIdx] || '');
    if (!userId) return;

    if (dir === 'out' && msg.includes('受け付けました')) {
      // Bot返信：「・M/D」形式
      var matches = msg.match(/・(\d{1,2})\/(\d{1,2})/g) || [];
      matches.forEach(function(dm) {
        var p = dm.replace('・', '').split('/');
        tryAdd(userId, parseInt(p[0], 10), parseInt(p[1], 10));
      });
    } else if (dir === 'in') {
      // ユーザー発言：M/D と M月D日 を抽出
      var s = msg;
      s = s.replace(/[^\n。]*(?:できませんか|ますか|でしょうか)[^\n。]*/g, ' ');
      var re1 = /(\d{1,2})\/(\d{1,2})/g, m1;
      while ((m1 = re1.exec(s)) !== null) tryAdd(userId, parseInt(m1[1], 10), parseInt(m1[2], 10));
      var re2 = /(\d{1,2})月(\d{1,2})日/g, m2;
      while ((m2 = re2.exec(s)) !== null) tryAdd(userId, parseInt(m2[1], 10), parseInt(m2[2], 10));
    }
  });

  Logger.log('=== backfillFromChatLog 完了: ' + added + '件追記 ===');
  return added;
}

// phone / zip_code / my_number 列をプレーンテキスト型に設定（数値・日付解釈を防ぐ）
// setup() から自動呼び出し。既存シートへの適用は GASエディタから直接実行してもよい
function _formatStaffTextColumns_() {
  try {
    const sh = getSheet(SHEET.STAFF);
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    ['phone', 'zip_code', 'my_number'].forEach(col => {
      const idx = hdr.indexOf(col);
      if (idx < 0) return;
      sh.getRange(2, idx + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
    Logger.log('Staff列フォーマット設定完了: phone / zip_code / my_number → テキスト型');
  } catch(e) {
    Logger.log('_formatStaffTextColumns_ スキップ: ' + e.message);
  }
}

// 既存 Staff シートにプロフィール列を末尾追加（1回限り実行）
function addStaffProfileColumns() {
  const sh = getSheet(SHEET.STAFF);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const toAdd = ['zip_code', 'address', 'birth_date', 'furigana', 'my_number']
    .filter(c => hdr.indexOf(c) < 0);
  if (!toAdd.length) { Logger.log('列追加不要（既に存在）'); return; }
  const startCol = sh.getLastColumn() + 1;
  sh.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  Logger.log('追加完了: ' + toAdd.join(', '));
}


// ─── 診断：6月以降の応募状態一覧（GASエディタから手動実行）───
// Applied シートの最新イベントをスタッフ名付きで一覧表示する
// 確認したい月を FROM_DATE で変更してください
function diagCheckAppliedStatus() {
  const FROM_DATE = '2026-06-01'; // 確認開始日（ここを変えて絞り込み）

  const staffMap = {};
  sheetToObjects(getSheet(SHEET.STAFF)).forEach(s => {
    staffMap[s.line_user_id] = s.name || s.display_name || s.line_user_id;
  });

  const allApplied = getAppliedAllWithArchives_();
  const latestMap = buildLatestAppliedMap_(allApplied);

  const rows = Object.values(latestMap)
    .filter(r => safeDate(r.date) >= FROM_DATE)
    .sort((a, b) => safeDate(a.date).localeCompare(safeDate(b.date)) || (a.user_id > b.user_id ? 1 : -1));

  Logger.log('=== 応募状態チェック（' + FROM_DATE + ' 以降） ===');
  Logger.log('合計: ' + rows.length + '件');
  Logger.log('');

  const byEvent = { applied: [], approved: [], cancelled: [], other: [] };
  rows.forEach(r => {
    const name = staffMap[r.user_id] || r.user_id;
    const line = safeDate(r.date) + '  [' + (r.event || '?') + ']  ' + name + '  (' + r.job_id + ')';
    Logger.log(line);
    const g = byEvent[r.event] || byEvent.other;
    g.push(name + '/' + safeDate(r.date));
  });

  Logger.log('');
  Logger.log('--- サマリ ---');
  Logger.log('承認待（applied）: ' + byEvent.applied.length + '件 → ' + byEvent.applied.join(', '));
  Logger.log('確定済（approved）: ' + byEvent.approved.length + '件 → ' + byEvent.approved.join(', '));
  Logger.log('キャンセル済（cancelled）: ' + byEvent.cancelled.length + '件 → ' + byEvent.cancelled.join(', '));
  Logger.log('=== 完了 ===');
}

// 田内あきこの line_user_id を Script Properties に保存する（1回だけ手動実行）
function setupTauchiUserId() {
  const tauchi = getAllStaff().find(function(s) { return s.name === '田内あきこ'; });
  if (!tauchi) {
    Logger.log('[setupTauchiUserId] 田内あきこ が Staff シートに見つかりません。名前を確認してください。');
    return;
  }
  if (!tauchi.line_user_id) {
    Logger.log('[setupTauchiUserId] 田内あきこ の line_user_id が空です。LINE連携を確認してください。');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('TAUCHI_LINE_USER_ID', tauchi.line_user_id);
  Logger.log('[setupTauchiUserId] 保存完了: ' + tauchi.line_user_id);
}
