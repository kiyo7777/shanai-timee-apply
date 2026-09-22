// 社内タイミー - AdminUi.gs
// doGet / doPost / 管理画面API


function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';

  // 応募フォームを外部静的ページ（GitHub Pages等）から使うためのJSON API。
  // GAS Web AppのHTML配信はGoogle側のsandbox iframeに包まれ、実機で画面幅が
  // 正しく取れない問題があるため、UI自体は外部ホスティングに置き、データ取得と
  // 応募送信（doPost の action=apply）だけこのWeb Appを叩く構成にする。
  if (page === 'apply_data') {
    const token = (e.parameter && e.parameter.token) || '';
    const td = _getTokenData(token);
    if (!td || td.type !== 'apply') {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, msg: 'トークン無効' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const tokenDates = _normalizeTokenDates_(td.dates);
    const tokenJobId = td.jobId ? String(td.jobId) : 'JOB001';
    const activeJobs = sheetToObjects(getSheet(SHEET.JOB))
      .filter(j => String(j.active).toLowerCase() === 'true')
      .map(_normalizeJob_);
    const filtered = activeJobs.filter(j => String(j.job_id) === tokenJobId);
    const jobs = filtered.length ? filtered : activeJobs;
    const staff = td.userId ? getStaff(td.userId) : null;
    const staffName = staff ? (staff.name || staff.display_name || '') : '';
    // 日付ごとの現在の状態（applied=応募中 / approved=確定）を返し、
    // フォーム側で応募済み日程を選択不可・状態表示できるようにする。
    // tokenDatesは常に未来日（募集日程）のためアーカイブは不要 → Appliedシートを
    // 1回だけ読んで全日付分をまとめて判定する（getShiftStatusを日付ごとに呼ぶと
    // 呼び出し回数分アーカイブ全読みが走り重くなるため避ける）
    const dateStatus = {};
    if (td.userId) {
      const latestMap = buildLatestAppliedMap_(
        getAppliedAll().filter(r => r.user_id === td.userId && (r.job_id || 'JOB001') === tokenJobId)
      );
      tokenDates.forEach(d => {
        const r = latestMap[td.userId + ':' + d + ':' + tokenJobId];
        if (r && (r.event === 'applied' || r.event === 'approved')) dateStatus[d] = r.event;
      });
    }
    return ContentService.createTextOutput(JSON.stringify({
      ok: true, jobs, dates: tokenDates, jobId: tokenJobId, staffName, dateStatus,
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'apply') {
    const token = (e.parameter && e.parameter.token) || '';
    const td = _getTokenData(token);
    const tokenDates = (td && td.type === 'apply') ? _normalizeTokenDates_(td.dates) : [];
    // 職種は管理側が指定する。トークンに職種指定が無い場合（スタッフ発信フォーム等）は
    // デフォルト JOB001（温泉施設清掃）に固定し、スタッフには職種を選ばせない。
    const tokenJobId = (td && td.type === 'apply' && td.jobId) ? String(td.jobId) : 'JOB001';
    const activeJobs = sheetToObjects(getSheet(SHEET.JOB))
      .filter(j => String(j.active).toLowerCase() === 'true')
      .map(_normalizeJob_);
    // 指定職種で絞り込み（該当なしの保険として activeJobs にフォールバック）
    const filtered = activeJobs.filter(j => String(j.job_id) === tokenJobId);
    const jobs = filtered.length ? filtered : activeJobs;
    const staff = (td && td.userId) ? getStaff(td.userId) : null;
    const staffName = staff ? (staff.name || staff.display_name || '') : '';
    const tpl = HtmlService.createTemplateFromFile('ApplyForm');
    tpl.token     = token;
    tpl.jobs      = JSON.stringify(jobs);
    tpl.dates     = JSON.stringify(tokenDates);
    tpl.jobId     = JSON.stringify(tokenJobId);
    tpl.staffName = staffName;
    return tpl.evaluate().setTitle('シフト応募').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'register') {
    const token = (e.parameter && e.parameter.token) || '';
    const tpl = HtmlService.createTemplateFromFile('RegisterForm');
    tpl.token  = token;
    tpl.liffId = _safeProp('LIFF_ID', '');
    return tpl.evaluate().setTitle('プロフィール登録・更新').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'checkin') {
    const token = (e.parameter && e.parameter.token) || '';
    const td = _getTokenData(token);
    const userId = (td && td.type === 'checkin') ? td.userId : '';
    const tpl = HtmlService.createTemplateFromFile('CheckIn');
    tpl.userId = userId;
    return tpl.evaluate().setTitle('チェックイン').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'store_checkin') {
    const staffList = sheetToObjects(getSheet(SHEET.STAFF))
      .filter(s => s.line_user_id)
      .map(s => ({ line_user_id: s.line_user_id, name: s.name || s.display_name || s.line_user_id }));
    const tpl = HtmlService.createTemplateFromFile('StoreCheckin');
    tpl.staffJson  = JSON.stringify(staffList);
    tpl.company    = _safeProp('COMPANY_NAME', '地涌の温泉');
    return tpl.evaluate().setTitle('チェックイン').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'qr_checkin') {
    const uid = (e.parameter && e.parameter.uid) || '';
    const tpl = HtmlService.createTemplateFromFile('QrCheckin');
    tpl.uid     = uid;
    tpl.company = _safeProp('COMPANY_NAME', '地涌の温泉');
    return tpl.evaluate().setTitle('チェックイン').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'scan_checkin') {
    const tpl = HtmlService.createTemplateFromFile('ScanCheckin');
    tpl.company = _safeProp('COMPANY_NAME', '地涌の温泉');
    return tpl.evaluate()
      .setTitle('スキャンチェックイン')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 管理画面
  const tpl = HtmlService.createTemplateFromFile('AdminPage');
  tpl.webAppUrl = ScriptApp.getService().getUrl();
  return tpl.evaluate()
    .setTitle('社内タイミー 管理画面')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================
// 管理画面 API（google.script.run経由）
// =============================

function api_getApplied() {
  const staffMap = {};
  sheetToObjects(getSheet(SHEET.STAFF)).forEach(s => {
    staffMap[s.line_user_id] = s;
  });

  // PayRecord を userId:date:jobId でインデックス（支払済みが最優先）
  const payMap = {};
  sheetToObjects(getSheet(SHEET.PAY_RECORD)).forEach(r => {
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    const rPaid = isPaid_(r.paid);
    const existing = payMap[key];
    const exPaid = existing && isPaid_(existing.paid);
    // 支払済みレコードを優先。未払いは支払済みで上書きされない
    if (!existing || (!exPaid && rPaid)) payMap[key] = r;
  });

  // サーバー側でlatestMap: userId:date:jobId ごとに最新イベント行のみ残す
  const allApplied = getAppliedAllWithArchives_();
  const latestMap = buildLatestAppliedMap_(allApplied);

  const rows = Object.values(latestMap).map(r => {
    const s = staffMap[r.user_id] || {};
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    const pr = payMap[key] || null;
    const paidBool = pr ? isPaid_(pr.paid) : null;
    return {
      ...r,
      name:         s.name || '',
      display_name: s.display_name || '',
      paid:         paidBool,
      paid_at:      pr ? pr.paid_at            : null,
      pay_id:       pr ? (pr.pay_id || '')      : null,
      net_pay:      pr ? (Number(pr.net_pay)    || 0) : null,
      wage:         pr ? (Number(pr.wage)       || 0) : null,
      transport:    pr ? (Number(pr.transport)  || 0) : null,
      withholding:  pr ? (Number(pr.withholding)|| 0) : null,
    };
  });
  return JSON.stringify(rows);
}

function api_getStaffList() {
  const staffList = getAllStaff();
  const allApplied = getAppliedAllWithArchives_();
  const paidSet = new Set(
    sheetToObjects(getSheet(SHEET.PAY_RECORD)).filter(r => isPaid_(r.paid)).map(r => r.user_id + '_' + safeDate(r.date))
  );
  const applyMap = {};
  allApplied.forEach(r => {
    if (!applyMap[r.user_id]) applyMap[r.user_id] = { applied: 0, approved: 0, last_date: '' };
    if (r.event === 'applied')  applyMap[r.user_id].applied++;
    if (r.event === 'approved') {
      applyMap[r.user_id].approved++;
      const d = safeDate(r.date);
      if (d > applyMap[r.user_id].last_date) applyMap[r.user_id].last_date = d;
    }
  });
  // キャンセル済み・未来シフトを除外し、最終状態が approved の過去分のみカウント
  const today_ = formatDate(getJstNow());
  const unpaidMap = {};
  Object.values(buildLatestAppliedMap_(allApplied)).forEach(r => {
    if (r.event === 'approved') {
      const d = safeDate(r.date);
      if (d > today_) return;
      if (!paidSet.has(r.user_id + '_' + d)) unpaidMap[r.user_id] = (unpaidMap[r.user_id] || 0) + 1;
    }
  });
  const result = staffList.map(s => ({
    ...s,
    apply_count:    (applyMap[s.line_user_id] || {}).applied   || 0,
    approved_count: (applyMap[s.line_user_id] || {}).approved  || 0,
    last_date:      (applyMap[s.line_user_id] || {}).last_date || '',
    unpaid_count:   unpaidMap[s.line_user_id] || 0,
  }));
  return JSON.stringify(result);
}

// 初期ロード一括取得（Applied・Staff・PayRecord を1回ずつだけ読む）
function api_loadAll() {
  // ── 1. 共通データ読み込み（各シート1回のみ）──
  const staffList  = getAllStaff();
  const allApplied = getAppliedAllWithArchives_();
  const allPay     = sheetToObjects(getSheet(SHEET.PAY_RECORD));
  const jobs       = sheetToObjects(getSheet(SHEET.JOB)).map(_normalizeJob_);

  // ── 2. applied: latestMap dedup + pay情報付与 ──
  const staffMap = {};
  staffList.forEach(s => { staffMap[s.line_user_id] = s; });

  const payMap = {};
  allPay.forEach(r => {
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    const rPaid = isPaid_(r.paid);
    const ex = payMap[key];
    const exPaid = ex && isPaid_(ex.paid);
    if (!ex || (!exPaid && rPaid)) payMap[key] = r;
  });

  const latestMap = buildLatestAppliedMap_(allApplied);

  const applied = Object.values(latestMap).map(r => {
    const s = staffMap[r.user_id] || {};
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    const pr = payMap[key] || null;
    const paidBool = pr ? isPaid_(pr.paid) : null;
    return {
      ...r,
      name:         s.name || '',
      display_name: s.display_name || '',
      paid:         paidBool,
      paid_at:      pr ? pr.paid_at             : null,
      pay_id:       pr ? (pr.pay_id || '')       : null,
      net_pay:      pr ? (Number(pr.net_pay)     || 0) : null,
      wage:         pr ? (Number(pr.wage)        || 0) : null,
      transport:    pr ? (Number(pr.transport)   || 0) : null,
      withholding:  pr ? (Number(pr.withholding) || 0) : null,
    };
  });

  // ── 3. staff: applyMap + unpaidMap 付与 ──
  const applyMap = {};
  allApplied.forEach(r => {
    if (!applyMap[r.user_id]) applyMap[r.user_id] = { applied: 0, approved: 0, last_date: '' };
    if (r.event === 'applied')  applyMap[r.user_id].applied++;
    if (r.event === 'approved') {
      applyMap[r.user_id].approved++;
      const d = safeDate(r.date);
      if (d > applyMap[r.user_id].last_date) applyMap[r.user_id].last_date = d;
    }
  });
  const paidSet = new Set(allPay.filter(r => isPaid_(r.paid)).map(r => r.user_id + '_' + safeDate(r.date)));
  const unpaidMap = {};
  // キャンセル済み・未来シフトを除外し、最終状態が approved の過去分のみカウント
  const today__ = formatDate(getJstNow());
  Object.values(buildLatestAppliedMap_(allApplied)).forEach(r => {
    if (r.event === 'approved') {
      const d = safeDate(r.date);
      if (d > today__) return;
      if (!paidSet.has(r.user_id + '_' + d)) unpaidMap[r.user_id] = (unpaidMap[r.user_id] || 0) + 1;
    }
  });
  const staff = staffList.map(s => ({
    ...s,
    apply_count:    (applyMap[s.line_user_id] || {}).applied   || 0,
    approved_count: (applyMap[s.line_user_id] || {}).approved  || 0,
    last_date:      (applyMap[s.line_user_id] || {}).last_date || '',
    unpaid_count:   unpaidMap[s.line_user_id] || 0,
  }));

  return JSON.stringify({ applied, staff, jobs });
}

// シフトのみ軽量取得（Staff読み込みをスキップ）
// 承認・割当・キャンセル後のリフレッシュ用
function api_loadShifts() {
  const allApplied = getAppliedAllWithArchives_();
  const allPay     = sheetToObjects(getSheet(SHEET.PAY_RECORD));
  const jobs       = sheetToObjects(getSheet(SHEET.JOB)).map(_normalizeJob_);

  const payMap = {};
  allPay.forEach(r => {
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    const rPaid = isPaid_(r.paid);
    const ex = payMap[key];
    const exPaid = ex && isPaid_(ex.paid);
    if (!ex || (!exPaid && rPaid)) payMap[key] = r;
  });

  const latestMap = buildLatestAppliedMap_(allApplied);

  // cancelled 行の直前イベントを取得（approved→cancelled か applied→cancelled かを区別）
  const keyRows = {};
  allApplied.forEach(r => {
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    if (!keyRows[key]) keyRows[key] = [];
    keyRows[key].push(r);
  });
  const prevEventMap = {};
  Object.keys(keyRows).forEach(key => {
    const sorted = keyRows[key].slice().sort((a, b) => String(a.timestamp || '').localeCompare(String(b.timestamp || '')));
    if (sorted.length >= 2) prevEventMap[key] = sorted[sorted.length - 2].event;
  });

  const applied = Object.values(latestMap).map(r => {
    const key = (r.user_id || '') + ':' + safeDate(r.date) + ':' + (r.job_id || '');
    const pr = payMap[key] || null;
    const paidBool = pr ? isPaid_(pr.paid) : null;
    return {
      ...r,
      prev_event:  prevEventMap[key] || null,
      paid:        paidBool,
      paid_at:     pr ? pr.paid_at             : null,
      pay_id:      pr ? (pr.pay_id || '')       : null,
      net_pay:     pr ? (Number(pr.net_pay)     || 0) : null,
      wage:        pr ? (Number(pr.wage)        || 0) : null,
      transport:   pr ? (Number(pr.transport)   || 0) : null,
      withholding: pr ? (Number(pr.withholding) || 0) : null,
    };
  });

  return JSON.stringify({ applied, jobs });
}

// スタッフ情報のみ軽量取得（Applied・JobMaster読み込みをスキップ）
// 支払い処理後のリフレッシュ用
function api_loadStaffSummary() {
  const staffList = getAllStaff();
  const allPay    = sheetToObjects(getSheet(SHEET.PAY_RECORD));
  const allApplied = getAppliedAllWithArchives_();
  const paidSet = new Set(allPay.filter(r => isPaid_(r.paid)).map(r => r.user_id + '_' + safeDate(r.date)));

  const applyMap = {};
  allApplied.forEach(r => {
    if (!applyMap[r.user_id]) applyMap[r.user_id] = { applied: 0, approved: 0, last_date: '' };
    if (r.event === 'applied')  applyMap[r.user_id].applied++;
    if (r.event === 'approved') {
      applyMap[r.user_id].approved++;
      const d = safeDate(r.date);
      if (d > applyMap[r.user_id].last_date) applyMap[r.user_id].last_date = d;
    }
  });
  // キャンセル済み・未来シフトを除外し、最終状態が approved の過去分のみカウント
  const today___ = formatDate(getJstNow());
  const unpaidMap = {};
  Object.values(buildLatestAppliedMap_(allApplied)).forEach(r => {
    if (r.event === 'approved') {
      const d = safeDate(r.date);
      if (d > today___) return;
      if (!paidSet.has(r.user_id + '_' + d)) unpaidMap[r.user_id] = (unpaidMap[r.user_id] || 0) + 1;
    }
  });

  const staff = staffList.map(s => ({
    ...s,
    apply_count:    (applyMap[s.line_user_id] || {}).applied   || 0,
    approved_count: (applyMap[s.line_user_id] || {}).approved  || 0,
    last_date:      (applyMap[s.line_user_id] || {}).last_date || '',
    unpaid_count: unpaidMap[s.line_user_id] || 0,
  }));

  return JSON.stringify({ staff });
}

function getAppliedAllWithArchives_(options) {
  return getRowsWithArchives_(SHEET.APPLIED, 'Applied', options).map(r => ({
    ...r,
    date: safeDate(r.date),
  }));
}

function getAllWorkLogWithArchives_(options) {
  return getRowsWithArchives_(SHEET.WORKLOG, 'WorkLog', options);
}

function getWorkLogByMonthIncludingArchive_(year, month) {
  const prefix = year + '-' + String(month).padStart(2, '0');
  const rows = getAllWorkLogWithArchives_({ yearMonth: prefix });
  const corrMap = _buildWLCorrMap_(rows);
  return rows
    .filter(r => !r.corrects)
    .filter(r => r.action !== 'deleted')
    .filter(r => !(corrMap[r.log_id] && corrMap[r.log_id].action === 'deleted'));
}

function getRowsWithArchives_(currentSheetName, archivePrefix, options) {
  options = options || {};
  const rows = [];
  function pushRowsFromSheet_(sh) {
    sheetToObjects(sh).forEach(r => {
      const d = safeDate(r.date || r.timestamp);
      if (options.yearMonth && !d.startsWith(options.yearMonth)) return;
      if (options.userId && String(r.user_id || '') !== String(options.userId)) return;
      rows.push(r);
    });
  }
  try {
    pushRowsFromSheet_(getSheet(currentSheetName));
  } catch(e) {
    Logger.log('[getRowsWithArchives_] current skip: ' + currentSheetName + ' ' + e.message);
  }
  getSpreadsheet().getSheets().forEach(sh => {
    const name = sh.getName();
    if (!new RegExp('^' + archivePrefix + '_\\d{4}_\\d{2}$').test(name)) return;
    if (options.yearMonth && name !== archivePrefix + '_' + options.yearMonth.replace('-', '_')) return;
    pushRowsFromSheet_(sh);
  });
  if (options.dedupe === false) return rows;

  const keyed = {};
  const unkeyed = [];
  rows.forEach(r => {
    const key = r.dedupe_key || r.log_id || '';
    if (!key) {
      unkeyed.push(r);
      return;
    }
    if (isNewerRow_(r, keyed[key])) keyed[key] = r;
  });
  return unkeyed.concat(Object.values(keyed));
}

function _normalizeJob_(j) {
  return {
    ...j,
    start_time: _safeTime_(j.start_time),
    end_time:   _safeTime_(j.end_time),
    active:     j.active === true || String(j.active).toLowerCase() === 'true',
  };
}

function api_getJobMaster() {
  const jobs = sheetToObjects(getSheet(SHEET.JOB))
    .filter(j => String(j.active).toLowerCase() === 'true')
    .map(_normalizeJob_);
  return JSON.stringify(jobs);
}

function api_getJobMasterAll() {
  return JSON.stringify(sheetToObjects(getSheet(SHEET.JOB)).map(_normalizeJob_));
}

function api_approveShift(userId, date, jobId) {
  try {
    return JSON.stringify(approveShift(userId, date, jobId));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_approveShifts(approvalsJson) {
  try {
    var approvals = JSON.parse(approvalsJson);
    return JSON.stringify(approveShifts(approvals));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 管理者がカレンダー上から直接シフトを追加する
function api_addShiftForDate(userId, date, jobId, directApprove) {
  if (!userId || !date || !jobId) return JSON.stringify({ ok: false, msg: '必須項目が不足しています' });
  try {
    date = String(date).slice(0, 10);
    if (directApprove === true || directApprove === 'true') {
      const result = approveShift(userId, date, jobId);
      return JSON.stringify(result);
    } else {
      const added = appendApplied('applied', userId, date, jobId, 'admin', '');
      if (!added) return JSON.stringify({ ok: false, msg: '既に応募済みです（重複）' });
      return JSON.stringify({ ok: true, msg: '応募を追加しました' });
    }
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 確定＆支払いを一括処理
function api_approveAndPay(userId, date, jobId) {
  if (!userId || !date || !jobId) return JSON.stringify({ ok: false, msg: '必須項目が不足しています' });
  date = String(date).slice(0, 10);
  if (date > formatDate(getJstNow())) {
    return JSON.stringify({ ok: false, msg: '未来日の支払いはできません' });
  }
  const job = getJob(jobId);

  // 1. 未確定なら確定（LINE通知含む）
  const status = getShiftStatus(userId, date, jobId);
  if (status !== 'approved') {
    const apResult = approveShift(userId, date, jobId);
    if (!apResult.ok) return JSON.stringify({ ok: false, msg: '確定処理失敗: ' + apResult.msg });
  }

  // 2. PayRecord 作成（重複スキップ内蔵）
  recordPay(userId, date, jobId, job);

  // 3. pay_id を取得して支払済みにする
  const sh = getSheet(SHEET.PAY_RECORD);
  if (sh.getLastRow() < 2) return JSON.stringify({ ok: false, msg: 'PayRecord 作成失敗' });
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const uidIdx  = hdr.indexOf('user_id');
  const dateIdx = hdr.indexOf('date');
  const jobIdx  = hdr.indexOf('job_id');
  const pidIdx  = hdr.indexOf('pay_id');
  const match = rows.find(r =>
    r[uidIdx] === userId && safeDate(r[dateIdx]) === date && r[jobIdx] === jobId
  );
  if (!match) return JSON.stringify({ ok: false, msg: 'PayRecord が見つかりません' });
  markAsPaid(match[pidIdx]);

  return JSON.stringify({ ok: true, msg: '確定・支払い完了' });
}

function api_cancelShift(userId, date, jobId, reason) {
  return JSON.stringify(cancelShift(userId, date, jobId, reason));
}

// Applied キャンセル + WorkLog の approved/manual/worked を論理削除（完全削除）
function api_fullDeleteShift(userId, date, jobId) {
  if (!userId || !date) return JSON.stringify({ ok: false, msg: '必須パラメータ不足' });
  date = String(date).slice(0, 10);

  const errors = [];
  let wlDeleted = 0, prDeleted = 0;

  // 1. Applied をキャンセル（jobId が空でも全ジョブ対象）
  try {
    if (jobId) {
      cancelShift(userId, date, jobId, 'admin_full_delete');
    } else {
      const allApplied = Object.values(buildLatestAppliedMap_(getAppliedAllWithArchives_())).filter(r =>
        r.user_id === userId && safeDate(r.date) === date && r.event === 'approved'
      );
      allApplied.forEach(r => cancelShift(userId, date, r.job_id, 'admin_full_delete'));
    }
  } catch(e) {
    errors.push('Applied: ' + e.message);
    Logger.log('[fullDeleteShift] Applied error: ' + e.message);
  }

  // 2. WorkLog の対象行を論理削除
  try {
    const sh = getSheet(SHEET.WORKLOG);
    const hm = buildHeaderMap(sh);
    const rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
    const uidIdx = hm['user_id'], dateIdx = hm['date'], jidIdx = hm['job_id'];
    const actionIdx = hm['action'], lidIdx = hm['log_id'], corIdx = hm['corrects'];

    const targets = rows.filter((r) => {
      if (r[corIdx]) return false;
      if (r[uidIdx] !== userId) return false;
      if (safeDate(r[dateIdx]) !== date) return false;
      if (jobId && r[jidIdx] !== jobId) return false;
      return ['approved', 'manual', 'worked', 'checkin', 'checkout'].includes(r[actionIdx]);
    });

    const now = formatDateTime(getJstNow());
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const currentLogIds = new Set();
    targets.forEach(orig => {
      const corrRow = orig.slice();
      if (lidIdx !== undefined && orig[lidIdx]) currentLogIds.add(orig[lidIdx]);
      if (hm['timestamp'] !== undefined) corrRow[hm['timestamp']] = now;
      if (lidIdx         !== undefined) corrRow[lidIdx]            = uuid();
      if (corIdx         !== undefined) corrRow[corIdx]            = orig[lidIdx] || '';
      if (actionIdx      !== undefined) corrRow[actionIdx]         = 'deleted';
      sh.appendRow(corrRow);
    });

    const allLogRows = getAllWorkLogWithArchives_({ userId }).filter(r => safeDate(r.date) === date);
    const allCorrMap = _buildWLCorrMap_(allLogRows);
    const archiveTargets = allLogRows.filter(r => {
      if (!r.log_id || currentLogIds.has(r.log_id)) return false;
      if (r.corrects || r.action === 'deleted') return false;
      if (allCorrMap[r.log_id] && allCorrMap[r.log_id].action === 'deleted') return false;
      if (jobId && r.job_id !== jobId) return false;
      return ['approved', 'manual', 'worked', 'checkin', 'checkout'].includes(r.action);
    });
    archiveTargets.forEach(orig => {
      const corrObj = Object.assign({}, orig, {
        timestamp: now,
        log_id: uuid(),
        corrects: orig.log_id || '',
        action: 'deleted',
      });
      sh.appendRow(objToRow(hdr, corrObj));
    });

    wlDeleted = targets.length + archiveTargets.length;
  } catch(e) {
    errors.push('WorkLog: ' + e.message);
    Logger.log('[fullDeleteShift] WorkLog error: ' + e.message);
  }

  // 3. PayRecord の未払い行を物理削除（支払済みは除外）
  try {
    const paySh = getSheet(SHEET.PAY_RECORD);
    if (paySh.getLastRow() >= 2) {
      const payHm = buildHeaderMap(paySh);
      const uidPayIdx  = payHm['user_id'];
      const datePayIdx = payHm['date'];
      const jidPayIdx  = payHm['job_id'];
      const paidPayIdx = payHm['paid'];
      if (uidPayIdx !== undefined && datePayIdx !== undefined && paidPayIdx !== undefined) {
        const payRows = paySh.getRange(2, 1, paySh.getLastRow() - 1, paySh.getLastColumn()).getValues();
        const toDeletePay = [];
        payRows.forEach((r, i) => {
          if (r[uidPayIdx] !== userId) return;
          if (safeDate(r[datePayIdx]) !== date) return;
          if (jobId && jidPayIdx !== undefined && r[jidPayIdx] !== jobId) return;
          const paid = r[paidPayIdx] === true || String(r[paidPayIdx]).toLowerCase() === 'true';
          if (!paid) toDeletePay.push(i + 2);
        });
        for (let i = toDeletePay.length - 1; i >= 0; i--) paySh.deleteRow(toDeletePay[i]);
        prDeleted = toDeletePay.length;
      }
    }
  } catch(e) {
    errors.push('PayRecord: ' + e.message);
    Logger.log('[fullDeleteShift] PayRecord error: ' + e.message);
  }

  if (errors.length > 0) {
    return JSON.stringify({ ok: false, msg: '一部失敗: ' + errors.join(', '), wlDeleted, prDeleted });
  }
  return JSON.stringify({ ok: true, msg: '削除しました', wlDeleted, prDeleted });
}

function api_broadcastToAll(msg) {
  try {
    const count = broadcastToAll(msg);
    return JSON.stringify({ ok: true, count });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 物理削除（管理者専用・初期データ誤り等の修正用）
// Applied / WorkLog / PayRecord から userId+date+jobId に一致する行を完全削除
function api_adminHardDelete(userId, date, jobId) {
  if (!userId || !date || !jobId) return JSON.stringify({ ok: false, msg: '必須パラメータ不足' });
  if (_safeProp('ALLOW_HARD_DELETE', 'false') !== 'true') {
    return JSON.stringify({ ok: false, msg: '完全削除は無効です。ALLOW_HARD_DELETE=true の時だけ実行できます' });
  }
  date = String(date).slice(0, 10);
  const ss = getSpreadsheet();

  function deleteMatchingRows(sheetName, uidCol, dateCol, jidCol) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return 0;
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const ui = hdr.indexOf(uidCol), di = hdr.indexOf(dateCol), ji = hdr.indexOf(jidCol);
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const toDelete = [];
    data.forEach((r, i) => {
      if (r[ui] === userId && safeDate(r[di]) === date && r[ji] === jobId) toDelete.push(i + 2);
    });
    for (let i = toDelete.length - 1; i >= 0; i--) sh.deleteRow(toDelete[i]);
    return toDelete.length;
  }

  const a = deleteMatchingRows(SHEET.APPLIED,    'user_id', 'date',      'job_id');
  const w = deleteMatchingRows(SHEET.WORKLOG,    'user_id', 'date',      'job_id');
  const p = deleteMatchingRows(SHEET.PAY_RECORD, 'user_id', 'date',      'job_id');
  const c = deleteMatchingRows(SHEET.CANCEL,     'user_id', 'date',      'job_id');

  Logger.log(`adminHardDelete: Applied=${a} WorkLog=${w} PayRecord=${p} Cancel=${c}`);
  return JSON.stringify({ ok: true, msg: `完全削除しました（Applied:${a} WorkLog:${w} PayRecord:${p}）` });
}

function api_getMonthlyPayroll(year, month) {
  return JSON.stringify(getMonthlyPayroll(Number(year), Number(month)));
}

function api_getWorkLogData(yearMonth, fromArchive) {
  return JSON.stringify(getWorkLogData(yearMonth, fromArchive));
}

function api_getArchiveSheetNames() {
  return JSON.stringify(getArchiveSheetNames());
}

function api_saveEvaluation(userId, workDate, data) {
  try {
  const d = typeof data === 'string' ? JSON.parse(data) : data;
  const sh = getSheet(SHEET.EVAL);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const scores = [
    d.punctuality, d.attitude, d.instruction,
    d.accuracy, d.speed, d.report,
    d.teamwork, d.rehire,
  ].map(Number);
  const avg = scores.reduce((s, v) => s + v, 0) / scores.length;
  const rounded = Math.round(avg * 10) / 10;
  let level = 'C';
  if (rounded >= 4.5) level = 'S';
  else if (rounded >= 4.0) level = 'A';
  else if (rounded >= 3.0) level = 'B';
  else if (rounded >= 2.0) level = 'C';
  else level = 'D';

  sh.appendRow(objToRow(hdr, {
    timestamp:          formatDateTime(getJstNow()),
    user_id:            userId,
    work_date:          workDate,
    score_punctuality:  scores[0],
    score_attitude:     scores[1],
    score_instruction:  scores[2],
    score_accuracy:     scores[3],
    score_speed:        scores[4],
    score_report:       scores[5],
    score_teamwork:     scores[6],
    score_rehire:       scores[7],
    score_avg:          rounded,
    level,
    strong_point:       d.strong_point || '',
    caution:            d.caution || '',
    next_policy:        d.next_policy || '',
  }));

  refreshStaffEval(userId);
  return JSON.stringify({ ok: true, level, avg: rounded });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_resendUrl(userId, type) {
  const token = _issueToken(userId, type);
  const baseUrl = ScriptApp.getService().getUrl();
  const url = baseUrl + '?page=' + type + '&token=' + token;
  pushMessage(userId, 'リンクをお送りします：\n' + url);
  return JSON.stringify({ ok: true, url });
}

function api_upsertJob(data) {
  const d = typeof data === 'string' ? JSON.parse(data) : data;
  const sh = getSheet(SHEET.JOB);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
  const jidIdx = hdr.indexOf('job_id');
  const now = formatDateTime(getJstNow());

  if (d.job_id) {
    const found = rows.findIndex(r => r[jidIdx] === d.job_id);
    if (found >= 0) {
      sh.getRange(found + 2, 1, 1, hdr.length).setValues([objToRow(hdr, { ...d })]);
      return JSON.stringify({ ok: true, action: 'updated' });
    }
  }
  // 新規
  const newId = d.job_id || ('JOB' + String(rows.length + 1).padStart(3, '0'));
  sh.appendRow(objToRow(hdr, { ...d, job_id: newId, created_at: now }));
  return JSON.stringify({ ok: true, action: 'created', job_id: newId });
}

function api_markAsPaid(payId) {
  try {
    return JSON.stringify({ ok: markAsPaid(payId) });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_markMultiplePaid(payIdsJson) {
  try {
    const ids = JSON.parse(payIdsJson || '[]');
    let count = 0;
    ids.forEach(id => { if (id && markAsPaid(id)) count++; });
    return JSON.stringify({ ok: true, count });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// ─── vol.2 追加 API ───────────────────────

// LINE一括募集メッセージ送信
function api_sendRecruitMessage(datesJson, jobId) {
  return api_sendRecruitWithTokens(jobId, datesJson);
}

// 管理者による直接シフト割当（承認済みで追加）
function api_addShiftDirect(userId, date, jobId, notify) {
  if (!userId || !date || !jobId) return JSON.stringify({ ok: false, msg: '必須項目が不足しています' });
  try {
    date = String(date).slice(0, 10);
    const sendNotify = notify === true || notify === 'true' || notify === undefined;
    const result = approveShift(userId, date, jobId, sendNotify);
    return JSON.stringify(result);
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 一括支払確定（api_markMultiplePaid の別名）
function api_markAsPaidBulk(payIdsJson) {
  return api_markMultiplePaid(payIdsJson);
}

// pay_id がない場合は PayRecord を作成してから支払済みにする
function api_markPaidByKeys(recordsJson) {
  try {
    const records = JSON.parse(recordsJson || '[]');
    const sh = getSheet(SHEET.PAY_RECORD);
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const existingRows = sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
      : [];
    const uidIdx  = hdr.indexOf('user_id');
    const dateIdx = hdr.indexOf('date');
    const jobIdx  = hdr.indexOf('job_id');
    const now = formatDateTime(getJstNow());
    let count = 0;

    records.forEach(r => {
      if (r.pay_id) {
        if (markAsPaid(r.pay_id)) count++;
      } else {
        const exists = existingRows.some(row =>
          row[uidIdx] === r.user_id &&
          safeDate(row[dateIdx]) === safeDate(r.date) &&
          row[jobIdx] === r.job_id
        );
        if (exists) { Logger.log('PayRecord重複スキップ: ' + [r.user_id, r.date, r.job_id].join(':')); return; }
        const wage  = Number(r.wage  || 0);
        const trans = Number(r.transport || 0);
        const wh    = calcWithholding(wage, trans).withholding;
        const newId = uuid();
        sh.appendRow(objToRow(hdr, {
          pay_id:         newId,
          user_id:        r.user_id,
          date:           r.date,
          job_id:         r.job_id,
          hourly_wage:    0,
          actual_hours:   0,
          wage:           wage,
          transport:      trans,
          withholding:    wh,
          net_pay:        wage + trans - wh,
          paid:           'true',
          paid_at:        now,
          payment_method: 'cash',
          receipt_url:    '',
          note:           '',
          created_at:     now,
        }));
        count++;
      }
    });

    return JSON.stringify({ ok: true, count });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 職種でスタッフをフィルタ（job_id列がなければ全員返す）
function api_getStaffByJobId(jobId) {
  const all = getAllStaff();
  if (!jobId) return JSON.stringify(all);
  const filtered = all.filter(s => {
    if (!s.job_id) return true; // job_id未設定は全職種対応
    return String(s.job_id).split(',').map(x => x.trim()).includes(String(jobId));
  });
  return JSON.stringify(filtered.length ? filtered : all);
}

// WorkLog に補正エントリを追記（append-only 維持）
// dataJson: { date?, job_id?, start_time?, end_time?, transport?, action? }
// action='deleted' の場合は論理削除
function api_addWorkLogCorrection(originalLogId, dataJson) {
  if (!originalLogId) return JSON.stringify({ ok: false, msg: 'log_id が必要です' });
  const data = typeof dataJson === 'string' ? JSON.parse(dataJson) : dataJson;
  const sh = getSheet(SHEET.WORKLOG);
  const hm = buildHeaderMap(sh);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  // 元の行を検索
  const rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
  const lidIdx = hm['log_id'];
  const orig = lidIdx !== undefined ? rows.find(r => r[lidIdx] === originalLogId) : null;
  let sourceRow = orig;
  if (!sourceRow) {
    const archived = getAllWorkLogWithArchives_().find(r => r.log_id === originalLogId);
    if (archived) sourceRow = objToRow(hdr, archived);
  }
  if (!sourceRow) return JSON.stringify({ ok: false, msg: '対象ログが見つかりません' });

  // 補正行を追記
  const now = formatDateTime(getJstNow());
  const corrRow = sourceRow.slice(); // 元の行をベースにコピー

  if (hm['timestamp']    !== undefined) corrRow[hm['timestamp']]  = now;
  if (hm['log_id']       !== undefined) corrRow[hm['log_id']]     = uuid();
  if (hm['corrects']     !== undefined) corrRow[hm['corrects']]   = originalLogId;
  if (hm['action']       !== undefined) corrRow[hm['action']]     = data.action || 'correction';

  if (data.date      && hm['date']       !== undefined) corrRow[hm['date']]       = data.date;
  if (data.job_id    && hm['job_id']     !== undefined) corrRow[hm['job_id']]     = data.job_id;
  if (data.start_time&& hm['start_time'] !== undefined) corrRow[hm['start_time']] = data.start_time;
  if (data.end_time  && hm['end_time']   !== undefined) corrRow[hm['end_time']]   = data.end_time;
  if (data.transport !== undefined && hm['transport'] !== undefined) corrRow[hm['transport']] = Number(data.transport) || 0;

  // 時間再計算（削除でない場合）
  if (data.action !== 'deleted' && data.start_time && data.end_time) {
    const job = getJob(data.job_id || (hm['job_id'] !== undefined ? sourceRow[hm['job_id']] : ''));
    if (job) {
      const calc = calcDailyWage(job.wage, data.start_time, data.end_time, Number(data.transport) || 0);
      const pay  = calcWithholding(calc.dailyWage, calc.transport);
      if (hm['actual_hours'] !== undefined) corrRow[hm['actual_hours']] = calc.hours;
      if (hm['wage']         !== undefined) corrRow[hm['wage']]         = calc.dailyWage;
      if (hm['withholding']  !== undefined) corrRow[hm['withholding']]  = pay.withholding;
      if (hm['net_pay']      !== undefined) corrRow[hm['net_pay']]      = pay.netPay;
    }
  }

  sh.appendRow(corrRow);

  // PayRecord同期（削除でない補正かつ給与再計算済みの場合）
  if (data.action !== 'deleted' && data.start_time && data.end_time &&
      hm['user_id'] !== undefined && hm['date'] !== undefined && hm['job_id'] !== undefined) {
    const prSh  = getSheet(SHEET.PAY_RECORD);
    const prHdr = prSh.getRange(1, 1, 1, prSh.getLastColumn()).getValues()[0];
    if (prSh.getLastRow() > 1) {
      const prRows = prSh.getRange(2, 1, prSh.getLastRow() - 1, prSh.getLastColumn()).getValues();
      const pUid = prHdr.indexOf('user_id'), pDate = prHdr.indexOf('date'), pJob = prHdr.indexOf('job_id');
      const pWage = prHdr.indexOf('wage'), pWh = prHdr.indexOf('withholding'), pNet = prHdr.indexOf('net_pay');
      const pPaid = prHdr.indexOf('paid');
      const tUid = corrRow[hm['user_id']], tDate = safeDate(corrRow[hm['date']]), tJob = corrRow[hm['job_id']];
      for (let i = 0; i < prRows.length; i++) {
        if (prRows[i][pUid] === tUid && safeDate(prRows[i][pDate]) === tDate && prRows[i][pJob] === tJob) {
          if (pPaid >= 0 && isPaid_(prRows[i][pPaid])) break;
          // 1行まとめて読み直して書き戻す（バッチ処理）
          const prRowFull = prSh.getRange(i + 2, 1, 1, prHdr.length).getValues()[0];
          if (pWage >= 0) prRowFull[pWage] = corrRow[hm['wage']] || 0;
          if (pWh   >= 0) prRowFull[pWh]   = corrRow[hm['withholding']] || 0;
          if (pNet  >= 0) prRowFull[pNet]   = corrRow[hm['net_pay']] || 0;
          prSh.getRange(i + 2, 1, 1, prHdr.length).setValues([prRowFull]);
          break;
        }
      }
    }
  }

  return JSON.stringify({ ok: true });
}

// WorkLog に手動エントリを追記（action='manual'）
function api_addWorkLogManual(userId, date, jobId, startTime, endTime, transport) {
  if (!userId || !date || !jobId || !startTime || !endTime) {
    return JSON.stringify({ ok: false, msg: '必須項目が不足しています' });
  }
  const targetDate = String(date).slice(0, 10);
  const existing = sheetToObjects(getSheet(SHEET.WORKLOG)).filter(r =>
    r.user_id === userId && safeDate(r.date) === targetDate && r.job_id === jobId && r.action === 'manual'
  );
  if (existing.length > 0) {
    return JSON.stringify({ ok: false, msg: '同日・同職種の手動記録が既に存在します' });
  }
  const job = getJob(jobId);
  const hourlyWage = job ? Number(job.wage) : 0;
  const trans = Number(transport) || (job ? Number(job.transport) : 0);
  const calc = calcDailyWage(hourlyWage, startTime, endTime, trans);
  const pay  = calcWithholding(calc.dailyWage, calc.transport);

  appendWorkLog({
    user_id:      userId,
    date:         String(date).slice(0, 10),
    job_id:       jobId,
    action:       'manual',
    start_time:   startTime,
    end_time:     endTime,
    actual_hours: calc.hours,
    wage:         calc.dailyWage,
    transport:    calc.transport,
    withholding:  pay.withholding,
    net_pay:      pay.netPay,
    note:         '管理者手動追記',
  });

  // PayRecord にも追記
  recordPay(userId, String(date).slice(0, 10), jobId, job, {
    hourly_wage: hourlyWage,
    actual_hours: calc.hours,
    wage: calc.dailyWage,
    transport: calc.transport,
    withholding: pay.withholding,
    net_pay: pay.netPay,
  });
  return JSON.stringify({ ok: true });
}

// ──────────────────────────────────────────

function api_updatePayAmount(payId, hourlyWage, hours, transport) {
  try {
    const sh = getSheet(SHEET.PAY_RECORD);
    if (sh.getLastRow() < 2) return JSON.stringify({ ok: false, msg: 'データなし' });
    const hm   = buildHeaderMap(sh);
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const i    = rows.findIndex(r => r[hm['pay_id']] === payId);
    if (i < 0) return JSON.stringify({ ok: false, msg: '記録が見つかりません' });
    if (hm['paid'] !== undefined && isPaid_(rows[i][hm['paid']])) {
      return JSON.stringify({ ok: false, msg: '支払済みの金額は変更できません' });
    }

    if (Number(hourlyWage) < 0 || Number(hours) < 0) {
      return JSON.stringify({ ok: false, msg: '時給・時間は0以上を指定してください' });
    }
    const wage = Math.round(Number(hourlyWage) * Number(hours));
    const trans = Number(transport) || 0;
    const withholding = calcWithholding(wage, trans).withholding;
    const net_pay = wage + trans - withholding;
    const row = i + 2;

    // 1行分をまとめて読んで書き直す（バッチ処理）
    const fullRow = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    fullRow[hm['wage']]        = wage;
    fullRow[hm['transport']]   = trans;
    fullRow[hm['withholding']] = withholding;
    fullRow[hm['net_pay']]     = net_pay;
    if (hm['hourly_wage']  !== undefined) fullRow[hm['hourly_wage']]  = Number(hourlyWage);
    if (hm['actual_hours'] !== undefined) fullRow[hm['actual_hours']] = Number(hours);
    sh.getRange(row, 1, 1, sh.getLastColumn()).setValues([fullRow]);

    return JSON.stringify({ ok: true, wage, transport: trans, net_pay });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_getPayRecords(year, month) {
  return JSON.stringify(getPayRecordsByMonth(Number(year), Number(month)));
}

// 会計サマリー：指定月の支払い総額と手数料を返す
// 手数料率は Script Properties の FEE_RATE（既定20）。料金モデル未確定のため可変。
function api_getFeeSummary(year, month) {
  try {
    const s = getPaySummary(Number(year), Number(month));
    const feeRate = Number(_safeProp('FEE_RATE', '20')) || 20;
    const feeAmount = Math.round((Number(s.total_wage) || 0) * feeRate / 100);
    return JSON.stringify({
      ok:            true,
      year:          Number(year),
      month:         Number(month),
      count:         s.rows.length,
      total_net_pay: s.total_net_pay,
      total_wage:    s.total_wage,
      fee_rate:      feeRate,
      fee_amount:    feeAmount,
    });
  } catch (e) {
    console.error('[api_getFeeSummary] ' + e);
    return JSON.stringify({ ok: false, msg: '集計に失敗しました: ' + (e && e.message || e) });
  }
}

function api_getUnpaidRecords() {
  const today = formatDate(getJstNow());
  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s; });
  const jobMap = {};
  sheetToObjects(getSheet(SHEET.JOB)).forEach(j => { jobMap[j.job_id] = j; });

  // PayRecord を user_id:date:job_id でインデックス（支払済みキーと金額情報を保持）
  const paidKeys = new Set();
  const payMap = {};
  sheetToObjects(getSheet(SHEET.PAY_RECORD)).forEach(r => {
    const d = safeDate(r.date);
    if (!d) return;
    const key = r.user_id + ':' + d + ':' + r.job_id;
    const isPaid = isPaid_(r.paid);
    if (isPaid) paidKeys.add(key);
    if (!payMap[key] || (!isPaid_(payMap[key].paid) && isPaid)) payMap[key] = r;
  });

  // Applied の最新状態からキャンセル済みキーを収集
  // LINE経由キャンセルは WorkLog に deleted 補正を書かないため、Applied で判定する
  // job_id 空文字は 'JOB001' に補完（buildLatestAppliedMap_ の normalizeAppliedRow_ と合わせる）
  const cancelledKeys = new Set(
    Object.values(buildLatestAppliedMap_(getAppliedAllWithArchives_()))
      .filter(r => r.event === 'cancelled')
      .map(r => r.user_id + ':' + safeDate(r.date) + ':' + (r.job_id || 'JOB001'))
  );

  // WorkLog の実績行を正本として未払い一覧を生成（補正削除を除外）
  const allWLogs = getAllWorkLogWithArchives_();
  const wlCorrMap = _buildWLCorrMap_(allWLogs);
  const rank = { approved: 1, worked: 3, manual: 3, checkout: 4 };
  const effectiveMap = {};
  allWLogs
    .filter(r => !r.corrects)
    .filter(r => r.action !== 'deleted')
    .filter(r => !(wlCorrMap[r.log_id] && wlCorrMap[r.log_id].action === 'deleted'))
    .filter(r => {
      if (!rank[r.action]) return false;
      const d = safeDate(r.date);
      return d && d <= today;
    })
    .forEach(r => {
      const d = safeDate(r.date);
      const key = r.user_id + ':' + d + ':' + r.job_id;
      const ex = effectiveMap[key];
      if (!ex || rank[r.action] > rank[ex.action] || (rank[r.action] === rank[ex.action] && isNewerRow_(r, ex))) {
        effectiveMap[key] = r;
      }
    });

  const result = [];
  Object.keys(effectiveMap).forEach(key => {
    if (paidKeys.has(key)) return; // 支払済みはスキップ
    const r = effectiveMap[key];
    // approved のみ（実績なし）+ Applied でキャンセル済み → スキップ
    if (rank[r.action] <= 1 && cancelledKeys.has(key)) return;
    const d = safeDate(r.date);
    const pay = payMap[key] || {};
    const s = staffMap[r.user_id] || {};
    result.push({
      pay_id:     pay.pay_id || '',
      user_id:    r.user_id,
      date:       d,
      job_id:     r.job_id,
      wage:       pay.wage != null ? pay.wage : (r.wage || 0),
      transport:  pay.transport != null ? pay.transport : (r.transport || 0),
      net_pay:    pay.net_pay != null ? pay.net_pay : calcWithholding(Number(r.wage||0), Number(r.transport||0)).netPay,
      paid:       false,
      staff_name: s.name || s.display_name || r.user_id,
      job_title:  (jobMap[r.job_id] || {}).title || r.job_id,
    });
  });

  result.sort((a, b) => a.date > b.date ? 1 : a.date < b.date ? -1 : 0);
  return JSON.stringify(result);
}

function api_sendLineMessage(userId, text) {
  if (!userId || !text || !text.trim()) return JSON.stringify({ ok: false, msg: '内容を入力してください' });
  try {
    const ok = pushMessage(userId, text.trim());
    return JSON.stringify(ok ? { ok: true } : { ok: false, msg: 'LINE送信に失敗しました' });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_generateReceipt(userId, date) {
  try {
    return JSON.stringify(generateReceipt(userId, date));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_generatePayslip(userId, yearMonth) {
  try {
    return JSON.stringify(generatePayslip(userId, yearMonth));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 全スタッフ賃金台帳を1ファイルで生成
function api_generateBulkPayslip(yearMonth) {
  try {
    return JSON.stringify(generateBulkPayslipCombined(yearMonth));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_generateAttendance(userId, yearMonth) {
  try {
    return JSON.stringify(generateAttendanceSheet(userId, yearMonth));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_backfillFromChatLog() {
  try {
    const added = backfillFromChatLog();
    return JSON.stringify({ ok: true, added: added });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_notifyAprilApproved() {
  return JSON.stringify(notifyAprilApproved());
}

function api_calcYearlySummary(year) {
  return JSON.stringify(calcYearlySummary(Number(year)));
}

function api_getPaySummaryByYear(year) {
  const rows = sheetToObjects(getSheet(SHEET.PAY_SUMM))
    .filter(r => Number(r.year) === Number(year));
  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s; });
  const result = rows.map(r => ({
    ...r,
    staff_name: (staffMap[r.user_id] || {}).name || (staffMap[r.user_id] || {}).display_name || r.user_id,
  }));
  return JSON.stringify(result);
}

function api_generatePaymentReport(userId, year) {
  try {
    return JSON.stringify(generatePaymentReport(userId, Number(year)));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_generateWithholdingCert(userId, year, endDate) {
  try {
    return JSON.stringify(generateWithholdingCert(userId, Number(year), endDate || null));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_generateBulkWithholdingCert(year, endDate) {
  try {
    return JSON.stringify(generateBulkWithholdingCert(Number(year), endDate || null));
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// Drive書類フォルダのURLを返す（名前で検索）
function api_getDocFolderUrl(folderName) {
  try {
    // 賃金台帳のみ実際のPDF保存先（generateBulkPayslipCombined/60_Docs.gs L301と同一フォルダ）を返す。
    // 他の書類タイプは従来通り getDriveFolder（社内タイミー配下のサブフォルダ）を使う。
    if (folderName === '賃金台帳') {
      const folder = DriveApp.getFolderById('198ubbF5s16KukcE0H67fy8gdh8Jw-7k5');
      return JSON.stringify({ ok: true, url: folder.getUrl() });
    }
    const folder = getDriveFolder(folderName);
    return JSON.stringify({ ok: true, url: folder.getUrl() });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 納期の特例集計: half=1→1〜6月（7/10納付）, half=2→7〜12月（翌1/20納付）
function api_getWithholdingByHalf(year, half) {
  const y = Number(year);
  const h = Number(half);
  const startDate = h === 1 ? y + '-01-01' : y + '-07-01';
  const endDate   = h === 1 ? y + '-06-30' : y + '-12-31';
  const dueDate   = h === 1 ? (y + '年7月10日') : ((y + 1) + '年1月20日');

  const _allLogsHalf = getAllWorkLogWithArchives_({});
  const _halfCorrMap = _buildWLCorrMap_(_allLogsHalf);
  // checkout が存在する日は approved を除外（二重カウント防止）
  const _halfFiltered = _allLogsHalf
    .filter(r => { const d = safeDate(r.date); return d >= String(startDate) && d <= String(endDate); })
    .filter(r => (r.action === 'checkout' || r.action === 'approved') && !r.corrects)
    .filter(r => !(_halfCorrMap[r.log_id] && _halfCorrMap[r.log_id].action === 'deleted'));
  const _checkoutKeys = new Set(
    _halfFiltered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const logs = _halfFiltered.filter(r =>
    r.action === 'checkout' || !_checkoutKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );

  const totalWage        = logs.reduce((s, r) => s + (Number(r.wage) || 0), 0);
  const totalWithholding = logs.reduce((s, r) => s + (Number(r.withholding) || 0), 0);
  const staffIds         = [...new Set(logs.map(r => r.user_id).filter(Boolean))];

  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s.name || s.display_name || s.line_user_id; });

  const breakdown = staffIds.map(uid => {
    const sl = logs.filter(r => r.user_id === uid);
    return {
      name:        staffMap[uid] || uid,
      wage:        sl.reduce((s, r) => s + (Number(r.wage) || 0), 0),
      withholding: sl.reduce((s, r) => s + (Number(r.withholding) || 0), 0),
      days:        new Set(sl.map(r => safeDate(r.date))).size,
    };
  }).sort((a, b) => b.wage - a.wage);

  return JSON.stringify({
    ok: true,
    year: y, half: h,
    period: startDate + ' 〜 ' + endDate,
    dueDate,
    totalWage, totalWithholding,
    staffCount: staffIds.length,
    breakdown,
  });
}

function api_getTodayCheckin() {
  const today = formatDate(getJstNow());
  const latestMap = buildLatestAppliedMap_(
    getAppliedAllWithArchives_().filter(r => safeDate(r.date) === today)
  );
  const approved = Object.values(latestMap)
    .filter(r => r.event === 'approved');
  const seen = new Set();
  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s; });

  // WorkLogを1回だけ取得してuserIdでグループ化（補正削除を除外）
  const todayRaw = sheetToObjects(getSheet(SHEET.WORKLOG))
    .filter(w => safeDate(w.date) === today);
  const todayCorrMap = _buildWLCorrMap_(todayRaw);
  const allLogs = todayRaw
    .filter(w => !w.corrects && w.action !== 'deleted')
    .filter(w => !(todayCorrMap[w.log_id] && todayCorrMap[w.log_id].action === 'deleted'));
  const logsByUser = {};
  allLogs.forEach(w => {
    if (!logsByUser[w.user_id]) logsByUser[w.user_id] = [];
    logsByUser[w.user_id].push(w);
  });

  return JSON.stringify(approved.filter(r => {
    if (seen.has(r.user_id)) return false;
    seen.add(r.user_id);
    return true;
  }).map(r => {
    const logs = logsByUser[r.user_id] || [];
    const checkin  = logs.find(w => w.action === 'checkin');
    const checkout = logs.find(w => w.action === 'checkout');
    const staff = staffMap[r.user_id] || {};
    return {
      user_id:      r.user_id,
      display_name: staff.name || staff.display_name || r.user_id,
      job_id:       r.job_id,
      checkin_time: checkin ? checkin.start_time : '',
      checkout_time: checkout ? checkout.end_time : '',
      status: checkout ? 'done' : (checkin ? 'working' : 'waiting'),
    };
  }));
}

// フォームからの応募処理（ApplyForm.html用）
function api_applyFromForm(token, datesJson, jobId) {
  const td = _getTokenData(token);
  if (!td || td.type !== 'apply') return JSON.stringify({ ok: false, msg: 'トークン無効' });
  const guard = _validateApplySelection_(td, datesJson, jobId || 'JOB001');
  if (!guard.ok) return JSON.stringify({ ok: false, msg: guard.msg });
  const dates = guard.dates;
  const targetJobId = guard.jobId;
  const added = [];
  dates.forEach(d => {
    const ok = appendApplied('applied', guard.userId, d, targetJobId, 'form', '');
    if (ok) added.push(d);
  });
  if (added.length > 0) {
    try {
      const staff = getStaff(guard.userId);
      const staffName = (staff && (staff.name || staff.display_name)) || guard.userId;
      const DOW = ['日','月','火','水','木','金','土'];
      const dateLabels = added.map(ds => {
        const [y, m, d] = ds.split('-').map(Number);
        return `${m}/${d}（${DOW[new Date(y, m - 1, d).getDay()]}）`;
      });
      pushMessage(SECRETS.ADMIN_USER_ID, `📩 フォームから応募がありました\n${staffName} さん\n${dateLabels.join('、')}`);
    } catch(e) { /* 無視 */ }
  }
  return JSON.stringify({ ok: true, added });
}

// LIFF経由の名前登録（リッチメニューからのアクセス用）
function api_registerFromLiff(userId, name, phone, zipCode, address, birthDate, furigana, myNumber) {
  if (!userId) return JSON.stringify({ ok: false, msg: 'ユーザーID不明' });
  const n = (name || '').trim();
  if (!n) return JSON.stringify({ ok: false, msg: '名前を入力してください' });
  upsertStaff(userId, '');
  const updates = { name: n };
  if (phone     && String(phone).trim())     updates.phone      = String(phone).trim();
  if (zipCode   && String(zipCode).trim())   updates.zip_code   = String(zipCode).trim();
  if (address   && String(address).trim())   updates.address    = String(address).trim();
  if (birthDate && String(birthDate).trim()) updates.birth_date = String(birthDate).trim();
  if (furigana  && String(furigana).trim())  updates.furigana   = String(furigana).trim();
  if (myNumber  && String(myNumber).trim())  updates.my_number  = String(myNumber).trim();
  updateStaff(userId, updates);
  try { pushMessage(userId, `✅ 登録完了！\n${n}さん、よろしくお願いします 🙏`); } catch(e) {}
  try { pushMessage(SECRETS.ADMIN_USER_ID, `📝 プロフィール登録\n${n}さんが情報を登録しました`); } catch(e) {}
  return JSON.stringify({ ok: true });
}

// フォームからの名前登録（RegisterForm.html用・トークン版）
function api_registerFromForm(token, name, phone, zipCode, address, birthDate, furigana, myNumber) {
  const td = _getTokenData(token);
  if (!td || td.type !== 'register') return JSON.stringify({ ok: false, msg: 'トークン無効' });
  const n = (name || '').trim();
  if (!n) return JSON.stringify({ ok: false, msg: '名前を入力してください' });
  const updates = { name: n };
  if (phone     && String(phone).trim())     updates.phone      = String(phone).trim();
  if (zipCode   && String(zipCode).trim())   updates.zip_code   = String(zipCode).trim();
  if (address   && String(address).trim())   updates.address    = String(address).trim();
  if (birthDate && String(birthDate).trim()) updates.birth_date = String(birthDate).trim();
  if (furigana  && String(furigana).trim())  updates.furigana   = String(furigana).trim();
  if (myNumber  && String(myNumber).trim())  updates.my_number  = String(myNumber).trim();
  updateStaff(td.userId, updates);
  try { pushMessage(td.userId, `✅ 登録完了！\n${n}さん、よろしくお願いします 🙏`); } catch(e) {}
  try { pushMessage(SECRETS.ADMIN_USER_ID, `📝 プロフィール登録\n${n}さんが情報を登録しました`); } catch(e) {}
  return JSON.stringify({ ok: true });
}

// 登録フォーム用：現在の登録値を返す（LIFFパス）
function api_getMyProfile(userId) {
  if (!userId) return JSON.stringify({ ok: false });
  const s = getStaff(userId);
  if (!s) return JSON.stringify({ ok: true, name: '', phone: '', zip_code: '', address: '', birth_date: '', furigana: '', my_number: '' });
  return JSON.stringify({
    ok: true,
    name:       s.name                                                     || '',
    phone:      (s.phone     instanceof Date) ? '' : String(s.phone     || ''),
    zip_code:   (s.zip_code  instanceof Date) ? '' : String(s.zip_code  || ''),
    address:    (s.address   instanceof Date) ? '' : String(s.address   || ''),
    birth_date: safeDate(s.birth_date),
    furigana:   (s.furigana  instanceof Date) ? '' : String(s.furigana  || ''),
    my_number:  (s.my_number instanceof Date) ? '' : String(s.my_number || ''),
  });
}

// 登録フォーム用：現在の登録値を返す（トークンパス）
function api_getProfileByToken(token) {
  const td = _getTokenData(token);
  if (!td || td.type !== 'register') return JSON.stringify({ ok: false });
  const s = getStaff(td.userId);
  if (!s) return JSON.stringify({ ok: true, name: '', phone: '', zip_code: '', address: '', birth_date: '', furigana: '', my_number: '' });
  return JSON.stringify({
    ok: true,
    name:       s.name                                                     || '',
    phone:      (s.phone     instanceof Date) ? '' : String(s.phone     || ''),
    zip_code:   (s.zip_code  instanceof Date) ? '' : String(s.zip_code  || ''),
    address:    (s.address   instanceof Date) ? '' : String(s.address   || ''),
    birth_date: safeDate(s.birth_date),
    furigana:   (s.furigana  instanceof Date) ? '' : String(s.furigana  || ''),
    my_number:  (s.my_number instanceof Date) ? '' : String(s.my_number || ''),
  });
}

// 個別スタッフにプロファイル登録リンクをLINE送信
function api_sendProfileLink(uid) {
  if (!uid) return JSON.stringify({ ok: false, msg: 'uid不明' });
  const baseUrl = ScriptApp.getService().getUrl();
  const token = _issueToken(uid, 'register', {}, 7 * 24 * 3600);
  const url = baseUrl + '?page=register&token=' + token;
  try {
    pushMessage(uid, '📝 個人情報の登録をお願いします\n\n氏名・住所・生年月日・電話番号を以下のリンクから登録してください。\n\n' + url + '\n\n※ リンクは7日間有効です');
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 全スタッフ（active）にプロファイル登録リンクをLINE一斉送信
function api_sendProfileLinkAll() {
  const staff = getAllStaff().filter(s => String(s.active).toLowerCase() !== 'false');
  let sent = 0, failed = 0;
  const baseUrl = ScriptApp.getService().getUrl();
  staff.forEach(s => {
    if (!s.line_user_id) { failed++; return; }
    const token = _issueToken(s.line_user_id, 'register', {}, 7 * 24 * 3600);
    const url = baseUrl + '?page=register&token=' + token;
    try {
      pushMessage(s.line_user_id, '📝 個人情報の登録をお願いします\n\n氏名・住所・生年月日・電話番号を以下のリンクから登録してください。\n\n' + url + '\n\n※ リンクは7日間有効です');
      sent++;
    } catch(e) {
      Logger.log('sendProfileLinkAll失敗: ' + s.line_user_id + ' ' + e.message);
      failed++;
    }
  });
  return JSON.stringify({ ok: true, sent, failed });
}

// チェックインWebページ用（CheckIn.html用）
function api_processCheckinWeb(userId) {
  const today = formatDate(getJstNow());
  const rawLogs = sheetToObjects(getSheet(SHEET.WORKLOG))
    .filter(r => r.user_id === userId && safeDate(r.date) === today);
  const webCorrMap = _buildWLCorrMap_(rawLogs);
  const logs = rawLogs
    .filter(r => !r.corrects && r.action !== 'deleted')
    .filter(r => !(webCorrMap[r.log_id] && webCorrMap[r.log_id].action === 'deleted'))
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  const checkedIn  = logs.find(r => r.action === 'checkin');
  const checkedOut = logs.find(r => r.action === 'checkout');
  const staff = getStaff(userId);
  const name  = staff ? (staff.name || staff.display_name || '') : '';

  if (checkedOut) {
    return JSON.stringify({ type: 'already_done', name });
  }

  const now = getJstNow();
  const timeStr = formatTime(now);

  if (checkedIn) {
    // チェックアウト
    const startParts = String(checkedIn.start_time).split(':').map(Number);
    const endParts   = timeStr.split(':').map(Number);
    const actualHours = Math.round(((endParts[0]*60+endParts[1]) - (startParts[0]*60+startParts[1])) / 6) / 10;
    const job = getJob(checkedIn.job_id);
    const dailyCalc = calcDailyWage(
      job ? job.wage : 0, checkedIn.start_time, timeStr, job ? job.transport : 0
    );
    const pay = calcWithholding(dailyCalc.dailyWage, dailyCalc.transport);
    appendWorkLog({
      user_id: userId, date: today, job_id: checkedIn.job_id,
      action: 'checkout', start_time: checkedIn.start_time, end_time: timeStr,
      actual_hours: dailyCalc.hours, wage: dailyCalc.dailyWage, transport: dailyCalc.transport,
      withholding: pay.withholding, net_pay: pay.netPay, note: '',
    });
    // 勤務完了 → PayRecord 追記
    recordPay(userId, today, checkedIn.job_id, job, {
      hourly_wage: job ? Number(job.wage) : 0,
      actual_hours: dailyCalc.hours,
      wage: dailyCalc.dailyWage,
      transport: dailyCalc.transport,
      withholding: pay.withholding,
      net_pay: pay.netPay,
    });
    return JSON.stringify({ type: 'checkout', name, actual_hours: dailyCalc.hours, net_pay: pay.netPay, start_time: checkedIn.start_time });
  }

  // チェックイン
  const approvedLogs = logs.filter(r => r.action === 'approved');
  const approved = approvedLogs.find(r => getShiftStatus(userId, today, r.job_id) === 'approved');
  const cancelledApproved = approvedLogs.find(r => getShiftStatus(userId, today, r.job_id) === 'cancelled');
  if (!approved && cancelledApproved) {
    return JSON.stringify({ type: 'error', name, msg: '本日のシフトはキャンセル済みです' });
  }
  const jobId = approved ? approved.job_id : 'JOB001';
  const job = getJob(jobId);
  appendWorkLog({
    user_id: userId, date: today, job_id: jobId,
    action: 'checkin', start_time: timeStr, end_time: '',
    actual_hours: '', wage: job ? job.wage : 0, transport: job ? job.transport : 0,
    withholding: '', net_pay: '', note: '',
  });
  return JSON.stringify({ type: 'checkin', name });
}

function api_getRecruitSettings() {
  const props = PropertiesService.getScriptProperties();
  const autoEnabled = props.getProperty('AUTO_RECRUIT');
  return JSON.stringify({
    autoEnabled: autoEnabled !== 'false', // デフォルトtrue
  });
}

function api_setAutoRecruit(enabled) {
  PropertiesService.getScriptProperties().setProperty('AUTO_RECRUIT', enabled ? 'true' : 'false');
  return JSON.stringify({ ok: true, autoEnabled: !!enabled });
}

function api_manualBroadcast(msg) {
  if (!msg || !msg.trim()) return JSON.stringify({ ok: false, msg: 'メッセージが空です' });
  const count = broadcastToAll(msg.trim());
  return JSON.stringify({ ok: true, count });
}

// カレンダー用：月ごとのWorkLog件数（日付→checkout件数）
function api_getWorkLogByMonth(ym) {
  const [y, m] = ym.split('-').map(Number);
  const logs = getWorkLogByMonthIncludingArchive_(y, m).filter(r => r.action === 'checkout');
  const map = {};
  logs.forEach(r => {
    map[safeDate(r.date)] = (map[safeDate(r.date)] || 0) + 1;
  });
  return JSON.stringify(map);
}

// 管理画面から手動で応募を追加
function api_addApplied(userId, date, jobId) {
  const ok = appendApplied('applied', userId, date, jobId, 'admin', '');
  return JSON.stringify({ ok: !!ok });
}

// 募集一斉送信（職種+日付 → スタッフ1人ずつ個別URL付き）
function api_sendRecruitWithTokens(jobId, datesJson) {
  const job = getJob(jobId);
  if (!job) return JSON.stringify({ ok: false, msg: '職種が見つかりません' });

  const dates = _normalizeTokenDates_(datesJson);
  if (!dates || dates.length === 0) return JSON.stringify({ ok: false, msg: '日付を選択してください' });

  // コピーしやすいよう1行1日付（曜日あり）形式にする。replyExamplesをdatesListと同一にすることで
  // ensureRecruitMessageParts の重複チェックが効き、LINE返信用リストの二重掲載を防止する
  const datesList = dates.map(formatRecruitDateLine).join('\n');
  const replyExamples = datesList;

  const template = _getMsgTemplate_();
  const phoneRule = _getMsgConfig_('PHONE_RULE');
  const staffList = sheetToObjects(getSheet(SHEET.STAFF)).filter(s => s.line_user_id);
  let count = 0;

  // 日当は時給×勤務時間で計算する（job.wageは時給。「日当」表記に時給をそのまま出さない）
  const dailyCalc = calcDailyWage(job.wage, job.start_time, job.end_time, job.transport);

  staffList.forEach(s => {
    const token = _issueToken(s.line_user_id, 'apply', { jobId: job.job_id, dates: dates }, _ttlToEndOfDayJst_(3));
    const applyUrl = STATIC_APPLY_URL + '?token=' + token;
    let msg = template
      .replace('{{dates_list}}', datesList)
      .replace('{{start_time}}', _safeTime_(job.start_time))
      .replace('{{end_time}}', _safeTime_(job.end_time))
      .replace('{{daily_wage}}', dailyCalc.dailyWage.toLocaleString())
      .replace('{{transport}}', dailyCalc.transport.toLocaleString())
      .replace('{{apply_url}}', applyUrl)
      .replace('{{reply_examples}}', replyExamples)
      .replace('{{phone_rule}}', phoneRule);
    msg = ensureRecruitMessageParts(msg, applyUrl, replyExamples);
    try {
      if (pushMessage(s.line_user_id, msg)) count++;
      Utilities.sleep(200);
    } catch(e) {
      console.error('[Recruit] 送信失敗: ' + s.line_user_id);
    }
  });
  return JSON.stringify({ ok: true, count });
}

function api_getMsgConfig() {
  try {
    const rows = sheetToObjects(getSheet(SHEET.MSG_CONFIG));
    const result = {};
    rows.forEach(r => { if (r.key) result[r.key] = r.value; });
    return JSON.stringify(result);
  } catch(e) {
    return JSON.stringify({});
  }
}

function api_getStaffDetail(userId) {
  const staff = getStaff(userId);
  const jobMap = {};
  sheetToObjects(getSheet(SHEET.JOB)).forEach(j => { jobMap[j.job_id] = j; });

  const allApplied = getAppliedAllWithArchives_({ userId });
  const rawWork = getAllWorkLogWithArchives_({ userId })
    .map(r => ({ ...r,
      date:       safeDate(r.date),
      start_time: _safeTime_(r.start_time),
      end_time:   _safeTime_(r.end_time),
      job_title:  (jobMap[r.job_id] || {}).title || r.job_id,
    }));

  // 補正エントリを適用（corrects列がある場合のみ）
  const correctionsMap = {};
  rawWork.filter(r => r.corrects).forEach(r => {
    const ex = correctionsMap[r.corrects];
    if (!ex || toEpochMs_(r.timestamp) > toEpochMs_(ex.timestamp)) correctionsMap[r.corrects] = r;
  });
  const allWork = rawWork
    .filter(r => !r.corrects) // 補正エントリ自体は除外
    .map(r => {
      const corr = correctionsMap[r.log_id];
      if (!corr) return r;
      if (corr.action === 'deleted') return null;
      return { ...r, ...corr, job_title: (jobMap[corr.job_id] || jobMap[r.job_id] || {}).title || corr.job_id || r.job_id, _corrected: true };
    })
    .filter(Boolean);

  const allCancel  = sheetToObjects(getSheet(SHEET.CANCEL)).filter(r => r.user_id === userId);
  const allEval    = sheetToObjects(getSheet(SHEET.EVAL)).filter(r => r.user_id === userId);
  const allPay     = sheetToObjects(getSheet(SHEET.PAY_RECORD)).filter(r => r.user_id === userId)
    .map(r => ({ ...r, date: safeDate(r.date), job_title: (jobMap[r.job_id] || {}).title || r.job_id }));
  const allDoc     = sheetToObjects(getSheet(SHEET.DOC)).filter(r => r.user_id === userId);
  const allChatLog = sheetToObjects(getSheet(SHEET.CHAT_LOG)).filter(r => r.user_id === userId)
    .slice(-20)
    .map(r => ({ timestamp: String(r.timestamp || '').slice(0, 16), dir: r.direction || '', text: String(r.message || '') }));
  return JSON.stringify({ staff, allApplied, allWork, allCancel, allEval, allPay, allDoc, allChatLog });
}

function api_updateStaffMemo(userId, memo) {
  updateStaff(userId, { memo: memo || '' });
  return JSON.stringify({ ok: true });
}

function api_updateStaffJob(userId, jobId) {
  updateStaff(userId, { job_id: jobId || '' });
  return JSON.stringify({ ok: true });
}

// WorkLogに記録があるスタッフをjob_idが未設定の場合に自動登録
function api_autoAssignJobStaff(jobId) {
  if (!jobId) return JSON.stringify({ ok: false, msg: 'jobIdが不正です' });
  const logs = getAllWorkLogWithArchives_().filter(r => r.job_id === jobId);
  const userIds = [...new Set(logs.map(r => r.user_id).filter(Boolean))];
  const staffList = getAllStaff();
  const staffMap = {};
  staffList.forEach(s => { staffMap[s.line_user_id] = s; });
  let count = 0;
  userIds.forEach(uid => {
    const s = staffMap[uid];
    if (!s) return;
    if (!s.job_id) { // job_id未設定のみ更新
      updateStaff(uid, { job_id: jobId });
      count++;
    }
  });
  return JSON.stringify({ ok: true, count, total: userIds.length });
}

function api_updateStaffName(userId, name) {
  if (!userId || !name || !name.trim()) return JSON.stringify({ ok: false, msg: '名前を入力してください' });
  updateStaff(userId, { name: name.trim() });
  return JSON.stringify({ ok: true });
}

function api_updateStaff(userId, name, jobId) {
  if (!userId || !name || !name.trim()) return JSON.stringify({ ok: false, msg: '名前を入力してください' });
  updateStaff(userId, { name: name.trim(), job_id: jobId || '' });
  return JSON.stringify({ ok: true });
}

function api_deactivateStaff(userId) {
  if (!userId) return JSON.stringify({ ok: false, msg: 'userId が不正です' });
  updateStaff(userId, { active: false });
  return JSON.stringify({ ok: true });
}

function api_updateJob(jobId, wage, transport) {
  if (!jobId) return JSON.stringify({ ok: false, msg: 'jobId が不正です' });
  const sh = getSheet(SHEET.JOB);
  if (sh.getLastRow() < 2) return JSON.stringify({ ok: false, msg: 'データなし' });
  const hdr  = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const jIdx = hdr.indexOf('job_id');
  const wIdx = hdr.indexOf('wage'); const tIdx = hdr.indexOf('transport');
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][jIdx]) === String(jobId)) {
      // 既読の data[i] をコピーして書き直す（ループ内の再 getValues 不要）
      const row = data[i].slice();
      if (wIdx >= 0) row[wIdx] = Number(wage) || 0;
      if (tIdx >= 0) row[tIdx] = Number(transport) || 0;
      sh.getRange(i + 2, 1, 1, hdr.length).setValues([row]);
      return JSON.stringify({ ok: true });
    }
  }
  return JSON.stringify({ ok: false, msg: '職種が見つかりません' });
}

// =============================
// QRチェックイン API（個人QRコード用）
// =============================

// QRチェックイン機能の有効/無効取得
function api_getQrCheckinEnabled() {
  const val = PropertiesService.getScriptProperties().getProperty('QR_CHECKIN_ENABLED');
  return JSON.stringify({ enabled: val !== 'false' });
}

// QRチェックイン機能の有効/無効設定
function api_setQrCheckinEnabled(enabled) {
  const val = (enabled === true || enabled === 'true') ? 'true' : 'false';
  PropertiesService.getScriptProperties().setProperty('QR_CHECKIN_ENABLED', val);
  return JSON.stringify({ ok: true });
}

// 本日のチェックイン状態を返す
function api_getQrStatus(uid) {
  const qrEnabled = PropertiesService.getScriptProperties().getProperty('QR_CHECKIN_ENABLED');
  if (qrEnabled === 'false') return JSON.stringify({ disabled: true });
  const staff = getStaff(uid);
  const name = staff ? (staff.name || staff.display_name || '') : '';
  const today = formatDate(getJstNow());
  const qrRaw = sheetToObjects(getSheet(SHEET.WORKLOG))
    .filter(r => r.user_id === uid && safeDate(r.date) === today);
  const qrCorrMap = _buildWLCorrMap_(qrRaw);
  const logs = qrRaw
    .filter(r => !r.corrects && r.action !== 'deleted')
    .filter(r => !(qrCorrMap[r.log_id] && qrCorrMap[r.log_id].action === 'deleted'));
  const checkin  = logs.find(r => r.action === 'checkin');
  const checkout = logs.find(r => r.action === 'checkout');
  let status = 'none';
  let checkin_time = '';
  if (checkout) {
    status = 'checkout';
  } else if (checkin) {
    status = 'checkin';
    checkin_time = _safeTime_(checkin.start_time);
  }
  return JSON.stringify({ name, status, checkin_time });
}

// チェックイン / チェックアウト処理
function api_processQrCheckin(uid, action) {
  if (!uid) return JSON.stringify({ ok: false, msg: 'QRコードが無効です' });
  try {
  const qrEnabled = PropertiesService.getScriptProperties().getProperty('QR_CHECKIN_ENABLED');
  if (qrEnabled === 'false') return JSON.stringify({ ok: false, msg: 'チェックイン機能は現在停止中です' });
  const today = formatDate(getJstNow());
  const now = getJstNow();
  const timeStr = formatTime(now);
  const qrRaw2 = sheetToObjects(getSheet(SHEET.WORKLOG))
    .filter(r => r.user_id === uid && safeDate(r.date) === today);
  const qrCorrMap2 = _buildWLCorrMap_(qrRaw2);
  const logs = qrRaw2
    .filter(r => !r.corrects && r.action !== 'deleted')
    .filter(r => !(qrCorrMap2[r.log_id] && qrCorrMap2[r.log_id].action === 'deleted'));
  const checkin  = logs.find(r => r.action === 'checkin');
  const checkout = logs.find(r => r.action === 'checkout');

  if (action === 'checkin') {
    if (checkin) return JSON.stringify({ ok: false, msg: '本日は既にチェックイン済みです' });
    const latestMap = buildLatestAppliedMap_(
      getAppliedAllWithArchives_().filter(r => r.user_id === uid && safeDate(r.date) === today)
    );
    const approved = Object.values(latestMap).find(r => r.event === 'approved');
    const cancelled = Object.values(latestMap).find(r => r.event === 'cancelled');
    if (!approved && cancelled) return JSON.stringify({ ok: false, msg: '本日のシフトはキャンセル済みです' });
    const jobId = approved ? approved.job_id : 'JOB001';
    const job = getJob(jobId);
    appendWorkLog({
      user_id: uid, date: today, job_id: jobId,
      action: 'checkin', start_time: timeStr, end_time: '',
      actual_hours: '', wage: job ? job.wage : 0, transport: job ? job.transport : 0,
      withholding: '', net_pay: '', note: 'QR',
    });
    return JSON.stringify({ ok: true, time: timeStr });
  }

  if (action === 'checkout') {
    if (checkout) return JSON.stringify({ ok: false, msg: '本日は既にチェックアウト済みです' });
    if (!checkin)  return JSON.stringify({ ok: false, msg: 'チェックインがありません' });
    const startParts = String(checkin.start_time).split(':').map(Number);
    const endParts   = timeStr.split(':').map(Number);
    const actualHours = Math.round(((endParts[0]*60+endParts[1]) - (startParts[0]*60+startParts[1])) / 6) / 10;
    const job = getJob(checkin.job_id);
    const dailyCalc = calcDailyWage(
      job ? job.wage : 0, checkin.start_time, timeStr, job ? job.transport : 0
    );
    const pay = calcWithholding(dailyCalc.dailyWage, dailyCalc.transport);
    appendWorkLog({
      user_id: uid, date: today, job_id: checkin.job_id,
      action: 'checkout', start_time: checkin.start_time, end_time: timeStr,
      actual_hours: dailyCalc.hours, wage: dailyCalc.dailyWage, transport: dailyCalc.transport,
      withholding: pay.withholding, net_pay: pay.netPay, note: 'QR',
    });
    // 勤務完了 → PayRecord 追記
    recordPay(uid, today, checkin.job_id, job, {
      hourly_wage: job ? Number(job.wage) : 0,
      actual_hours: dailyCalc.hours,
      wage: dailyCalc.dailyWage,
      transport: dailyCalc.transport,
      withholding: pay.withholding,
      net_pay: pay.netPay,
    });
    return JSON.stringify({
      ok: true,
      start_time:   _safeTime_(checkin.start_time),
      end_time:     timeStr,
      actual_hours: dailyCalc.hours,
      wage:         dailyCalc.dailyWage,
      transport:    dailyCalc.transport,
      withholding:  pay.withholding,
      net_pay:      pay.netPay,
    });
  }

  return JSON.stringify({ ok: false, msg: '不明なアクション' });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 管理者が手動で勤務完了（パターンB）
function api_completeWork(userId, date, jobId) {
  try {
    const job = getJob(jobId);
    const dailyCalc = job
      ? calcDailyWage(job.wage, job.start_time, job.end_time, job.transport)
      : { hours: 0, dailyWage: 0, transport: 0, total: 0 };
    const pay = calcWithholding(dailyCalc.dailyWage, dailyCalc.transport);

    appendWorkLog({
      user_id:      userId,
      date:         date,
      job_id:       jobId,
      action:       'worked',
      start_time:   job ? _safeTime_(job.start_time) : '',
      end_time:     job ? _safeTime_(job.end_time) : '',
      actual_hours: dailyCalc.hours,
      wage:         dailyCalc.dailyWage,
      transport:    dailyCalc.transport,
      withholding:  pay.withholding,
      net_pay:      pay.netPay,
      note:         'manual',
    });

    const added = recordPay(userId, date, jobId, job, {
      hourly_wage: job ? Number(job.wage) : 0,
      actual_hours: dailyCalc.hours,
      wage: dailyCalc.dailyWage,
      transport: dailyCalc.transport,
      withholding: pay.withholding,
      net_pay: pay.netPay,
    });
    return JSON.stringify({ ok: true, added });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// 領収確認（WorkLogのnoteを更新）
function api_backupAllData() {
  try {
    const url = backupAllData();
    const name = '社内タイミー_バックアップ_' + formatDate(getJstNow());
    return JSON.stringify({ ok: true, url, name });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

function api_confirmReceipt(uid) {
  const today = formatDate(getJstNow());
  const sh = getSheet(SHEET.WORKLOG);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const uidIdx    = hdr.indexOf('user_id');
  const dateIdx   = hdr.indexOf('date');
  const actionIdx = hdr.indexOf('action');
  const noteIdx   = hdr.indexOf('note');
  if (sh.getLastRow() > 1 && noteIdx >= 0) {
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    let targetRow = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][uidIdx] === uid && String(data[i][dateIdx]) === today && data[i][actionIdx] === 'checkout') {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow >= 0) sh.getRange(targetRow, noteIdx + 1).setValue('QR_confirmed');
  }
  return JSON.stringify({ ok: true });
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { api_getDocFolderUrl };
}
