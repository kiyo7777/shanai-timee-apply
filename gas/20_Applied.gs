// 社内タイミー - 20_Applied.gs
// 応募記録（append-only）・状態集計

function appendApplied(event, userId, date, jobId, source, note) {
  const dedupeKey = [event, userId, date, jobId, source].join(':');
  if (isDuplicate(SHEET.APPLIED, dedupeKey)) {
    Logger.log('Applied重複スキップ: ' + dedupeKey);
    return false;
  }
  // source違い（form/line）の二重登録防止：同じuserId+date+jobIdで既にapplied/approvedなら追記しない
  if (event === 'applied') {
    const existing = getShiftStatus(userId, safeDate(date), jobId || 'JOB001');
    if (existing === 'applied' || existing === 'approved') {
      Logger.log('Applied source違い重複スキップ: ' + userId + ' ' + date);
      return false;
    }
  }
  const sh = getSheet(SHEET.APPLIED);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = formatDateTime(getJstNow());
  const row = objToRow(hdr, {
    timestamp:  now,
    dedupe_key: dedupeKey,
    event:      event,
    user_id:    userId,
    date:       safeDate(date),   // Date型混入防止（文字列に変換）
    job_id:     jobId || 'JOB001',
    source:     source || 'line',
    note:       note || '',
  });
  sh.appendRow(row);
  Logger.log('Applied追記: ' + event + ' ' + userId + ' ' + date);
  return true;
}

function normalizeAppliedRow_(r) {
  return {
    ...r,
    date: safeDate(r.date),
    job_id: r.job_id || 'JOB001',
  };
}

function getAppliedRowsForState_(includeArchives) {
  const rows = includeArchives && typeof getAppliedAllWithArchives_ === 'function'
    ? getAppliedAllWithArchives_()
    : getAppliedAll();
  return rows.map(normalizeAppliedRow_);
}

function buildLatestAppliedMap_(rows) {
  const latestMap = {};
  (rows || []).forEach(r0 => {
    const r = normalizeAppliedRow_(r0);
    const key = (r.user_id || '') + ':' + r.date + ':' + (r.job_id || '');
    if (!r.user_id || !r.date || !r.job_id) return;
    if (isNewerRow_(r, latestMap[key])) latestMap[key] = r;
  });
  return latestMap;
}

// 指定ユーザーの現在有効な応募（applied/approved）を取得
function getMyApplied(userId) {
  const today = formatDate(getJstNow());
  const latestMap = buildLatestAppliedMap_(
    getAppliedRowsForState_(true).filter(r => r.user_id === userId && safeDate(r.date) >= today)
  );
  return Object.values(latestMap)
    .filter(r => r.event === 'applied' || r.event === 'approved')
    .sort((a, b) => safeDate(a.date).localeCompare(safeDate(b.date)));
}

// 全件集計（カレンダー用）：{date: {job_id: {applied:[], approved:[]}}}
function getAppliedAll() {
  return sheetToObjects(getSheet(SHEET.APPLIED)).map(normalizeAppliedRow_);
}

// 日付×職種の最終状態を返す（applied > approved > cancelledの順で最新イベントを採用）
function getShiftStatus(userId, date, jobId) {
  const latestMap = buildLatestAppliedMap_(
    getAppliedRowsForState_(true).filter(r =>
      r.user_id === userId && safeDate(r.date) === date && (r.job_id || 'JOB001') === (jobId || 'JOB001')
    )
  );
  const rows = Object.values(latestMap);
  return rows.length > 0 ? rows[0].event : null;
}

function getCancellableShifts(userId, date) {
  const latestMap = buildLatestAppliedMap_(
    getAppliedRowsForState_(true).filter(r => r.user_id === userId && safeDate(r.date) === date)
  );
  return Object.values(latestMap)
    .filter(r => r.event === 'applied' || r.event === 'approved');
}

// 確定処理：approved追記 + WorkLog追記のみ（PayRecordは勤務完了時に追記）
function approveShift(userId, date, jobId, notify) {
  date = String(date).slice(0, 10); // ISO timestamp -> YYYY-MM-DD
  const status = getShiftStatus(userId, date, jobId);
  if (status === 'approved') return { ok: false, msg: '既に確定済みです' };

  // タイムスタンプ付きsourceでdedupe_key衝突を防ぐ
  // （承認→キャンセル後の再承認でも通過できるよう）
  // getShiftStatus の approved チェックがメインの二重防止ロジック
  const src = 'admin_' + formatDateTime(getJstNow()).replace(/[^0-9]/g, '').slice(0, 14);
  const added = appendApplied('approved', userId, date, jobId, src, '');
  if (!added) return { ok: false, msg: '承認に失敗しました' };

  // WorkLog に scheduled 追記
  const job = getJob(jobId);
  // 予定時間から日当を計算（時給 × 予定時間 + 交通費）
  const dailyCalc = job
    ? calcDailyWage(job.wage, job.start_time, job.end_time, job.transport)
    : { hours: 0, dailyWage: 0, transport: 0, total: 0 };
  const pay = calcWithholding(dailyCalc.dailyWage, dailyCalc.transport);
  const logId = uuid();
  appendWorkLog({
    log_id:    logId,
    user_id:   userId,
    date:      date,
    job_id:    jobId,
    action:    'approved',
    start_time:   job ? _safeTime_(job.start_time) : '',
    end_time:     job ? _safeTime_(job.end_time) : '',
    actual_hours: dailyCalc.hours,
    wage:         dailyCalc.dailyWage,
    transport:    dailyCalc.transport,
    withholding:  pay.withholding,
    net_pay:      pay.netPay,
    note:         '',
  });

  // LINEに通知（notify=false の場合はスキップ）
  if (notify !== false) {
    const dateJa = formatDateJa(date);
    const title = job ? job.title : jobId;
    pushMessage(userId,
      `✅ シフト確定しました！\n${dateJa} ${title}\n${job ? _safeTime_(job.start_time) + '〜' + _safeTime_(job.end_time) : ''}\nよろしくお願いします 🙏`
    );
  }

  return { ok: true, msg: '確定しました' };
}

// キャンセル処理：cancelled追記 + CancelLog追記
function cancelShift(userId, date, jobId, reason, source) {
  date = String(date).slice(0, 10); // ISO timestamp -> YYYY-MM-DD（approveShift と対称）
  // 既にキャンセル済みなら早期リターン（approveShift と対称）
  const status = getShiftStatus(userId, date, jobId);
  if (status === 'cancelled') return { ok: true, msg: 'キャンセル済みです' };

  // タイムスタンプ付き source で dedupe_key 衝突を防ぐ
  // （承認→キャンセル後の再キャンセルでも Applied に追記できるよう approveShift と同方式）
  const ts = formatDateTime(getJstNow()).replace(/[^0-9]/g, '').slice(0, 14);
  const src = (source || 'admin') + '_' + ts;
  const added = appendApplied('cancelled', userId, date, jobId, src, reason || '');
  if (!added) return { ok: true, msg: 'キャンセル済み（重複スキップ）' };

  // キャンセル種別判定
  const today = formatDate(getJstNow());
  const cancelType = (date === today) ? 'same_day' : 'before_day';

  const sh = getSheet(SHEET.CANCEL);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = formatDateTime(getJstNow());
  sh.appendRow(objToRow(hdr, {
    timestamp:   now,
    user_id:     userId,
    date:        date,
    job_id:      jobId,
    cancel_type: cancelType,
    reason:      reason || '',
    notified_at: now,
  }));

  // 確定後キャンセル: total_cancel をインクリメント（当日なら same_day_cancel も）
  if (status === 'approved') {
    const staff = getStaff(userId);
    if (staff) {
      updateStaff(userId, {
        total_cancel:    (Number(staff.total_cancel) || 0) + 1,
        ...(cancelType === 'same_day' ? { same_day_cancel: (Number(staff.same_day_cancel) || 0) + 1 } : {}),
      });
    }
    if (cancelType === 'same_day') {
      try {
        pushMessage(SECRETS.ADMIN_USER_ID,
          `⚠️ 当日キャンセル通知\n${userId}\n日付: ${date}\n職種: ${jobId}`
        );
      } catch(e) { /* ADMIN_USER_ID 未設定の場合は無視 */ }
    }
  }

  return { ok: true, msg: 'キャンセルしました' };
}

// まとめ確定（approvals = [{userId, date, jobId}, ...]）
function approveShifts(approvals) {
  var count = 0;
  (approvals || []).forEach(function(a) {
    try {
      var r = approveShift(a.userId, a.date, a.jobId);
      if (r.ok) count++;
    } catch(e) {
      console.error('[approveShifts] ' + e.message);
    }
    Utilities.sleep(200);
  });
  return { ok: true, approved: count };
}

// 職種マスタ取得ヘルパー
function getJob(jobId) {
  const jobs = sheetToObjects(getSheet(SHEET.JOB));
  return jobs.find(j => j.job_id === jobId) || null;
}
