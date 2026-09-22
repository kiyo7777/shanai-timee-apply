// 社内タイミー - 30_WorkLog.gs
// 勤怠記録（append-only）

function appendWorkLog(fields) {
  const sh = getSheet(SHEET.WORKLOG);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = formatDateTime(getJstNow());
  const row = objToRow(hdr, Object.assign({ timestamp: now, log_id: uuid() }, fields));
  sh.appendRow(row);
  const lidIdx = hdr.indexOf('log_id');
  return lidIdx >= 0 ? row[lidIdx] : '';
}

// 当日のWorkLogを取得（deleted・補正削除を除外）
function getTodayWorkLog(userId) {
  const today = formatDate(getJstNow());
  // 補正行が翌日以降に追記される場合があるため、全期間行で corrMap を構築してから日付フィルタを適用
  const allUserRows = sheetToObjects(getSheet(SHEET.WORKLOG))
    .filter(r => r.user_id === userId);
  const corrMap = _buildWLCorrMap_(allUserRows);
  return allUserRows
    .filter(r => safeDate(r.date) === today)
    .filter(r => !r.corrects && r.action !== 'deleted')
    .filter(r => !(corrMap[r.log_id] && corrMap[r.log_id].action === 'deleted'))
    .sort((a, b) => toEpochMs_(b.timestamp) - toEpochMs_(a.timestamp));
}

// チェックイン処理
function processCheckin(userId, replyToken) {
  const today = formatDate(getJstNow());
  const logs = getTodayWorkLog(userId);

  // approved レコード確認
  const approvedLogs = logs.filter(r => r.action === 'approved');
  const approved = approvedLogs.find(r => getShiftStatus(userId, today, r.job_id) === 'approved');
  const cancelledApproved = approvedLogs.find(r => getShiftStatus(userId, today, r.job_id) === 'cancelled');
  const checkedIn = logs.find(r => r.action === 'checkin');
  const checkedOut = logs.find(r => r.action === 'checkout');

  function replyAndLog(msg) {
    replyMessage(replyToken, msg);
    _appendChatLog(userId, 'out', msg);
  }

  if (checkedOut) {
    replyAndLog('本日の勤務記録は完了しています ✅');
    return;
  }

  if (checkedIn) {
    // チェックアウト処理
    processCheckout(userId, replyToken);
    return;
  }

  if (!approved && cancelledApproved) {
    replyAndLog('本日のシフトはキャンセル済みです。必要な場合は担当者へ確認してください。');
    return;
  }

  // チェックイン
  const now = getJstNow();
  const startTime = formatTime(now);
  const staffForJob = getStaff(userId);
  const jobId = approved ? approved.job_id : ((staffForJob && staffForJob.job_id) ? staffForJob.job_id : 'JOB001');
  const job = getJob(jobId);

  appendWorkLog({
    user_id:    userId,
    date:       today,
    job_id:     jobId,
    action:     'checkin',
    start_time: startTime,
    end_time:   '',
    actual_hours: '',
    wage:       job ? job.wage : 0,
    transport:  job ? job.transport : 0,
    withholding: '',
    net_pay:    '',
    note:       '',
  });

  const name = (staffForJob && staffForJob.name) ? staffForJob.name : (staffForJob ? staffForJob.display_name : 'スタッフ');
  replyAndLog(`おはようございます！${name}さん ☀️\nチェックイン記録しました\n⏰ ${startTime}`);
}

// チェックアウト処理
function processCheckout(userId, replyToken) {
  const today = formatDate(getJstNow());
  const logs = getTodayWorkLog(userId);
  const checkedIn = logs.find(r => r.action === 'checkin');
  const checkedOut = logs.find(r => r.action === 'checkout');

  function replyAndLog(msg) {
    replyMessage(replyToken, msg);
    _appendChatLog(userId, 'out', msg);
  }

  if (checkedOut) {
    replyAndLog('本日の勤務記録は完了しています ✅');
    return;
  }

  if (!checkedIn) {
    replyAndLog('チェックインの記録が見つかりません。\nまずチェックインしてください。');
    return;
  }

  const now = getJstNow();
  const endTime = formatTime(now);

  const job = getJob(checkedIn.job_id);
  const dailyCalc = calcDailyWage(
    job ? job.wage : 0, checkedIn.start_time, endTime, job ? job.transport : 0
  );
  const payCalc = calcWithholding(dailyCalc.dailyWage, dailyCalc.transport);

  appendWorkLog({
    user_id:      userId,
    date:         today,
    job_id:       checkedIn.job_id,
    action:       'checkout',
    start_time:   checkedIn.start_time,
    end_time:     endTime,
    actual_hours: dailyCalc.hours,
    wage:         dailyCalc.dailyWage,
    transport:    dailyCalc.transport,
    withholding:  payCalc.withholding,
    net_pay:      payCalc.netPay,
    note:         '',
  });

  // 勤務完了 → PayRecord に追記（実績値で記録）
  const payAdded = recordPay(userId, today, checkedIn.job_id, job, {
    hourly_wage: job ? Number(job.wage) : 0,
    actual_hours: dailyCalc.hours,
    wage: dailyCalc.dailyWage,
    transport: dailyCalc.transport,
    withholding: payCalc.withholding,
    net_pay: payCalc.netPay,
  });
  if (!payAdded) Logger.log('[processCheckout] PayRecord重複スキップ: ' + userId + ' ' + today);

  replyAndLog(
    `おつかれさまでした！ 🙏\n⏰ ${_safeTime_(checkedIn.start_time)} 〜 ${endTime}\n実働: ${dailyCalc.hours}時間\n\n日当: ¥${dailyCalc.dailyWage.toLocaleString()}\n交通費: ¥${dailyCalc.transport.toLocaleString()}\n源泉: ¥${payCalc.withholding.toLocaleString()}\n手取り: ¥${(dailyCalc.dailyWage - payCalc.withholding + dailyCalc.transport).toLocaleString()}`
  );
}

// corrects 列のある補正行から「削除済みlog_id」マップを構築
function _buildWLCorrMap_(rows) {
  const corrMap = {};
  rows.filter(r => r.corrects).forEach(r => {
    const ex = corrMap[r.corrects];
    if (!ex || toEpochMs_(r.timestamp) > toEpochMs_(ex.timestamp)) corrMap[r.corrects] = r;
  });
  return corrMap;
}

// 月別勤怠取得（deleted 論理削除行・補正削除を除外）
function getWorkLogByMonth(year, month) {
  const prefix = year + '-' + String(month).padStart(2, '0');
  // 補正行が月をまたぐ場合があるため、全行でcorrMapを構築してから月フィルタを適用する
  const allRows = sheetToObjects(getSheet(SHEET.WORKLOG));
  const corrMap = _buildWLCorrMap_(allRows);
  return allRows
    .filter(r => safeDate(r.date).startsWith(prefix))
    .filter(r => !r.corrects)
    .filter(r => r.action !== 'deleted')
    .filter(r => !(corrMap[r.log_id] && corrMap[r.log_id].action === 'deleted'));
}

// アーカイブシートからも含めた勤怠取得
function getWorkLogData(yearMonth, fromArchive) {
  const sheetName = fromArchive ? ('WorkLog_' + yearMonth.replace('-', '_')) : SHEET.WORKLOG;
  try {
    const sh = getSheet(sheetName);
    const rows = sheetToObjects(sh);
    if (fromArchive) return rows;
    // 補正行が月をまたぐ場合があるため、全行でcorrMapを構築してから月フィルタを適用する
    const corrMap = _buildWLCorrMap_(rows);
    return rows
      .filter(r => safeDate(r.date).startsWith(yearMonth))
      .filter(r => !r.corrects && r.action !== 'deleted')
      .filter(r => !(corrMap[r.log_id] && corrMap[r.log_id].action === 'deleted'));
  } catch(e) {
    return [];
  }
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { _buildWLCorrMap_ };
}
