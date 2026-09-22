// 社内タイミー - 50_PayRecord.gs
// 支払台帳（PayRecord / PaySummary）

function recordPay(userId, date, jobId, job, actual) {
  const j = job || getJob(jobId);
  const a = actual || {};
  const hasActualWage = Object.prototype.hasOwnProperty.call(a, 'wage');
  const hasActualHours = Object.prototype.hasOwnProperty.call(a, 'actual_hours');
  const hourlyWage = Object.prototype.hasOwnProperty.call(a, 'hourly_wage')
    ? Number(a.hourly_wage)
    : (j ? Number(j.wage) : 0);
  const calc = (hasActualWage && hasActualHours)
    ? {
        hours: Number(a.actual_hours) || 0,
        dailyWage: Number(a.wage) || 0,
        transport: Object.prototype.hasOwnProperty.call(a, 'transport') ? (Number(a.transport) || 0) : (j ? Number(j.transport) : 0),
        total: (Number(a.wage) || 0) + (Object.prototype.hasOwnProperty.call(a, 'transport') ? (Number(a.transport) || 0) : (j ? Number(j.transport) : 0)),
      }
    : (j
      ? calcDailyWage(hourlyWage, j.start_time, j.end_time, j.transport)
      : { hours: 0, dailyWage: 0, transport: 0, total: 0 });
  const pay0 = calcWithholding(calc.dailyWage, calc.transport);
  const pay = {
    withholding: Object.prototype.hasOwnProperty.call(a, 'withholding') ? (Number(a.withholding) || 0) : pay0.withholding,
    netPay: Object.prototype.hasOwnProperty.call(a, 'net_pay') ? (Number(a.net_pay) || 0) : pay0.netPay,
  };
  const now = formatDateTime(getJstNow());

  const sh = getSheet(SHEET.PAY_RECORD);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  // 重複チェック（同一 user_id + date + job_id）
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const uidIdx  = hdr.indexOf('user_id');
    const dateIdx = hdr.indexOf('date');
    const jobIdx  = hdr.indexOf('job_id');
    const dateStr = safeDate(date);
    const dup = rows.find(r =>
      r[uidIdx] === userId && safeDate(r[dateIdx]) === dateStr && r[jobIdx] === jobId
    );
    if (dup) {
      Logger.log('PayRecord重複スキップ: ' + userId + ' ' + date);
      return false;
    }
  }

  sh.appendRow(objToRow(hdr, {
    pay_id:          uuid(),
    user_id:         userId,
    date:            safeDate(date),
    job_id:          jobId,
    hourly_wage:     hourlyWage,
    actual_hours:    calc.hours,
    wage:            calc.dailyWage,
    transport:       calc.transport,
    withholding:     pay.withholding,
    net_pay:         pay.netPay,
    paid:            false,
    paid_at:         '',
    payment_method:  'cash',
    receipt_url:     '',
    note:            '',
    created_at:      now,
  }));
  return true;
}

function markAsPaid(payId) {
  const sh = getSheet(SHEET.PAY_RECORD);
  if (sh.getLastRow() < 2) return false;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const pidIdx  = hdr.indexOf('pay_id');
  const paidIdx = hdr.indexOf('paid');
  const patIdx  = hdr.indexOf('paid_at');
  const now = formatDateTime(getJstNow());

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][pidIdx] === payId) {
      if (paidIdx >= 0 && isPaid_(rows[i][paidIdx])) return true;
      sh.getRange(i + 2, paidIdx + 1).setValue('true');
      sh.getRange(i + 2, patIdx + 1).setValue(now);
      return true;
    }
  }
  return false;
}

// 未来の PayRecord（確定時に誤追記されたもの）を削除する一回限りの修復関数
// GASエディタから手動実行すること
function cleanFuturePayRecords() {
  const sh = getSheet(SHEET.PAY_RECORD);
  const hm = buildHeaderMap(sh);
  const today = formatDate(getJstNow());
  if (sh.getLastRow() < 2) { Logger.log('PayRecord: データなし'); return; }
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const toDelete = [];
  rows.forEach((row, i) => {
    const date = safeDate(row[hm['date']]);
    if (date > today && !isPaid_(row[hm['paid']])) {
      toDelete.push(i + 2);
    }
  });
  toDelete.reverse().forEach(n => sh.deleteRow(n)); // 後ろから削除（行番号ずれ防止）
  Logger.log('未来のPayRecord削除: ' + toDelete.length + '件');
}

function getPaySummary(year, month) {
  const prefix = year + '-' + String(month).padStart(2, '0');
  const rows = sheetToObjects(getSheet(SHEET.PAY_RECORD))
    .filter(r => safeDate(r.date).startsWith(prefix));
  const total_wage        = rows.reduce((s, r) => s + (Number(r.wage) || 0), 0);
  const total_transport   = rows.reduce((s, r) => s + (Number(r.transport) || 0), 0);
  const total_withholding = rows.reduce((s, r) => s + (Number(r.withholding) || 0), 0);
  const total_net_pay     = rows.reduce((s, r) => s + (Number(r.net_pay) || 0), 0);
  const unpaid_count      = rows.filter(r => !isPaid_(r.paid)).length;
  return { rows, total_wage, total_transport, total_withholding, total_net_pay, unpaid_count };
}

// 月別支払一覧（Staff・JobMaster 結合）
function getPayRecordsByMonth(year, month) {
  const summary = getPaySummary(year, month);
  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s; });
  const jobMap = {};
  sheetToObjects(getSheet(SHEET.JOB)).forEach(j => { jobMap[j.job_id] = j; });

  const rows = summary.rows.map(r => ({
    ...r,
    date:       safeDate(r.date),
    staff_name: (staffMap[r.user_id] || {}).name || (staffMap[r.user_id] || {}).display_name || r.user_id,
    job_title:  (jobMap[r.job_id] || {}).title  || r.job_id,
    start_time: _safeTime_((jobMap[r.job_id] || {}).start_time) || '',
    end_time:   _safeTime_((jobMap[r.job_id] || {}).end_time)   || '',
  }));
  return { ...summary, rows };
}

// 対象月のPayRecord行が全件支払済み（paid=true）かどうかを判定する純粋関数
// payRecords: 対象月に絞り込み済みの PayRecord 行オブジェクト配列
// 対象月に1件もPayRecordが無い場合（＝その月に勤務実績のあるスタッフがいない）は false を返す
// （呼び出し側で「対象外」として扱う想定。checkAndGenerateMonthlyPayslip 参照）
function _isMonthFullyPaid_(payRecords) {
  if (!payRecords || !payRecords.length) return false;
  return payRecords.every(r => isPaid_(r.paid));
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { _isMonthFullyPaid_ };
}
