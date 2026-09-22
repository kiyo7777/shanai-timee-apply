// 社内タイミー - 40_Payroll.gs
// 給与計算・源泉徴収
// 法的根拠：所得税法第185条・日額表丙欄
// 日雇い労働者（労働基準法第21条）

function calcWithholding(wage, transport) {
  const w = Number(wage) || 0;
  const t = Math.max(0, Number(transport) || 0);
  const taxable = w; // 課税対象は日当のみ（交通費は非課税のため除外しない）
  if (taxable <= 0) return { taxableWage: 0, withholding: 0, netPay: w };

  let withholding = 0;
  try {
    const rows = sheetToObjects(getSheet(SHEET.TAX));
    const row = rows.find(r =>
      taxable >= Number(r.wage_from) && taxable <= Number(r.wage_to)
    );
    if (row) {
      const wt = Number(row.withholding);
      // TAX行の値が不正（空・#NUM!等）なら近似フォールバック。NaNをそのまま採用しない。
      withholding = Number.isFinite(wt) ? wt : (taxable >= 13000 ? Math.floor(taxable * 0.051) : 0);
    } else if (taxable >= 13000) {
      withholding = Math.floor(taxable * 0.051);
    }
  } catch(e) {
    // フォールバック：近似計算
    if (taxable < 9300)       withholding = 0;
    else if (taxable < 10000) withholding = 200;
    else if (taxable < 11000) withholding = 300;
    else                      withholding = Math.floor(taxable * 0.051);
  }

  return {
    taxableWage: taxable,
    withholding:  withholding,
    netPay:       w + t - withholding,
  };
}

// 月次給与一覧（WorkLog × Staff × JobMaster 結合）
function getMonthlyPayroll(year, month) {
  const PAYABLE = ['checkout', 'approved', 'manual', 'worked'];
  const filtered = getWorkLogByMonth(year, month)
    .filter(r => PAYABLE.includes(r.action));
  // checkout がある日は他を除外（二重カウント防止）
  const checkoutKeys = new Set(
    filtered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const wlogs = filtered.filter(r =>
    r.action === 'checkout' || !checkoutKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s; });
  const jobMap = {};
  sheetToObjects(getSheet(SHEET.JOB)).forEach(j => { jobMap[j.job_id] = j; });

  return wlogs.map(r => ({
    ...r,
    staff_name:    (staffMap[r.user_id] || {}).name || (staffMap[r.user_id] || {}).display_name || r.user_id,
    job_title:     (jobMap[r.job_id] || {}).title || r.job_id,
  }));
}

// 年次集計 → PaySummaryにupsert
function calcYearlySummary(year) {
  // アーカイブシートも含めて読み込む
  const allYearRows = getAllWorkLogWithArchives_({})
    .filter(r => safeDate(r.date).startsWith(String(year)));
  const yearCorrMap = _buildWLCorrMap_(allYearRows);
  const PAYABLE = ['checkout', 'approved', 'manual', 'worked'];
  const filtered = allYearRows
    .filter(r => !r.corrects && PAYABLE.includes(r.action))
    .filter(r => !(yearCorrMap[r.log_id] && yearCorrMap[r.log_id].action === 'deleted'));
  // checkout がある日は他を除外（二重カウント防止）
  const checkoutKeys = new Set(
    filtered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const wlogs = filtered.filter(r =>
    r.action === 'checkout' || !checkoutKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );

  const summary = {};
  wlogs.forEach(r => {
    const uid = r.user_id;
    if (!uid) return;
    if (!summary[uid]) {
      summary[uid] = {
        user_id: uid, year: year,
        total_wage: 0, total_transport: 0,
        total_withholding: 0, total_net_pay: 0,
        work_days: new Set(),
      };
    }
    summary[uid].total_wage      += Number(r.wage) || 0;
    summary[uid].total_transport += Number(r.transport) || 0;
    summary[uid].total_withholding += Number(r.withholding) || 0;
    summary[uid].total_net_pay   += Number(r.net_pay) || 0;
    summary[uid].work_days.add(safeDate(r.date));
  });

  const sh = getSheet(SHEET.PAY_SUMM);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = formatDateTime(getJstNow());

  // ループ外で1回だけ読み込む（ループ内 getValues() 禁止）
  const uidIdx  = hdr.indexOf('user_id');
  const yearIdx = hdr.indexOf('year');
  let existingRows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];

  const reportSubmittedIdx = hdr.indexOf('report_submitted');

  Object.values(summary).forEach(s => {
    const workDays = s.work_days.size;
    s.work_days = workDays;
    s.needs_report = s.total_wage >= 500000;
    s.updated_at = now;

    // upsert（existingRows はループ前に取得済み）
    const found = existingRows.findIndex(r => r[uidIdx] === s.user_id && Number(r[yearIdx]) === year);
    if (found >= 0) {
      // 既存行の report_submitted を引き継ぐ（再集計で提出済みフラグを上書きしない）
      const existingSubmitted = reportSubmittedIdx >= 0 ? existingRows[found][reportSubmittedIdx] : false;
      s.report_submitted = isPaid_(existingSubmitted);
      sh.getRange(found + 2, 1, 1, hdr.length).setValues([objToRow(hdr, s)]);
    } else {
      s.report_submitted = false;
      sh.appendRow(objToRow(hdr, s));
    }
  });

  Logger.log('calcYearlySummary完了: ' + year);
  return Object.values(summary);
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { calcWithholding };
}
