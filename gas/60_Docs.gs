// 社内タイミー - 60_Docs.gs
// 書類PDF生成・Drive保存
// 法的根拠：
//   賃金台帳：労働基準法第108条（5年間保存義務）
//   支払調書：所得税法第225条（年間50万円超は提出義務）

function getDriveFolder(name) {
  const rootName = '社内タイミー';
  const roots = DriveApp.getRootFolder().getFoldersByName(rootName);
  const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
  const subs = root.getFoldersByName(name);
  return subs.hasNext() ? subs.next() : root.createFolder(name);
}

function _recordDocLog(userId, docType, targetPeriod, fileUrl) {
  const sh = getSheet(SHEET.DOC);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const docId = uuid();
  sh.appendRow(objToRow(hdr, {
    doc_id:        docId,
    user_id:       userId,
    doc_type:      docType,
    target_period: targetPeriod,
    file_url:      fileUrl,
    created_at:    formatDateTime(getJstNow()),
    created_by:    'admin',
  }));
  return docId;
}

// Google DocsでHTMLをPDFに変換してDriveに保存
function _htmlToPdf(html, fileName, folder) {
  const blob = Utilities.newBlob(html, MimeType.HTML, fileName + '.html');
  let doc;
  try {
    doc = Drive.Files.insert(
      { title: fileName, mimeType: MimeType.GOOGLE_DOCS },
      blob
    );
  } catch(e) {
    throw new Error('Drive.Files.insert 失敗: ' + e.message);
  }
  try {
    const pdfBlob = DriveApp.getFileById(doc.id).getAs(MimeType.PDF);
    pdfBlob.setName(fileName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    return pdfFile.getUrl();
  } finally {
    // 成功・失敗どちらの場合も中間Docsを削除
    try { DriveApp.getFileById(doc.id).setTrashed(true); } catch(_) {}
  }
}

// 領収書PDF生成
function generateReceipt(userId, date) {
  const allLogs = sheetToObjects(getSheet(SHEET.WORKLOG))
    .filter(r => r.user_id === userId && safeDate(r.date) === date);
  const corrMap = _buildWLCorrMap_(allLogs);
  const logs = allLogs
    .filter(r => !r.corrects && r.action !== 'deleted')
    .filter(r => !(corrMap[r.log_id] && corrMap[r.log_id].action === 'deleted'));
  const log = logs.find(r => r.action === 'checkout') || logs.find(r => r.action === 'approved');
  if (!log) return { ok: false, msg: '勤怠記録が見つかりません' };

  const staff = getStaff(userId);
  const job   = getJob(log.job_id);
  const company = {
    name: _safeGetProp('COMPANY_NAME'),
    address: _safeGetProp('COMPANY_ADDRESS'),
    tel: _safeGetProp('COMPANY_TEL'),
  };

  const html = _buildReceiptHtml({
    staff, log, job, company, date,
    dateJa: formatDateJa(date),
  });

  const folder = getDriveFolder('領収書');
  const fileName = `領収書_${(staff ? (staff.name || staff.display_name) : userId)}_${date}`;
  try {
    const url = _htmlToPdf(html, fileName, folder);
    const docId = _recordDocLog(userId, 'receipt', date, url);
    pushMessage(userId, `📄 領収書を発行しました\n${url}`);
    return { ok: true, url };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function _buildReceiptHtml(d) {
  const name = d.staff ? (d.staff.name || d.staff.display_name || d.staff.line_user_id) : '';
  const title = d.job ? d.job.title : d.log.job_id;
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    body{font-family:sans-serif;margin:40px;} h1{font-size:24px;text-align:center;}
    table{width:100%;border-collapse:collapse;margin-top:20px;}
    td,th{border:1px solid #ccc;padding:8px;} th{background:#f5f5f5;}
    .total{font-weight:bold;font-size:18px;} .company{text-align:right;margin-top:30px;}
  </style></head><body>
  <h1>領 収 書</h1>
  <p>${name} 様</p>
  <table>
    <tr><th>勤務日</th><td>${d.dateJa}</td></tr>
    <tr><th>職種</th><td>${title}</td></tr>
    <tr><th>勤務時間</th><td>${_docFmtTime(d.log.start_time)} 〜 ${_docFmtTime(d.log.end_time)}</td></tr>
    <tr><th>日当</th><td>¥${Number(d.log.wage || 0).toLocaleString()}</td></tr>
    <tr><th>交通費</th><td>¥${Number(d.log.transport || 0).toLocaleString()}</td></tr>
    <tr><th>源泉徴収額</th><td>¥${Number(d.log.withholding || 0).toLocaleString()}</td></tr>
    <tr><th class="total">お支払金額（手取り）</th><td class="total">¥${Number(d.log.net_pay || 0).toLocaleString()}</td></tr>
  </table>
  <div class="company">
    <p>${d.company.name}</p>
    <p>${d.company.address}</p>
    <p>TEL: ${d.company.tel}</p>
    <p>発行日: ${formatDate(getJstNow())}</p>
  </div>
  </body></html>`;
}

// 賃金台帳PDF生成
function generatePayslip(userId, yearMonth) {
  const [y, m] = yearMonth.split('-').map(Number);
  const _allPayslipLogs = getWorkLogByMonthIncludingArchive_(y, m)
    .filter(r => r.user_id === userId);
  const _payslipCorrMap = _buildWLCorrMap_(_allPayslipLogs);
  const _payslipFiltered = _allPayslipLogs
    .filter(r => (r.action === 'checkout' || r.action === 'approved') && !r.corrects)
    .filter(r => !(_payslipCorrMap[r.log_id] && _payslipCorrMap[r.log_id].action === 'deleted'));
  const _payslipCOKeys = new Set(
    _payslipFiltered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const logs = _payslipFiltered.filter(r =>
    r.action === 'checkout' || !_payslipCOKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const staff = getStaff(userId);
  const name = staff ? (staff.name || staff.display_name || userId) : userId;
  const company = {
    name: _safeGetProp('COMPANY_NAME'),
  };

  // ループ前にジョブマスタを一括取得（ループ内での繰り返しシートアクセスを防ぐ）
  const jobMap = {};
  sheetToObjects(getSheet(SHEET.JOB)).forEach(j => { jobMap[j.job_id] = j; });

  const rows = logs.map(r => {
    const job = jobMap[r.job_id] || null;
    return `<tr>
      <td>${formatDateJa(r.date)}</td>
      <td>${job ? job.title : r.job_id}</td>
      <td>${_docFmtTime(r.start_time)} 〜 ${_docFmtTime(r.end_time)}</td>
      <td>${r.actual_hours || ''}</td>
      <td>¥${Number(r.wage || 0).toLocaleString()}</td>
      <td>¥${Number(r.transport || 0).toLocaleString()}</td>
      <td>¥${Number(r.withholding || 0).toLocaleString()}</td>
      <td>¥${Number(r.net_pay || 0).toLocaleString()}</td>
    </tr>`;
  }).join('');

  const totalWage = logs.reduce((s, r) => s + Number(r.wage || 0), 0);
  const totalTransport = logs.reduce((s, r) => s + Number(r.transport || 0), 0);
  const totalWithholding = logs.reduce((s, r) => s + Number(r.withholding || 0), 0);
  const totalNetPay = logs.reduce((s, r) => s + Number(r.net_pay || 0), 0);

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    body{font-family:sans-serif;margin:30px;} h1{text-align:center;}
    table{width:100%;border-collapse:collapse;font-size:12px;}
    td,th{border:1px solid #ccc;padding:6px;} th{background:#f0f0f0;}
    .total-row{font-weight:bold;background:#ffe;}
  </style></head><body>
  <h1>賃金台帳</h1>
  <p><b>${company.name}</b> &nbsp; 対象期間: ${yearMonth} &nbsp; 氏名: ${name}</p>
  <table>
    <tr><th>日付</th><th>職種</th><th>勤務時間</th><th>実働(h)</th>
        <th>日当</th><th>交通費</th><th>源泉</th><th>手取り</th></tr>
    ${rows}
    <tr class="total-row">
      <td colspan="4">合計</td>
      <td>¥${totalWage.toLocaleString()}</td>
      <td>¥${totalTransport.toLocaleString()}</td>
      <td>¥${totalWithholding.toLocaleString()}</td>
      <td>¥${totalNetPay.toLocaleString()}</td>
    </tr>
  </table>
  <p style="font-size:11px;color:#888;margin-top:20px;">
    ※ 労働基準法第108条に基づき5年間保存
  </p>
  </body></html>`;

  // 保存先: D-102固定フォルダ（📁ボタン=api_getDocFolderUrl、一括生成=generateBulkPayslipCombined L301と同一）
  const folder = DriveApp.getFolderById('198ubbF5s16KukcE0H67fy8gdh8Jw-7k5');
  const fileName = `賃金台帳_${name}_${yearMonth}`;
  try {
    const url = _htmlToPdf(html, fileName, folder);
    _recordDocLog(userId, 'payslip', yearMonth, url);
    return { ok: true, url };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// 日付を "YYYY-MM-DD" 文字列に統一
function _docFmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) return formatDate(v);
  return String(v).slice(0, 10);
}

// 時刻を "HH:MM" 文字列に統一（Spreadsheet は Dec 30 1899 の Date で返る）
// getHours() はローカル時刻依存のため _safeTime_ / formatTime を経由して JST 補正済み UTC ゲッターを使う
function _docFmtTime(v) {
  if (!v) return '';
  if (v instanceof Date) return _safeTime_(v);
  return String(v).slice(0,5);
}

// 全スタッフ賃金台帳を1ファイルで生成（DocumentApp でページ区切り確実化）
function generateBulkPayslipCombined(yearMonth) {
  const [y, m] = yearMonth.split('-').map(Number);
  const _rawBulkLogs = getWorkLogByMonthIncludingArchive_(y, m);
  const _bulkCorrMap = _buildWLCorrMap_(_rawBulkLogs);
  const _bulkFiltered = _rawBulkLogs
    .filter(r => (r.action === 'checkout' || r.action === 'approved') && !r.corrects)
    .filter(r => !(_bulkCorrMap[r.log_id] && _bulkCorrMap[r.log_id].action === 'deleted'));
  const _bulkCOKeys = new Set(
    _bulkFiltered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const allLogs = _bulkFiltered.filter(r =>
    r.action === 'checkout' || !_bulkCOKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const userIds = [...new Set(allLogs.map(r => r.user_id).filter(Boolean))];
  if (!userIds.length) return { ok: false, msg: '対象月に勤務記録がありません' };

  const staffMap = {};
  getAllStaff().forEach(s => { staffMap[s.line_user_id] = s; });
  // ループ前にジョブマスタを一括取得（ループ内での繰り返しシートアクセスを防ぐ）
  const jobMap = {};
  sheetToObjects(getSheet(SHEET.JOB)).forEach(j => { jobMap[j.job_id] = j; });
  const companyName = _safeGetProp('COMPANY_NAME');
  const fileName = `賃金台帳_全スタッフ_${yearMonth}`;

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();
  body.clear();

  userIds.forEach((uid, idx) => {
    const logs = allLogs
      .filter(r => r.user_id === uid)
      .sort((a, b) => safeDate(a.date) > safeDate(b.date) ? 1 : safeDate(a.date) < safeDate(b.date) ? -1 : 0);
    const s = staffMap[uid] || {};
    const name = s.name || s.display_name || uid;

    const totalWage = logs.reduce((acc, r) => acc + Number(r.wage || 0), 0);
    const totalTransport = logs.reduce((acc, r) => acc + Number(r.transport || 0), 0);
    const totalWithholding = logs.reduce((acc, r) => acc + Number(r.withholding || 0), 0);
    const totalNetPay = logs.reduce((acc, r) => acc + Number(r.net_pay || 0), 0);

    body.appendParagraph('賃金台帳')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph(`${companyName}　対象期間: ${yearMonth}　氏名: ${name}`)
      .setFontSize(11);

    const headers = ['日付', '職種', '勤務時間', '実働(h)', '日当', '交通費', '源泉', '手取り'];
    const tableData = [headers];
    logs.forEach(r => {
      const job = jobMap[r.job_id] || null;
      tableData.push([
        formatDateJa(r.date),
        job ? job.title : r.job_id,
        `${_docFmtTime(r.start_time)} 〜 ${_docFmtTime(r.end_time)}`,
        String(r.actual_hours || ''),
        `¥${Number(r.wage || 0).toLocaleString()}`,
        `¥${Number(r.transport || 0).toLocaleString()}`,
        `¥${Number(r.withholding || 0).toLocaleString()}`,
        `¥${Number(r.net_pay || 0).toLocaleString()}`,
      ]);
    });
    tableData.push([
      '合計', '', '', '',
      `¥${totalWage.toLocaleString()}`,
      `¥${totalTransport.toLocaleString()}`,
      `¥${totalWithholding.toLocaleString()}`,
      `¥${totalNetPay.toLocaleString()}`,
    ]);
    body.appendTable(tableData);

    body.appendParagraph('※ 労働基準法第108条に基づき5年間保存').setFontSize(9);

    if (idx < userIds.length - 1) {
      body.appendPageBreak();
    }
  });

  doc.saveAndClose();

  try {
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs(MimeType.PDF);
    pdfBlob.setName(fileName + '.pdf');
    // 保存先: D-102固定フォルダ（管理画面の📁リンク=api_getDocFolderUrlとは別系統。2026-09-18に1M4qi...へ変更を検討したが、
    // ユーザー判断により198ubb...を正として維持することに決定。理由・経緯は implementation-notes 参照）
    const folder = DriveApp.getFolderById('198ubbF5s16KukcE0H67fy8gdh8Jw-7k5');
    const pdfFile = folder.createFile(pdfBlob);
    docFile.setTrashed(true);
    const url = pdfFile.getUrl();
    _recordDocLog('bulk', 'payslip_bulk', yearMonth, url);
    return { ok: true, url, total: userIds.length };
  } catch(e) {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(_) {}
    return { ok: false, msg: e.message };
  }
}

// DocLogに指定した doc_type + target_period の記録が既に存在するか判定する純粋関数
function _hasDocLogFor_(docLogRows, docType, targetPeriod) {
  return (docLogRows || []).some(r => r.doc_type === docType && r.target_period === targetPeriod);
}

// 実行日がその月の1〜7日の範囲内か（＝毎週月曜発火のトリガーの中で「第一月曜日」に当たるか）を判定する純粋関数
// 2026-09-18: トリガーをGAS標準機能で「第n月曜日」指定できないため、毎週月曜発火＋この関数での日付ガードで代替
function _isWithinFirstWeekOfMonth_(date) {
  const d = date instanceof Date ? date.getDate() : NaN;
  return d >= 1 && d <= 7;
}

// 対象月（前月）のPayRecordが全スタッフ分支払済みになっていたら、
// 賃金台帳_全スタッフPDFを自動生成してDriveに保存する。
// 対象月に勤務実績のあるスタッフ（PayRecordが存在する）だけを判定対象とする。
// 冪等: DocLog(doc_type='payslip_bulk', target_period=対象月) が既に存在する場合は再生成しない。
// トリガー: 毎週月曜17時（01_Setup.gs setupTriggers() 参照）。
// 「第一月曜日」はGASトリガーで直接指定できないため、毎週発火させた上で
// 実行日が1〜7日の範囲内（＝その月の最初の月曜）でなければ即returnする。
function checkAndGenerateMonthlyPayslip() {
  const now = getJstNow();
  if (!_isWithinFirstWeekOfMonth_(now)) {
    Logger.log('[checkAndGenerateMonthlyPayslip] 第一月曜日ではないためスキップ (day=' + now.getDate() + ')');
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[checkAndGenerateMonthlyPayslip] ロック取得失敗のためスキップ');
    return;
  }
  try {
    // 前月を算出（archiveLastMonth と同じ方式。now は関数先頭のガードで取得済みのものを再利用）
    const year  = now.getFullYear();
    const month = now.getMonth(); // 0始まり → 前月
    const lastMonth = month === 0 ? 12 : month;
    const lastYear  = month === 0 ? year - 1 : year;
    const yearMonth = lastYear + '-' + String(lastMonth).padStart(2, '0');

    const payRecords = sheetToObjects(getSheet(SHEET.PAY_RECORD))
      .filter(r => safeDate(r.date).startsWith(yearMonth));

    if (!payRecords.length) {
      Logger.log('[checkAndGenerateMonthlyPayslip] 対象月(' + yearMonth + ')にPayRecordなし。スキップ');
      return;
    }
    if (!_isMonthFullyPaid_(payRecords)) {
      Logger.log('[checkAndGenerateMonthlyPayslip] 対象月(' + yearMonth + ')は未払いが残っているためスキップ');
      return;
    }

    const docLogRows = sheetToObjects(getSheet(SHEET.DOC));
    if (_hasDocLogFor_(docLogRows, 'payslip_bulk', yearMonth)) {
      Logger.log('[checkAndGenerateMonthlyPayslip] 対象月(' + yearMonth + ')は生成済み（DocLogあり）。スキップ');
      return;
    }

    const result = generateBulkPayslipCombined(yearMonth);
    if (result && result.ok) {
      Logger.log('[checkAndGenerateMonthlyPayslip] 生成完了: ' + yearMonth + ' total=' + result.total + ' url=' + result.url);
    } else {
      Logger.log('[checkAndGenerateMonthlyPayslip] 生成失敗: ' + yearMonth + ' msg=' + (result && result.msg));
    }
  } catch (e) {
    Logger.log('[checkAndGenerateMonthlyPayslip] エラー: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// 出勤簿PDF生成
function generateAttendanceSheet(userId, yearMonth) {
  const [y, m] = yearMonth.split('-').map(Number);
  const _allAttendLogs = getWorkLogByMonthIncludingArchive_(y, m).filter(r => r.user_id === userId);
  const _attendCorrMap = _buildWLCorrMap_(_allAttendLogs);
  const allLogs = _allAttendLogs
    .filter(r => !r.corrects && r.action !== 'deleted')
    .filter(r => !(_attendCorrMap[r.log_id] && _attendCorrMap[r.log_id].action === 'deleted'));

  // 日付ごとに最終レコード（checkout → approved → checkin の優先順）を1件に集約
  const byDate = {};
  allLogs.forEach(r => {
    const ds = _docFmtDate(r.date);
    if (!ds) return;
    const prev = byDate[ds];
    const rank = { checkout: 3, approved: 2, checkin: 1 };
    const rRank = rank[r.action] || 0;
    const pRank = prev ? (rank[prev.action] || 0) : -1;
    if (!prev || rRank > pRank) byDate[ds] = r;
  });

  const staff = getStaff(userId);
  const name = staff ? (staff.name || staff.display_name || userId) : userId;
  const company = _safeGetProp('COMPANY_NAME') || '';

  const DOW = ['日','月','火','水','木','金','土'];
  const dates = Object.keys(byDate).sort();

  let totalDays = 0, totalHours = 0, totalWage = 0, totalTransport = 0, totalNet = 0;

  const rows = dates.map(ds => {
    const r = byDate[ds];
    const [yr2, mo2, da2] = ds.split('-').map(Number);
    const dt = new Date(yr2, mo2 - 1, da2);
    const dow = DOW[dt.getDay()];
    const isSun = dt.getDay() === 0, isSat = dt.getDay() === 6;
    const color = isSun ? '#c00' : isSat ? '#006' : '#000';
    const st = _docFmtTime(r.start_time);
    const et = _docFmtTime(r.end_time);
    const h  = Number(r.actual_hours || 0);
    const w  = Number(r.wage || 0);
    const tr = Number(r.transport || 0);
    const net= Number(r.net_pay || 0);
    if (r.action === 'checkout' || r.action === 'approved') {
      totalDays++; totalHours += h; totalWage += w; totalTransport += tr; totalNet += net;
    }
    return `<tr>
      <td style="color:${color};">${mo2}/${String(da2).padStart(2,'0')}</td>
      <td style="color:${color};">${dow}</td>
      <td>${st}</td>
      <td>${et}</td>
      <td>${h || ''}</td>
      <td style="text-align:right;">${w ? '¥'+w.toLocaleString() : ''}</td>
      <td style="text-align:right;">${tr ? '¥'+tr.toLocaleString() : ''}</td>
      <td style="text-align:right;">${net ? '¥'+net.toLocaleString() : ''}</td>
    </tr>`;
  }).join('');

  const [yy, mm] = yearMonth.split('-');
  const periodLabel = `${yy}年${Number(mm)}月`;

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    body{font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif;margin:30px;font-size:12px;}
    h1{text-align:center;font-size:18px;letter-spacing:4px;margin-bottom:4px;}
    .meta{display:flex;justify-content:space-between;margin-bottom:16px;font-size:12px;}
    table{width:100%;border-collapse:collapse;}
    td,th{border:1px solid #999;padding:5px 6px;text-align:center;}
    th{background:#e8e8e8;font-weight:bold;}
    .total-row{font-weight:bold;background:#f5f5dc;}
    .seal-box{display:flex;justify-content:flex-end;gap:16px;margin-top:20px;}
    .seal{border:1px solid #999;width:60px;height:60px;text-align:center;padding-top:8px;font-size:11px;}
    p.note{font-size:10px;color:#666;margin-top:12px;}
  </style></head><body>
  <h1>出 勤 簿</h1>
  <div class="meta">
    <span>氏名：<b>${name}</b></span>
    <span>期間：<b>${periodLabel}</b></span>
    <span>事業所：${company}</span>
  </div>
  <table>
    <tr>
      <th>日付</th><th>曜日</th><th>出勤</th><th>退勤</th>
      <th>実働(h)</th><th>日当</th><th>交通費</th><th>手取り</th>
    </tr>
    ${rows}
    <tr class="total-row">
      <td colspan="4" style="text-align:right;">合　計</td>
      <td>${totalDays}日 / ${totalHours}h</td>
      <td style="text-align:right;">¥${totalWage.toLocaleString()}</td>
      <td style="text-align:right;">¥${totalTransport.toLocaleString()}</td>
      <td style="text-align:right;">¥${totalNet.toLocaleString()}</td>
    </tr>
  </table>
  <div class="seal-box">
    <div class="seal">確認<br><br></div>
    <div class="seal">承認<br><br></div>
  </div>
  <p class="note">※ 労働基準法第108条に基づき5年間保存</p>
  </body></html>`;

  const folder = getDriveFolder('出勤簿');
  const fileName = `出勤簿_${name}_${yearMonth}`;
  try {
    const url = _htmlToPdf(html, fileName, folder);
    _recordDocLog(userId, 'attendance', yearMonth, url);
    return { ok: true, url };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// 支払調書PDF生成（年間50万円超のみ）
function generatePaymentReport(userId, year) {
  const staff = getStaff(userId);
  const name = staff ? (staff.name || staff.display_name || userId) : userId;

  // アーカイブシートも含めて年間全データを取得（月次アーカイブ移行後も正確に集計）
  const _allPayLogs = getAllWorkLogWithArchives_({ userId: String(userId) });
  const _payCorrMap = _buildWLCorrMap_(_allPayLogs);
  const _payFiltered = _allPayLogs
    .filter(r => safeDate(r.date).startsWith(String(year)))
    .filter(r => !r.corrects && (r.action === 'checkout' || r.action === 'approved'))
    .filter(r => !(_payCorrMap[r.log_id] && _payCorrMap[r.log_id].action === 'deleted'));
  const _payCOKeys = new Set(
    _payFiltered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const wlogs = _payFiltered.filter(r =>
    r.action === 'checkout' || !_payCOKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const totalWage = wlogs.reduce((s, r) => s + Number(r.wage || 0), 0);
  const totalWithholding = wlogs.reduce((s, r) => s + Number(r.withholding || 0), 0);

  if (totalWage < 500000) {
    return { ok: false, msg: `年間支払金額 ¥${totalWage.toLocaleString()} は50万円以下のため支払調書の作成対象外です（所得税法第225条）` };
  }

  const company = {
    name:    _safeGetProp('COMPANY_NAME'),
    address: _safeGetProp('COMPANY_ADDRESS'),
    tel:     _safeGetProp('COMPANY_TEL'),
  };

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    body{font-family:sans-serif;margin:40px;}
    h1{text-align:center;} table{width:100%;border-collapse:collapse;}
    td,th{border:1px solid #ccc;padding:8px;} th{background:#f5f5f5;}
    .warn{color:red;font-size:12px;margin-top:20px;}
  </style></head><body>
  <h1>報酬、料金、契約金及び賞金の支払調書</h1>
  <table>
    <tr><th>支払を受ける者</th><td>${name}</td></tr>
    <tr><th>対象年</th><td>${year}年</td></tr>
    <tr><th>支払金額</th><td>¥${totalWage.toLocaleString()}</td></tr>
    <tr><th>源泉徴収税額</th><td>¥${totalWithholding.toLocaleString()}</td></tr>
  </table>
  <p>支払者: ${company.name} ${company.address} TEL:${company.tel}</p>
  <p class="warn">⚠️ 税理士に確認を推奨します（所得税法第225条）</p>
  </body></html>`;

  const folder = getDriveFolder('支払調書');
  const fileName = `支払調書_${name}_${year}`;
  try {
    const url = _htmlToPdf(html, fileName, folder);
    _recordDocLog(userId, 'payment_report', String(year), url);
    return { ok: true, url };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// 源泉徴収票PDF生成（給与所得の源泉徴収票・日額表丙欄適用）
// endDate 省略時は ${year}-12-31（全年）。上半期なら ${year}-06-30 を渡す
function generateWithholdingCert(userId, year, endDate) {
  const startDate = String(year) + '-01-01';
  const end       = endDate ? String(endDate) : (String(year) + '-12-31');

  const staff = getStaff(userId);
  const name = staff ? (staff.name || staff.display_name || userId) : userId;
  const zipCode   = staff ? (staff.zip_code   || '') : '';
  const address   = staff ? (staff.address    || '') : '';
  const birthDate = staff ? (staff.birth_date || '') : '';

  const _allWithLogs = getAllWorkLogWithArchives_({ userId: String(userId) });
  const _withCorrMap = _buildWLCorrMap_(_allWithLogs);
  const _withFiltered = _allWithLogs
    .filter(r => { const d = safeDate(r.date); return d >= startDate && d <= end; })
    .filter(r => (r.action === 'checkout' || r.action === 'approved') && !r.corrects)
    .filter(r => !(_withCorrMap[r.log_id] && _withCorrMap[r.log_id].action === 'deleted'));
  const _withCOKeys = new Set(
    _withFiltered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const wlogs = _withFiltered.filter(r =>
    r.action === 'checkout' || !_withCOKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );

  const totalWage        = wlogs.reduce((s, r) => s + (Number(r.wage) || 0), 0);
  const totalWithholding = wlogs.reduce((s, r) => s + (Number(r.withholding) || 0), 0);
  const workDays         = new Set(wlogs.map(r => safeDate(r.date))).size;

  const company = {
    name:    _safeGetProp('COMPANY_NAME'),
    address: _safeGetProp('COMPANY_ADDRESS'),
    tel:     _safeGetProp('COMPANY_TEL'),
  };

  const addressLine = zipCode ? '〒' + zipCode + ' ' + address : (address || '（未登録）');
  const issueDate = formatDate(getJstNow());

  // 期間ラベル: 例「2026年1月1日 〜 2026年6月30日」
  const [sy, sm, sd] = startDate.split('-').map(Number);
  const [ey, em, ed] = end.split('-').map(Number);
  const periodLabel = `${sy}年${sm}月${sd}日 〜 ${ey}年${em}月${ed}日`;
  const yearLabel   = sy === ey ? sy + '年分' : sy + '〜' + ey + '年分';

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    body{font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif;margin:40px;font-size:13px;}
    h1{text-align:center;font-size:18px;letter-spacing:4px;border-bottom:2px solid #333;padding-bottom:8px;margin-bottom:20px;}
    .year-label{text-align:right;font-size:13px;margin-top:-20px;margin-bottom:20px;}
    table{width:100%;border-collapse:collapse;margin-bottom:16px;}
    td,th{border:1px solid #999;padding:8px 10px;}
    th{background:#f0f0f0;font-weight:bold;width:200px;text-align:left;}
    .section-title{font-weight:bold;margin:16px 0 6px;font-size:14px;border-left:4px solid #333;padding-left:8px;}
    .note{font-size:11px;color:#666;margin-top:24px;}
    .warn{color:#c00;font-size:12px;margin-top:8px;}
  </style></head><body>
  <h1>給与所得の源泉徴収票</h1>
  <div class="year-label">${yearLabel}</div>

  <div class="section-title">支払を受ける者</div>
  <table>
    <tr><th>住所</th><td>${addressLine}</td></tr>
    <tr><th>氏名</th><td>${name}</td></tr>
    <tr><th>生年月日</th><td>${birthDate || '（未登録）'}</td></tr>
  </table>

  <div class="section-title">支払内容</div>
  <table>
    <tr><th>対象期間</th><td>${periodLabel}</td></tr>
    <tr><th>種別</th><td>給与</td></tr>
    <tr><th>支払金額</th><td>¥${totalWage.toLocaleString()}</td></tr>
    <tr><th>源泉徴収税額</th><td>¥${totalWithholding.toLocaleString()}</td></tr>
    <tr><th>社会保険料等の金額</th><td>¥0</td></tr>
    <tr><th>給与所得控除後の金額</th><td>─（日額表丙欄適用につき省略）</td></tr>
    <tr><th>所得控除の額の合計額</th><td>─</td></tr>
    <tr><th>勤務日数</th><td>${workDays}日</td></tr>
  </table>

  <div class="section-title">支払者</div>
  <table>
    <tr><th>名称</th><td>${company.name}</td></tr>
    <tr><th>住所</th><td>${company.address}</td></tr>
    <tr><th>電話</th><td>${company.tel}</td></tr>
  </table>

  <p class="note">発行日: ${issueDate}　　※ 所得税法第185条 日額表丙欄適用（日雇労働者）</p>
  <p class="warn">⚠️ この書類は参考資料です。確定申告の際は税理士にご確認ください。</p>
  </body></html>`;

  const folder = getDriveFolder('源泉徴収票');
  const periodSuffix = end === String(year) + '-12-31' ? '' : '_' + end.replace(/-/g, '');
  const fileName = `源泉徴収票_${name}_${year}${periodSuffix}`;
  try {
    const url = _htmlToPdf(html, fileName, folder);
    _recordDocLog(userId, 'withholding_cert', startDate + '_' + end, url);
    return { ok: true, url };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// 全スタッフ源泉徴収票を1PDFに一括生成（DocumentApp方式）
// endDate 省略時は ${year}-12-31（全年）。上半期なら ${year}-06-30 を渡す
function generateBulkWithholdingCert(year, endDate) {
  const startDate = String(year) + '-01-01';
  const end       = endDate ? String(endDate) : (String(year) + '-12-31');

  const allStaff = getAllStaff().filter(s => String(s.active).toLowerCase() !== 'false' && s.line_user_id);
  if (!allStaff.length) return { ok: false, msg: '対象スタッフがいません' };

  const companyName = _safeGetProp('COMPANY_NAME');
  const companyAddr = _safeGetProp('COMPANY_ADDRESS');
  const companyTel  = _safeGetProp('COMPANY_TEL');
  const issueDate   = formatDate(getJstNow());

  const [sy, sm, sd] = startDate.split('-').map(Number);
  const [ey, em, ed] = end.split('-').map(Number);
  const periodLabel = `${sy}年${sm}月${sd}日 〜 ${ey}年${em}月${ed}日`;
  const yearLabel   = sy === ey ? sy + '年分' : sy + '〜' + ey + '年分';

  const periodSuffix = end === String(year) + '-12-31' ? '' : '_' + end.replace(/-/g, '');
  const fileName = `源泉徴収票_全スタッフ_${year}${periodSuffix}`;

  const _allValidLogs = getAllWorkLogWithArchives_({});
  const _validCorrMap = _buildWLCorrMap_(_allValidLogs);
  const _validFiltered = _allValidLogs
    .filter(r => { const d = safeDate(r.date); return d >= startDate && d <= end; })
    .filter(r => (r.action === 'checkout' || r.action === 'approved') && !r.corrects)
    .filter(r => !(_validCorrMap[r.log_id] && _validCorrMap[r.log_id].action === 'deleted'));
  const _validCOKeys = new Set(
    _validFiltered.filter(r => r.action === 'checkout').map(r => r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );
  const validLogs = _validFiltered.filter(r =>
    r.action === 'checkout' || !_validCOKeys.has(r.user_id + ':' + safeDate(r.date) + ':' + r.job_id)
  );

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();
  body.clear();

  let certCount = 0;
  allStaff.forEach((s, idx) => {
    const uid  = s.line_user_id;
    const name = s.name || s.display_name || uid;
    const zipCode   = s.zip_code   || '';
    const address   = s.address    || '';
    const birthDate = s.birth_date || '';
    const addressLine = zipCode ? '〒' + zipCode + ' ' + address : (address || '（未登録）');

    const logs = validLogs.filter(r => r.user_id === uid);
    if (!logs.length) return;

    const totalWage        = logs.reduce((acc, r) => acc + (Number(r.wage) || 0), 0);
    const totalWithholding = logs.reduce((acc, r) => acc + (Number(r.withholding) || 0), 0);
    const workDays         = new Set(logs.map(r => safeDate(r.date))).size;

    if (certCount > 0) body.appendPageBreak();

    body.appendParagraph('給与所得の源泉徴収票　' + yearLabel)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    const tableData = [
      ['支払を受ける者 — 住所', addressLine],
      ['支払を受ける者 — 氏名', name],
      ['生年月日', birthDate || '（未登録）'],
      ['対象期間', periodLabel],
      ['種別', '給与'],
      ['支払金額', '¥' + totalWage.toLocaleString()],
      ['源泉徴収税額', '¥' + totalWithholding.toLocaleString()],
      ['社会保険料等の金額', '¥0'],
      ['勤務日数', workDays + '日'],
      ['支払者', companyName + '　' + companyAddr + '　TEL:' + companyTel],
    ];
    body.appendTable(tableData);
    body.appendParagraph('発行日: ' + issueDate + '　※ 所得税法第185条 日額表丙欄適用').setFontSize(9);
    certCount++;
  });

  if (!certCount) {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(_) {}
    return { ok: false, msg: String(year) + '年の勤務実績があるスタッフがいません' };
  }

  doc.saveAndClose();
  try {
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs(MimeType.PDF);
    pdfBlob.setName(fileName + '.pdf');
    const folder = DriveApp.getFolderById('198ubbF5s16KukcE0H67fy8gdh8Jw-7k5');
    const pdfFile = folder.createFile(pdfBlob);
    docFile.setTrashed(true);
    const url = pdfFile.getUrl();
    _recordDocLog('bulk', 'withholding_cert_bulk', String(year) + (endDate ? '_' + end : ''), url);
    return { ok: true, url, total: certCount };
  } catch(e) {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(_) {}
    return { ok: false, msg: e.message };
  }
}

function _safeGetProp(key) {
  return _safeProp(key, '');
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { _hasDocLogFor_, _isWithinFirstWeekOfMonth_ };
}
