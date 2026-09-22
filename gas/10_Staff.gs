// 社内タイミー - 10_Staff.gs
// スタッフマスタ CRUD
// userId最優先：LINEメッセージ受信の瞬間にStaffへ確保

function upsertStaff(userId, displayName) {
  if (!userId) return null;
  const sh = getSheet(SHEET.STAFF);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const now = formatDateTime(getJstNow());

  if (sh.getLastRow() > 1) {
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const uidIdx = hdr.indexOf('line_user_id');
    const dnIdx  = hdr.indexOf('display_name');
    const uaIdx  = hdr.indexOf('updated_at');
    for (let i = 0; i < data.length; i++) {
      if (data[i][uidIdx] === userId) {
        // displayName を更新
        if (displayName && data[i][dnIdx] !== displayName) {
          const rowNum = i + 2;
          const rowData = sh.getRange(rowNum, 1, 1, hdr.length).getValues()[0];
          rowData[dnIdx] = displayName;
          rowData[uaIdx] = now;
          sh.getRange(rowNum, 1, 1, hdr.length).setValues([rowData]);
        }
        return getStaff(userId);
      }
    }
  }

  // 新規登録（LINE表示名を name の初期値として使用）
  const row = objToRow(hdr, {
    line_user_id:    userId,
    display_name:    displayName || '',
    name:            displayName || '',
    active:          'true',
    job_id:          'JOB001',
    tax_type:        'hei',
    bank_info:       '',
    same_day_cancel: 0,
    total_cancel:    0,
    caution_flag:    false,
    eval_avg:        '',
    eval_level:      '',
    eval_count:      0,
    memo:            '',
    registered_at:   now,
    updated_at:      now,
  });
  sh.appendRow(row);
  Logger.log('Staff新規登録: ' + userId);
  return getStaff(userId);
}

function getStaff(userId) {
  const sh = getSheet(SHEET.STAFF);
  if (sh.getLastRow() < 2) return null;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const uidIdx = hdr.indexOf('line_user_id');
  for (const row of data) {
    if (row[uidIdx] === userId) {
      const obj = {};
      hdr.forEach((h, i) => { if (h) obj[String(h).trim()] = row[i]; });
      return obj;
    }
  }
  return null;
}

function updateStaff(userId, fields) {
  const sh = getSheet(SHEET.STAFF);
  if (sh.getLastRow() < 2) return false;
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const uidIdx = hdr.indexOf('line_user_id');
  const uaIdx  = hdr.indexOf('updated_at');
  const now = formatDateTime(getJstNow());

  for (let i = 0; i < data.length; i++) {
    if (data[i][uidIdx] === userId) {
      const rowNum = i + 2;
      // 1行まとめて読んで書き直す
      const rowData = sh.getRange(rowNum, 1, 1, hdr.length).getValues()[0];
      Object.keys(fields).forEach(key => {
        const colIdx = hdr.indexOf(key);
        if (colIdx >= 0) rowData[colIdx] = fields[key];
      });
      rowData[uaIdx] = now;
      sh.getRange(rowNum, 1, 1, hdr.length).setValues([rowData]);
      return true;
    }
  }
  return false;
}

// 全スタッフ一覧（統計付き）
function getAllStaff() {
  return sheetToObjects(getSheet(SHEET.STAFF));
}

// 評価平均・レベルを更新
function refreshStaffEval(userId) {
  const evals = sheetToObjects(getSheet(SHEET.EVAL)).filter(r => r.user_id === userId);
  if (evals.length === 0) return;
  const avg = evals.reduce((s, r) => s + Number(r.score_avg || 0), 0) / evals.length;
  const rounded = Math.round(avg * 10) / 10;
  let level = 'C';
  if (rounded >= 4.5) level = 'S';
  else if (rounded >= 4.0) level = 'A';
  else if (rounded >= 3.0) level = 'B';
  else if (rounded >= 2.0) level = 'C';
  else level = 'D';
  updateStaff(userId, { eval_avg: rounded, eval_level: level, eval_count: evals.length });
}
