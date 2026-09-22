// 社内タイミー - 70_Archive.gs
// 月次アーカイブ（Applied / WorkLog の前月データを別シートへ移動）

function archiveLastMonth() {
  const now = getJstNow();
  const year  = now.getFullYear();
  const month = now.getMonth(); // 0始まり → 前月
  const lastMonth = month === 0 ? 12 : month;
  const lastYear  = month === 0 ? year - 1 : year;
  const ym = lastYear + '_' + String(lastMonth).padStart(2, '0');
  const prefix = lastYear + '-' + String(lastMonth).padStart(2, '0');

  Logger.log('アーカイブ開始: ' + prefix);
  _archiveSheet(SHEET.APPLIED,  'Applied_' + ym,  prefix);
  _archiveSheet(SHEET.WORKLOG,  'WorkLog_' + ym,   prefix);
  Logger.log('アーカイブ完了');
}

function _archiveSheet(srcName, archiveName, prefix) {
  const ss = getSpreadsheet();
  const src = ss.getSheetByName(srcName);
  if (!src || src.getLastRow() < 2) return;

  const hdr  = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
  const data = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues();

  // 日付列を特定（date または timestamp）
  let dateIdx = hdr.indexOf('date');
  if (dateIdx < 0) dateIdx = hdr.indexOf('timestamp');
  if (dateIdx < 0) {
    Logger.log('[_archiveSheet] date/timestamp 列が見つかりません: ' + src.getName());
    return;
  }

  const toMove = data.filter(row => {
    const d = safeDate(row[dateIdx]);
    return d.startsWith(prefix);
  });
  if (toMove.length === 0) {
    Logger.log('移動対象なし: ' + srcName);
    return;
  }

  // アーカイブシート作成/取得
  let arch = ss.getSheetByName(archiveName);
  if (!arch) {
    arch = ss.insertSheet(archiveName);
    arch.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    Logger.log('アーカイブシート作成: ' + archiveName);
  }

  // 既存アーカイブのdedupeキー取得（冪等）
  const existingKeys = new Set();
  if (arch.getLastRow() > 1) {
    const archData = arch.getRange(2, 1, arch.getLastRow() - 1, arch.getLastColumn()).getValues();
    const archHdr  = arch.getRange(1, 1, 1, arch.getLastColumn()).getValues()[0];
    const dkIdx    = archHdr.indexOf('dedupe_key');
    const lidIdx   = archHdr.indexOf('log_id');
    archData.forEach(r => {
      if (dkIdx >= 0 && r[dkIdx]) existingKeys.add(r[dkIdx]);
      if (lidIdx >= 0 && r[lidIdx]) existingKeys.add(r[lidIdx]);
    });
  }
  const dkIdx  = hdr.indexOf('dedupe_key');
  const lidIdx = hdr.indexOf('log_id');

  // キーなし行は毎月重複アーカイブされるリスクがあるため、
  // キーを持つ行のみアーカイブ対象とする（キーなし行は元シートに残す）
  const newRows = toMove.filter(r => {
    const key = dkIdx >= 0 ? r[dkIdx] : (lidIdx >= 0 ? r[lidIdx] : null);
    if (!key) return false; // キーなし行はスキップ
    return !existingKeys.has(key);
  });

  if (newRows.length > 0) {
    arch.getRange(arch.getLastRow() + 1, 1, newRows.length, hdr.length).setValues(newRows);
  }

  // アーカイブに存在が確認できたキー（既存アーカイブ分＋今回書き込み分）
  const archivedKeys = new Set(existingKeys);
  newRows.forEach(r => {
    const key = dkIdx >= 0 ? r[dkIdx] : (lidIdx >= 0 ? r[lidIdx] : null);
    if (key) archivedKeys.add(key);
  });

  // D-002（DECISIONS.md）: Applied / WorkLog / CancelLog は行削除禁止。
  // アーカイブ後も元シートの行は削除しない（append-only を守る）。
  // アーカイブシートへのコピーのみ行い、元シートはそのまま保持する。
  // ※ シートが肥大化した場合は partialReset() で対処すること（手動実行）。

  Logger.log(`${srcName} → ${archiveName}: ${newRows.length}件コピー（元シート行は保持）`);
}

// アーカイブシート一覧を返す
function getArchiveSheetNames() {
  const sheets = getSpreadsheet().getSheets();
  const result = [];
  sheets.forEach(sh => {
    const name = sh.getName();
    const m = name.match(/^(Applied|WorkLog)_(\d{4})_(\d{2})$/);
    if (m) {
      result.push({
        name,
        source:    m[1],
        yearMonth: m[2] + '-' + m[3],
      });
    }
  });
  return result.sort((a, b) => b.yearMonth.localeCompare(a.yearMonth));
}

// ロールバック用：アーカイブを元シートに戻す
function restoreArchive(archiveName) {
  const ss = getSpreadsheet();
  const arch = ss.getSheetByName(archiveName);
  if (!arch || arch.getLastRow() < 2) return false;

  const m = archiveName.match(/^(Applied|WorkLog)_/);
  if (!m) return false;
  const srcName = m[1] === 'Applied' ? SHEET.APPLIED : SHEET.WORKLOG;

  const src = getSheet(srcName);
  const hdr = arch.getRange(1, 1, 1, arch.getLastColumn()).getValues()[0];
  const data = arch.getRange(2, 1, arch.getLastRow() - 1, arch.getLastColumn()).getValues();

  const existingKeys = new Set();
  const srcHdr = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
  const srcDkIdx = srcHdr.indexOf('dedupe_key');
  const srcLidIdx = srcHdr.indexOf('log_id');
  if (src.getLastRow() > 1) {
    const srcData = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues();
    srcData.forEach(r => {
      if (srcDkIdx >= 0 && r[srcDkIdx]) existingKeys.add(r[srcDkIdx]);
      if (srcLidIdx >= 0 && r[srcLidIdx]) existingKeys.add(r[srcLidIdx]);
    });
  }

  const dkIdx = hdr.indexOf('dedupe_key');
  const lidIdx = hdr.indexOf('log_id');
  const rows = data.filter(row => {
    const key = dkIdx >= 0 ? row[dkIdx] : (lidIdx >= 0 ? row[lidIdx] : '');
    return !key || !existingKeys.has(key);
  });
  if (rows.length > 0) {
    src.getRange(src.getLastRow() + 1, 1, rows.length, hdr.length).setValues(rows);
  }
  Logger.log('アーカイブ復元: ' + archiveName + ' → ' + srcName + ' ' + rows.length + '件');
  return true;
}

// 毎週月曜 8時：翌週日程を全スタッフへ個別URL付きで送信
function sendWeeklyRecruit() {
  const autoEnabled = PropertiesService.getScriptProperties().getProperty('AUTO_RECRUIT');
  if (autoEnabled === 'false') {
    Logger.log('AUTO_RECRUIT=false のためスキップ');
    return;
  }
  const now = getJstNow();
  // 翌週月〜日の日付リスト
  const days = [];
  for (let i = 1; i <= 7; i++) {
    const d = new Date(now);
    d.setDate(now.getDate() + i);
    days.push(formatDate(d));
  }

  const jobs = sheetToObjects(getSheet(SHEET.JOB)).filter(j => String(j.active).toLowerCase() === 'true');
  if (jobs.length === 0) { Logger.log('有効な職種なし'); return; }
  const job = jobs[0]; // 最初のアクティブ職種

  // コピーしやすいよう1行1日付（曜日あり）形式にする。replyExamplesをdatesListと同一にすることで
  // ensureRecruitMessageParts の重複チェックが効き、LINE返信用リストの二重掲載を防止する
  const datesList = days.map(formatRecruitDateLine).join('\n');
  const replyExamples = datesList;

  const template = _getMsgTemplate_();
  const phoneRule = _getMsgConfig_('PHONE_RULE');
  const staffList = sheetToObjects(getSheet(SHEET.STAFF))
    .filter(s => s.line_user_id && String(s.active).toLowerCase() !== 'false');
  let count = 0;

  // 日当は時給×勤務時間で計算する（job.wageは時給。「日当」表記に時給をそのまま出さない）
  const dailyCalc = calcDailyWage(job.wage, job.start_time, job.end_time, job.transport);

  staffList.forEach(s => {
    const token = _issueToken(s.line_user_id, 'apply', { jobId: job.job_id, dates: days }, _ttlToEndOfDayJst_(3));
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
      Logger.log('[WeeklyRecruit] 送信失敗 userId=' + s.line_user_id + ' name=' + (s.name || '') + ' err=' + e.message);
    }
  });
  Logger.log('sendWeeklyRecruit: ' + count + '件送信');
}

function _getMsgConfig_(key) {
  try {
    const rows = sheetToObjects(getSheet(SHEET.MSG_CONFIG));
    const row = rows.find(r => r.key === key);
    return row ? String(row.value) : '';
  } catch(e) {
    Logger.log('[_getMsgConfig_] 取得失敗: ' + e.message);
    return '';
  }
}

function _getMsgTemplate_() {
  const t = _getMsgConfig_('MSG_TEMPLATE');
  return t || DEFAULT_MSG_TEMPLATE;
}

// 毎日 18時：翌日確定スタッフへリマインド
function sendNextDayReminder() {
  const now = getJstNow();
  const tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);
  const tomorrowStr = formatDate(tomorrow);

  const latestMap = buildLatestAppliedMap_(
    getAppliedAllWithArchives_().filter(r => safeDate(r.date) === tomorrowStr)
  );
  const applied = Object.values(latestMap)
    .filter(r => r.event === 'approved');

  const seen = new Set();
  applied.forEach(r => {
    if (seen.has(r.user_id)) return;
    seen.add(r.user_id);
    const job = getJob(r.job_id);
    const company = _safeGetCompany();
    try {
      pushMessage(r.user_id,
        `明日もよろしくお願いします！🙏\n` +
        `${formatDateJa(tomorrowStr)} ${job ? _safeTime_(job.start_time) + '〜' + _safeTime_(job.end_time) : ''}\n` +
        `${company}\n体調に気をつけてお越しください`
      );
    } catch(e) {
      console.error('[NextDayReminder] 送信失敗: ' + r.user_id);
    }
    Utilities.sleep(200);
  });
  Logger.log('sendNextDayReminder: ' + seen.size + '件送信');
}

function _safeGetCompany() {
  return _safeProp('COMPANY_NAME', '');
}

// 4月の確定スタッフに一斉通知（管理画面ボタンから実行）
// 実行月の前月を対象とする
function notifyAprilApproved() {
  const now = getJstNow();
  const y = now.getFullYear();
  const m = now.getMonth(); // 0-indexed: 0=1月
  const prev = m === 0 ? (y - 1) + '-12' : y + '-' + String(m).padStart(2, '0');
  const latestMap = buildLatestAppliedMap_(
    getAppliedAllWithArchives_().filter(function(r) { return safeDate(r.date).startsWith(prev); })
  );
  const april = Object.values(latestMap)
    .filter(function(r) { return r.event === 'approved'; });

  if (!april.length) return { ok: false, msg: '4月の確定データがありません', count: 0 };

  const byUser = {};
  april.forEach(function(r) {
    if (!byUser[r.user_id]) byUser[r.user_id] = [];
    byUser[r.user_id].push(r);
  });

  const staffMap = {};
  getAllStaff().forEach(function(s) { staffMap[s.line_user_id] = s; });

  let sent = 0;
  Object.keys(byUser).forEach(function(uid) {
    const rows = byUser[uid].sort(function(a, b) { return a.date.localeCompare(b.date); });
    const staff = staffMap[uid] || {};
    const name = staff.name || staff.display_name || uid;
    const lines = rows.map(function(r) {
      const job = getJob(r.job_id);
      const time = job ? (_safeTime_(job.start_time) + '〜' + _safeTime_(job.end_time)) : '';
      return '・' + formatDateJa(safeDate(r.date)) + (job ? ' ' + job.title : '') + (time ? ' ' + time : '');
    }).join('\n');
    const msg = '【4月のシフト確定のご連絡】\n' + name + 'さん、以下の日程が確定しました✅\n' + lines +
      '\n当日はよろしくお願いします🙏\nキャンセルはあまりしないで、ください';
    try {
      pushMessage(uid, msg);
      sent++;
    } catch(e) {
      console.error('[notifyAprilApproved] 送信失敗: ' + uid);
    }
    Utilities.sleep(500);
  });

  Logger.log('notifyAprilApproved: ' + sent + '名に送信');
  return { ok: true, count: sent };
}

// 毎週月曜 19時：田内あきこへ翌週火・土の応募依頼を送信
function sendWeeklyApplyRequest_Tauchi() {
  const userId = _safeProp('TAUCHI_LINE_USER_ID', '');
  if (!userId) {
    Logger.log('[tauchiWeekly] TAUCHI_LINE_USER_ID が未設定です。setupTauchiUserId() を実行してください。');
    return;
  }

  const now = getJstNow();
  // 月曜以外での誤実行を防ぐ（0=日,1=月,...,6=土）
  if (now.getDay() !== 1) {
    Logger.log('[tauchiWeekly] 月曜以外のためスキップ: day=' + now.getDay());
    return;
  }

  // 翌週火曜(+8日)・土曜(+12日)を計算
  const tue = new Date(now.getTime() + 8 * 24 * 60 * 60 * 1000);
  const sat = new Date(now.getTime() + 12 * 24 * 60 * 60 * 1000);
  const tueDateStr = formatDate(tue);
  const satDateStr = formatDate(sat);

  const tauchi = getAllStaff().find(function(s) { return s.name === '田内あきこ'; });
  if (!tauchi) {
    Logger.log('[tauchiWeekly] 田内あきこ が Staff シートに見つかりません。');
    return;
  }

  const job = tauchi.job_id ? getJob(tauchi.job_id) : null;
  let timeStr = '';
  let wageLine = '';
  let transportLine = '';
  let totalLine = '';
  if (job) {
    timeStr = _safeTime_(job.start_time) + '〜' + _safeTime_(job.end_time);
    const calc = calcDailyWage(Number(job.wage) || 0, job.start_time, job.end_time, Number(job.transport) || 0);
    wageLine = '\n💰 時給：' + Number(job.wage).toLocaleString() + '円';
    transportLine = '\n🚌 交通費：' + Number(job.transport).toLocaleString() + '円';
    totalLine = '\n💴 合計：' + calc.total.toLocaleString() + '円';
  }

  const msg =
    '【シフト応募のお願い】\n' +
    (job ? job.title : '温泉スタッフ') + '\n\n' +
    formatRecruitDateLine(tueDateStr) + '\n' +
    formatRecruitDateLine(satDateStr) +
    (timeStr ? '\n\n⏰ 時間：' + timeStr : '') +
    wageLine + transportLine + totalLine + '\n\n' +
    '応募する日の数字を返信してください。\n' +
    '（例：' + formatRecruitReplyExample(tueDateStr) + '）';

  try {
    pushMessage(userId, msg);
    Logger.log('[tauchiWeekly] 送信完了: ' + tueDateStr + ' / ' + satDateStr);
  } catch(e) {
    Logger.log('[tauchiWeekly] 送信失敗: ' + e.message);
  }
}
