// 社内タイミー - 90_Webhook.gs
// LINE Webhook 処理

// ※ LINE署名検証について：
// GAS Web Apps はリクエストヘッダーを取得できないため HMAC-SHA256 検証は不可。
// URL クエリパラメータ方式も GAS の 302 リダイレクト問題により使用不可。
// Webhook URL 自体が推測困難な長い文字列のため、現状は URL の秘匿性に依存する。

function doPost(e) {
  const output = ContentService.createTextOutput('OK');

  try {
    const body = JSON.parse(e.postData.contents);

    // フォームからの直接POST（register / apply）
    if (body.action === 'register') {
      return _handleFormRegister(body);
    }
    if (body.action === 'apply') {
      return _handleFormApply(body);
    }

    // LINE Webhook
    const events = body.events;
    if (!events || events.length === 0) return output;
    events.forEach(_handleLineEvent_);
  } catch(err) {
    console.error('[Webhook] doPost error: ' + err.message);
  }

  return output;
}

function _handleLineEvent_(ev) {
  if (!ev) return;

  // 友だち追加 → ウェルカム＋氏名登録リンク
  if (ev.type === 'follow') {
    const userId = ev.source.userId;
    upsertStaff(userId, '');
    _updateDisplayNameAsync(userId);
    const url = _getWebAppUrl() + '?page=register&token=' + _issueToken(userId, 'register');
    pushMessage(userId,
      '地涌の温泉 社内タイミーへようこそ！👋\n\n' +
      'まずはじめに、お名前を登録してください。\n' +
      '以下のリンクから入力できます 📝\n\n' + url
    );
    return;
  }

  if (ev.type !== 'message') return;
  if (!ev.message || ev.message.type !== 'text') return;

  const userId      = ev.source.userId;
  const displayName = ev.source.displayName || '';
  const text        = ev.message.text || '';
  const replyToken  = ev.replyToken;

  // userId を即時確保（Step 3必須）
  upsertStaff(userId, displayName);

  // プロフィール非同期更新（displayName確実取得）
  _updateDisplayNameAsync(userId);

  _appendChatLog(userId, 'in', text);
  handleLineMessage(userId, text, replyToken);
}

function handleLineMessage(userId, text, replyToken) {
  const t = _normalizeMessageText_(text).trim();

  // 返信 + ChatLog同時保存のローカルヘルパー
  function reply(msg) {
    replyMessage(replyToken, msg);
    _appendChatLog(userId, 'out', msg);
  }

  // 名前・電話をログから自動抽出・更新（常時）
  _extractAndUpdateProfile(userId, t);

  // ── コマンドキーワード ──────────────────────────
  if (/登録|名前|更新|プロフィール/.test(t)) {
    const url = _getWebAppUrl() + '?page=register&token=' + _issueToken(userId, 'register');
    reply('プロフィールの登録・更新はこちらから 📝\n\n' + url);
    return;
  }

  if (/確認/.test(t)) {
    const myApplied = getMyApplied(userId);
    if (myApplied.length === 0) {
      reply('現在の応募はありません。\n希望日を送ってください😊');
    } else {
      const lines = myApplied.map(r => {
        const job = getJob(r.job_id);
        const status = r.event === 'approved' ? '✅確定' : '⏳審査中';
        return `${formatDateJa(r.date)} ${job ? job.title : r.job_id} ${status}`;
      });
      reply('現在の応募状況：\n' + lines.join('\n'));
    }
    return;
  }

  if (/キャンセル|取り消し|取消|取下げ|取下|辞退/.test(t)) {
    const today = formatDate(getJstNow());
    const dates = extractDates(t, today, { keepQuestions: true });
    if (dates.length === 0) {
      reply('キャンセルしたい日程を教えてください。\n例：5月9日をキャンセル');
      return;
    }

    const CANCEL_DEADLINE_DAYS = 2;
    const cancelled = [];
    const notFound = [];
    const tooLate = [];
    dates.forEach(d => {
      const diffDays = Math.floor((new Date(d + 'T00:00:00') - new Date(today + 'T00:00:00')) / 86400000);
      if (diffDays <= CANCEL_DEADLINE_DAYS) {
        tooLate.push(d);
        return;
      }
      const shifts = getCancellableShifts(userId, d);
      if (shifts.length === 0) {
        notFound.push(d);
        return;
      }
      shifts.forEach(s => {
        const r = cancelShift(userId, d, s.job_id, 'line_cancel', 'line');
        if (r && r.ok) cancelled.push({ date: d, jobId: s.job_id });
      });
    });

    if (tooLate.length > 0 && cancelled.length === 0) {
      reply('直前のキャンセルはLINEでは受け付けられません🙏\n' +
        tooLate.map(formatDateJa).join('\n') +
        '\n\n担当者に直接ご連絡ください。');
      return;
    }

    if (cancelled.length === 0 && tooLate.length === 0) {
      reply('該当する応募・確定シフトが見つかりませんでした。\n日付を確認してもう一度送ってください。');
      return;
    }

    const lines = cancelled.map(c => {
      const job = getJob(c.jobId);
      return '・' + formatDateJa(c.date) + (job ? ' ' + job.title : '');
    }).join('\n');
    let msg = 'キャンセルを受け付けました。\n' + lines + '\n\n担当者にも共有します。';
    if (tooLate.length > 0) {
      msg += '\n\n以下は直前のため担当者に直接ご連絡ください：\n' + tooLate.map(formatDateJa).join('\n');
    }
    if (notFound.length > 0) {
      msg += '\n\n確認できなかった日程：\n' + notFound.map(formatDateJa).join('\n');
    }
    reply(msg);

    try {
      pushMessage(SECRETS.ADMIN_USER_ID,
        '⚠️ LINEキャンセル受付\n' + userId + '\n' +
        cancelled.map(c => c.date + ' ' + c.jobId).join('\n')
      );
    } catch(e) { /* ADMIN_USER_ID 未設定は無視 */ }
    return;
  }

  if (/チェックイン|到着/.test(t)) {
    processCheckin(userId, replyToken);
    return;
  }

  if (/チェックアウト|退勤/.test(t)) {
    processCheckout(userId, replyToken);
    return;
  }

  if (/応募|シフト/.test(t)) {
    const url = STATIC_APPLY_URL + '?token=' + _issueToken(userId, 'apply');
    reply('応募フォームはこちらです 📅\n' + url + '\n\nまたは希望日を直接送っても受け付けます😊');
    return;
  }

  // ── 日付抽出 → 応募 ──
  // シンプルな形式（M/D・M月D日）はregex、曖昧なパターンのみGeminiへ
  const today = formatDate(getJstNow());
  let dates;
  if (_needsGemini_(t)) {
    const geminiDates = _extractDatesGemini_(t, today);
    dates = geminiDates !== null ? geminiDates : extractDates(t, today);
  } else {
    dates = extractDates(t, today);
    // 明示的な日付表記（M/D・M月D・Y.M.D）が無いのに数字だけで日付が拾えた場合は
    // エラーコード・電話番号・金額等の誤爆の可能性が高い → AIに実際の意図を確認する
    // （例: 「H-15-02のエラーコードが出ています」→ 15,02 を日付と誤認）
    if (dates.length > 0 && !_hasStrongDateSignal_(t)) {
      const geminiDates = _extractDatesGemini_(t, today);
      dates = geminiDates !== null ? geminiDates : []; // Gemini未設定/失敗時は誤登録防止のため見送り
    }
  }
  if (dates.length === 0) {
    // 日付なし・キーワードなし → スルー（返信しない）
    return;
  }

  // 応募処理
  const allJobs = sheetToObjects(getSheet(SHEET.JOB));
  const jobs = allJobs.filter(j => String(j.active).toLowerCase() === 'true');
  const staffInfo = getAllStaff().find(s => s.line_user_id === userId);
  const job = (staffInfo && staffInfo.job_id && jobs.find(j => j.job_id === staffInfo.job_id))
    || jobs[0];
  if (!job) {
    reply('現在募集中の職種がありません。');
    return;
  }

  const added = [], skipped = [];
  dates.forEach(d => {
    const ok = appendApplied('applied', userId, d, job.job_id, 'line', '');
    if (ok) added.push(d);
    else skipped.push(d);
  });

  if (added.length === 0 && skipped.length > 0) {
    reply('応募済みの日程です。確定連絡をお待ちください 🙏');
    return;
  }

  const dateLines = added.map(d => '・' + _formatDateWithDay(d)).join('\n');
  const skipNote = skipped.length > 0
    ? '\n\n※応募済みのためスキップ：\n' + skipped.map(d => '・' + _formatDateWithDay(d)).join('\n')
    : '';

  reply(
    `以下の日程で受け付けました✅\n${dateLines}${skipNote}\n\n` +
    `職種：${job.title}\n` +
    `${_safeTime_(job.start_time)}〜${_safeTime_(job.end_time)} / ${(Number(job.wage) || 0).toLocaleString()}円＋交通費${(Number(job.transport) || 0).toLocaleString()}円\n\n` +
    `確定連絡をお待ちください🙏`
  );

  if (added.length > 0) {
    try {
      pushMessage(SECRETS.ADMIN_USER_ID,
        `📩 新しい応募\n${userId}\n${added.join(', ')}\n職種：${job.title}`
      );
    } catch(e) { /* ADMIN_USER_ID 未設定は無視 */ }
  }
}

// regex で安全に処理できるか判定
// M/D・M月D日 だけなら regex で十分。M月D（日なし）が残る場合は Gemini へ
function _needsGemini_(text) {
  const n = _normalizeMessageText_(text);
  const stripped = n
    .replace(/\d{1,2}\/\d{1,2}/g, '')    // M/D を除去
    .replace(/\d{1,2}月\d{1,2}日/g, ''); // M月D日 を除去
  return /\d月\d/.test(stripped);         // M月D（日なし）が残っていれば Gemini が必要
}

// 明示的な日付表記（M/D・M月D・Y.M.D）が含まれるかどうか
// これが無いのに extractDates が数字だけを日付として拾った場合、
// エラーコード・電話番号・金額等を誤認している可能性が高い（D-007の⑧パターン）
function _hasStrongDateSignal_(text) {
  const n = _normalizeMessageText_(text);
  return /\d{1,2}\/\d{1,2}/.test(n)
    || /\d{1,2}\s*月\s*\d{1,2}/.test(n)
    || /\d{2,4}\.\d{1,2}\.\d{1,2}/.test(n);
}

// Gemini による日付抽出（主系）
// 戻り値: 日付配列 → Gemini 成功、null → 失敗/未設定 → regex フォールバックへ
function _extractDatesGemini_(text, baseDate) {
  const apiKey = _safeProp('GEMINI_API_KEY', '');
  if (!apiKey) return null;

  const limitDate = new Date(baseDate + 'T00:00:00');
  limitDate.setDate(limitDate.getDate() + 30);
  const limitStr = formatDate(limitDate);

  const prompt =
    'LINEメッセージから勤務希望日を抽出してください。\n' +
    '基準日（今日）: ' + baseDate + '\n' +
    '対象範囲: ' + baseDate + '〜' + limitStr + '（今日から30日以内の未来日のみ）\n\n' +
    'メッセージ: 「' + String(text).replace(/[」`\\]/g, '').slice(0, 200) + '」\n\n' +
    'ルール:\n' +
    '・勤務を希望している日付のみ抽出する\n' +
    '・質問・キャンセル・確認の文脈の日付は除外\n' +
    '・YYYY-MM-DD 形式のJSON配列だけ返す（説明不要）\n' +
    '・例: ["2026-06-09","2026-06-15"]\n' +
    '・該当なし: []';

  try {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey,
      {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0, maxOutputTokens: 150 },
        }),
        muteHttpExceptions: true,
      }
    );
    const json = JSON.parse(res.getContentText());
    const raw = ((json.candidates || [])[0] || {});
    const txt = ((raw.content || {}).parts || [])[0] || {};
    const str = txt.text || '';
    Logger.log('[Gemini] date raw: ' + str);
    const m = str.match(/\[[\s\S]*?\]/);
    if (!m) return [];
    const arr = JSON.parse(m[0]);
    if (!Array.isArray(arr)) return [];
    return arr.filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d) && d >= baseDate && d <= limitStr);
  } catch(e) {
    console.error('[Gemini] extractDates failed: ' + e.message);
    return null; // → regex フォールバック
  }
}

// 日付抽出：M/D・M.D・M月D日・D日・D曜日・Y.M.D 全形式対応
// 基準日から30日以内の未来日付のみ採用（重複除去・複数件全採用）
// 質問文（「できませんか」「ますか」を含む文節）の日付は除外
function extractDates(text, baseDate, options) {
  // 質問文を文節単位で除去（「できませんか」「ますか」「でしょうか」を含む行・文を削除）
  const normalized = _normalizeMessageText_(text);
  const w0 = (options && options.keepQuestions)
    ? normalized
    : normalized.replace(/[^\n。]*(?:できませんか|ますか|でしょうか)[^\n。]*/g, ' ');

  const results = [];
  const base  = new Date(baseDate + 'T00:00:00');
  const limit = new Date(base);
  limit.setDate(base.getDate() + 30);
  const year  = base.getFullYear();
  const month = base.getMonth() + 1;
  const seen  = new Set();

  function tryAdd(mo, da) {
    if (mo < 1 || mo > 12 || da < 1 || da > 31) return false;
    const d = new Date(year, mo - 1, da);
    if (d.getMonth() !== mo - 1) return false; // 月末越え
    if (d < base || d > limit) return false;
    const s = formatDate(d);
    if (!seen.has(s)) { results.push(s); seen.add(s); }
    return true;
  }

  let w = w0;

  // ① M/D（最優先で消費 — これを先に処理しないと「4/16.21.5/1」の16.21.5をY.M.Dが誤食いする）
  w = w.replace(/(\d{1,2})\/(\d{1,2})/g, (_, mo, da) => { tryAdd(+mo, +da); return ' '; });

  // ② M月D日 or M月D（「日」省略可。例：4月4日・6月9）
  w = w.replace(/(\d{1,2})\s*月\s*(\d{1,2})(?:\s*日)?/g, (_, mo, da) => { tryAdd(+mo, +da); return ' '; });

  // ③ Y.M.D（年.月.日 例：30.3.6 → 3/6 ※M/D消費後なので誤食いなし）
  w = w.replace(/\d{2,4}\.(\d{1,2})\.(\d{1,2})/g, (_, mo, da) => { tryAdd(+mo, +da); return ' '; });

  // ④ M.D（例：5.9、5.10）を先に月日として認識。月日として無効なら日付列に戻す。
  w = w.replace(/(\d{1,2})\.(\d{1,2})/g, (_, mo, da) => {
    return tryAdd(+mo, +da) ? ' ' : (mo + '、' + da);
  });

  // ⑤ D.D / D・D / D,D（区切り付き日付列 → 複数の「D日」として認識）
  // 例：「14.18」→ 14日・18日、「25.26」→ 25日・26日
  function addDay(day) {
    for (let off = 0; off <= 30; off++) {
      const c = new Date(base);
      c.setDate(base.getDate() + off);
      if (c.getDate() === day) {
        const s = formatDate(c);
        if (!seen.has(s)) { results.push(s); seen.add(s); }
        break;
      }
    }
  }
  w = w.replace(/(\d{1,2})[.・,、]+(\d{1,2})/g, (_, d1, d2) => {
    [+d1, +d2].forEach(addDay);
    return ' ';
  });

  // ⑥ D曜日（例：24火曜日、26木 → 今月24日・26日）
  w = w.replace(/(\d{1,2})\s*[日月火水木金土]曜日?/g, (_, da) => {
    const d = new Date(year, month - 1, +da);
    if (d.getMonth() === month - 1 && d >= base && d <= limit) {
      const s = formatDate(d);
      if (!seen.has(s)) { results.push(s); seen.add(s); }
    }
    return ' ';
  });

  // ⑦ D日（例：24日 → 30日以内で該当する最初の日）
  w = w.replace(/(\d{1,2})\s*日/g, (_, da) => { addDay(+da); return ' '; });

  // ⑧ 残った数字トークン（①〜⑦で消えなかった 1〜31 の整数）→ D日として扱う
  w.replace(/[^\d]/g, ' ').trim().split(/\s+/).forEach(tok => {
    if (!/^\d{1,2}$/.test(tok)) return;
    const day = +tok;
    if (day < 1 || day > 31) return;
    addDay(day);
  });

  return results.sort(); // 日付昇順で返す
}

function _normalizeMessageText_(text) {
  return String(text || '')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[　]/g, ' ')
    .replace(/[／]/g, '/')
    .replace(/[．]/g, '.')
    .replace(/[，]/g, ',');
}

// 日付を「M/D（曜）」形式にフォーマット（例：4/4（土））
function _formatDateWithDay(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const days = ['日','月','火','水','木','金','土'];
  return (d.getMonth() + 1) + '/' + d.getDate() + '（' + days[d.getDay()] + '）';
}

function _updateDisplayNameAsync(userId) {
  try {
    const profile = getLineProfile(userId);
    if (profile && profile.displayName) {
      upsertStaff(userId, profile.displayName);
      // name が空の場合も LINE 表示名で補完
      const staff = getStaff(userId);
      if (!staff || !staff.name) {
        updateStaff(userId, { name: profile.displayName });
      }
    }
  } catch(e) { /* 無視 */ }
}

// ── チャットログ ──────────────────────────────────
function _appendChatLog(userId, direction, message) {
  if (!userId) return; // userId 空文字でのログ禁止（line.md ルール）
  try {
    const ss = getSpreadsheet();
    let sh = ss.getSheetByName(SHEET.CHAT_LOG);
    if (!sh) {
      // シートがなければ作成
      sh = ss.insertSheet(SHEET.CHAT_LOG);
      sh.getRange(1, 1, 1, 4).setValues([['timestamp', 'user_id', 'direction', 'message']]);
    }
    sh.appendRow([formatDateTime(getJstNow()), userId, direction, message]);
  } catch(e) {
    console.error('[ChatLog] ' + e.message);
  }
}

// 直近 N 件のメッセージを取得（in のみ）
function _getRecentLog(userId, n) {
  try {
    const sh = getSheet(SHEET.CHAT_LOG);
    if (sh.getLastRow() < 2) return [];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    return data
      .filter(r => String(r[1]) === userId && String(r[2]) === 'in')
      .slice(-n)
      .map(r => String(r[3]));
  } catch(e) { return []; }
}

// 名前を抽出（「田中です」「名前は田中花子」「田中に変更」etc.）
function _extractName(text) {
  const m = text.match(/(?:名前[はわ]?|私[はわ]?)\s*([^\sです。\n]{2,10})(?:です|に変更|といいます)?/) ||
            text.match(/^([^\sです。\n]{2,10})(?:です|といいます)$/);
  return m ? m[1].trim() : null;
}

// 電話番号を抽出
function _extractPhone(text) {
  const m = text.match(/0\d{1,4}[\-\s]?\d{1,4}[\-\s]?\d{3,4}/);
  return m ? m[0].replace(/[\s]/g, '-') : null;
}

// ログ含む全文から名前・電話を探して Staff を自動更新
function _extractAndUpdateProfile(userId, currentText) {
  const logs = _getRecentLog(userId, 10);
  logs.push(currentText);
  const staff = getStaff(userId);
  if (!staff) return null; // スタッフ未登録の場合はスキップ
  const updates = {};

  for (const msg of logs) {
    if (!staff.name) {
      const name = _extractName(msg);
      if (name) updates.name = name;
    }
    if (!staff.phone) {
      const phone = _extractPhone(msg);
      if (phone) updates.phone = phone;
    }
  }
  if (Object.keys(updates).length > 0) {
    updateStaff(userId, updates);
    return updates;
  }
  return null;
}

function _getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// 発行日から daysAhead 日後の23:59:59 JST までの秒数を返す
function _ttlToEndOfDayJst_(daysAhead) {
  const JST_MS = 9 * 3600 * 1000;
  const DAY_MS = 86400 * 1000;
  const nowMs = Date.now();
  const todayStartUtcMs = Math.floor((nowMs + JST_MS) / DAY_MS) * DAY_MS - JST_MS;
  const expUtcMs = todayStartUtcMs + (daysAhead + 1) * DAY_MS - 1000;
  return Math.max(60, Math.floor((expUtcMs - nowMs) / 1000));
}

function _issueToken(userId, type, extra, ttlSeconds) {
  const token = uuid();
  const props = PropertiesService.getScriptProperties();
  const data = Object.assign({ userId, type, exp: Date.now() + (ttlSeconds || 3600) * 1000 }, extra || {});
  props.setProperty('token:' + token, JSON.stringify(data));
  return token;
}

function _getTokenData(token) {
  if (!token) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('token:' + token);
  if (!raw) return null;
  try {
    const data = JSON.parse(raw);
    if (data.exp && Date.now() > data.exp) {
      props.deleteProperty('token:' + token);
      return null;
    }
    return data;
  } catch(e) {
    return null;
  }
}

function _normalizeTokenDates_(dates) {
  if (!dates) return [];
  if (typeof dates === 'string') {
    try { dates = JSON.parse(dates); } catch(e) { dates = [dates]; }
  }
  if (!Array.isArray(dates)) return [];
  const seen = new Set();
  const result = [];
  dates.forEach(d => {
    const s = safeDate(d);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || seen.has(s)) return;
    seen.add(s);
    result.push(s);
  });
  return result;
}

function _validateApplySelection_(tokenData, dates, jobId) {
  if (!tokenData || tokenData.type !== 'apply') return { ok: false, msg: 'トークン種別が不正です' };
  if (!tokenData.userId) return { ok: false, msg: 'ユーザー情報がありません' };
  const normalizedDates = _normalizeTokenDates_(dates);
  const requestedJobId = String(jobId || tokenData.jobId || 'JOB001');
  const allowedDates = _normalizeTokenDates_(tokenData.dates);
  if (allowedDates.length > 0) {
    const allowed = new Set(allowedDates);
    const outside = normalizedDates.filter(d => !allowed.has(d));
    if (outside.length > 0) return { ok: false, msg: '募集対象外の日付が含まれています' };
  }
  if (tokenData.jobId && String(tokenData.jobId) !== requestedJobId) {
    return { ok: false, msg: '募集対象外の職種です' };
  }
  if (normalizedDates.length === 0) return { ok: false, msg: '日付を選択してください' };
  return { ok: true, dates: normalizedDates, jobId: requestedJobId, userId: tokenData.userId };
}

function _handleFormRegister(body) {
  const tokenData = _getTokenData(body.token);
  if (!tokenData || tokenData.type !== 'register') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, msg: 'トークン無効' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const name = (body.name || '').trim();
  if (!name) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, msg: '名前を入力してください' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  updateStaff(tokenData.userId, { name });
  pushMessage(tokenData.userId, `✅ 登録完了！\n${name}さん、よろしくお願いします 🙏`);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _handleFormApply(body) {
  const tokenData = _getTokenData(body.token);
  if (!tokenData || tokenData.type !== 'apply') {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, msg: 'トークン無効' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const guard = _validateApplySelection_(tokenData, body.dates || [], body.job_id || 'JOB001');
  if (!guard.ok) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, msg: guard.msg }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const dates  = guard.dates;
  const jobId  = guard.jobId;
  const userId = guard.userId;
  const added  = [];
  dates.forEach(d => {
    const ok = appendApplied('applied', userId, d, jobId, 'form', '');
    if (ok) added.push(d);
  });
  if (added.length > 0) {
    try {
      pushMessage(SECRETS.ADMIN_USER_ID,
        `📩 フォーム応募\n${userId}\n${added.join(', ')}`
      );
    } catch(e) { /* 無視 */ }
  }
  const skipped = dates.filter(d => !added.includes(d));
  return ContentService.createTextOutput(JSON.stringify({ ok: added.length > 0, added, skipped }))
    .setMimeType(ContentService.MimeType.JSON);
}

// テスト用エクスポート（GAS 環境では typeof module === 'undefined' のため無視される）
if (typeof module !== 'undefined') {
  module.exports = { extractDates, _hasStrongDateSignal_, _needsGemini_, _normalizeMessageText_ };
}
