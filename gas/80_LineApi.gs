// 社内タイミー - 80_LineApi.gs
// LINE Messaging API ラッパー（Push / Reply / GetProfile）

const LINE_API_BASE = 'https://api.line.me/v2/bot';

function _lineHeaders() {
  return {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + SECRETS.CHANNEL_ACCESS_TOKEN,
  };
}

// テキストメッセージを push 送信
function pushMessage(userId, text) {
  if (!userId || !text) return false;
  const payload = {
    to: userId,
    messages: [{ type: 'text', text: String(text) }],
  };
  try {
    const res = UrlFetchApp.fetch(LINE_API_BASE + '/message/push', {
      method: 'post',
      headers: _lineHeaders(),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      console.error('[LineApi] push error: ' + res.getContentText());
      return false;
    }
    return true;
  } catch(e) {
    console.error('[LineApi] push exception: ' + e.message);
  }
  return false;
}

// reply トークンで返信（1回のみ使用可）
function replyMessage(replyToken, text) {
  if (!replyToken || !text) return false;
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: String(text) }],
  };
  try {
    const res = UrlFetchApp.fetch(LINE_API_BASE + '/message/reply', {
      method: 'post',
      headers: _lineHeaders(),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      console.error('[LineApi] reply error: ' + res.getContentText());
      return false;
    }
    return true;
  } catch(e) {
    console.error('[LineApi] reply exception: ' + e.message);
  }
  return false;
}

// 複数メッセージを reply 送信（最大5件）
function replyMessages(replyToken, texts) {
  if (!replyToken || !texts || texts.length === 0) return false;
  const messages = texts.slice(0, 5).map(t => ({ type: 'text', text: String(t) }));
  const payload = { replyToken: replyToken, messages: messages };
  try {
    const res = UrlFetchApp.fetch(LINE_API_BASE + '/message/reply', {
      method: 'post',
      headers: _lineHeaders(),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      console.error('[LineApi] replyMessages error: ' + res.getContentText());
      return false;
    }
    return true;
  } catch(e) {
    console.error('[LineApi] replyMessages exception: ' + e.message);
  }
  return false;
}

// ユーザーのプロフィール取得
function getLineProfile(userId) {
  try {
    const res = UrlFetchApp.fetch(LINE_API_BASE + '/profile/' + userId, {
      method: 'get',
      headers: _lineHeaders(),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) {
      return JSON.parse(res.getContentText());
    }
  } catch(e) {
    console.error('[LineApi] getProfile exception: ' + e.message);
  }
  return null;
}

// 全スタッフへ一斉送信（200ms待機必須）
function broadcastToAll(text) {
  const staffList = getStaffList_();
  let count = 0;
  staffList.forEach(s => {
    if (!s.line_user_id) return;
    if (pushMessage(s.line_user_id, text)) count++;
    Utilities.sleep(200);
  });
  Logger.log('broadcastToAll: ' + count + '件送信');
  return count;
}

function getStaffList_() {
  try {
    return sheetToObjects(getSheet(SHEET.STAFF));
  } catch(e) {
    return [];
  }
}
