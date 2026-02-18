/**
 * LINE Bot + Google Sheet 冰箱庫存系統
 *
 * 重要：
 * 1) 請先在 Script Properties 設定：
 *    - LINE_CHANNEL_ACCESS_TOKEN
 *    - LINE_CHANNEL_SECRET
 *    - SPREADSHEET_ID
 * 2) 執行 initializeSheets() 建立 Users/Items/Logs 分頁與表頭
 */

const SHEET_NAMES = {
  USERS: 'Users',
  ITEMS: 'Items',
  LOGS: 'Logs',
};

const USER_HEADERS = [
  'userId',
  'displayName',
  'defaultReminderDays',
  'createdAt',
  'updatedAt',
];

const ITEM_HEADERS = [
  'itemId',
  'userId',
  'itemName',
  'quantity',
  'unit',
  'expiryDate',
  'createdAt',
  'updatedAt',
  'status',
];

const LOG_HEADERS = [
  'logId',
  'timestamp',
  'userId',
  'action',
  'itemName',
  'quantityChange',
  'note',
  'rawMessage',
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('web/index')
    .setTitle('冰箱庫存管理介面')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const signature = e.postData.headers['X-Line-Signature'] || e.postData.headers['x-line-signature'];
    if (!verifyLineSignature_(e.postData.contents, signature)) {
      return ContentService.createTextOutput('invalid signature').setMimeType(ContentService.MimeType.TEXT);
    }

    if (!body.events || !body.events.length) {
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }

    body.events.forEach((event) => {
      if (event.type !== 'message' || event.message.type !== 'text') return;
      const userId = event.source.userId;
      const input = (event.message.text || '').trim();
      const reply = handleCommand_(userId, input);
      if (event.replyToken) {
        replyLineMessage_(event.replyToken, reply);
      }
    });

    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput('error').setMimeType(ContentService.MimeType.TEXT);
  }
}

function initializeSheets() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, SHEET_NAMES.USERS, USER_HEADERS);
  ensureSheet_(ss, SHEET_NAMES.ITEMS, ITEM_HEADERS);
  ensureSheet_(ss, SHEET_NAMES.LOGS, LOG_HEADERS);
  return '初始化完成：Users / Items / Logs';
}

function handleCommand_(userId, text) {
  upsertUser_(userId);

  if (/^新增\s+/.test(text)) return addItemCommand_(userId, text);
  if (/^取出\s+/.test(text)) return takeItemCommand_(userId, text);
  if (text === '清單') return listItemsCommand_(userId);
  if (/^快過期/.test(text)) return expiringItemsCommand_(userId, text);
  if (/^設定提醒\s+/.test(text)) return setReminderCommand_(userId, text);
  return [
    '指令格式：',
    '1) 新增 品名 數量 [單位] [YYYY-MM-DD]',
    '   例：新增 牛奶 2 瓶 2026-02-20',
    '2) 取出 品名 數量',
    '   例：取出 牛奶 1',
    '3) 清單',
    '4) 快過期 [天數]',
    '   例：快過期 3',
    '5) 設定提醒 天數',
    '   例：設定提醒 5',
  ].join('\n');
}

function addItemCommand_(userId, text) {
  const parts = text.split(/\s+/);
  if (parts.length < 3) return '格式錯誤：新增 品名 數量 [單位] [YYYY-MM-DD]';

  const itemName = parts[1];
  const quantity = Number(parts[2]);
  if (Number.isNaN(quantity) || quantity <= 0) return '數量需為正數。';

  let unit = '個';
  let expiryDate = '';

  if (parts[3]) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(parts[3])) {
      expiryDate = parts[3];
    } else {
      unit = parts[3];
    }
  }
  if (parts[4]) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(parts[4])) return '日期格式需為 YYYY-MM-DD';
    expiryDate = parts[4];
  }

  const now = new Date();
  const itemSheet = getSheet_(SHEET_NAMES.ITEMS);
  const itemId = Utilities.getUuid();
  itemSheet.appendRow([
    itemId,
    userId,
    itemName,
    quantity,
    unit,
    expiryDate,
    now,
    now,
    'active',
  ]);

  addLog_(userId, 'ADD', itemName, quantity, `unit=${unit}, expiry=${expiryDate || 'N/A'}`, text);
  return `已新增：${itemName} ${quantity}${unit}${expiryDate ? `（到期：${expiryDate}）` : ''}`;
}

function takeItemCommand_(userId, text) {
  const parts = text.split(/\s+/);
  if (parts.length < 3) return '格式錯誤：取出 品名 數量';

  const itemName = parts[1];
  const takeQty = Number(parts[2]);
  if (Number.isNaN(takeQty) || takeQty <= 0) return '數量需為正數。';

  const sheet = getSheet_(SHEET_NAMES.ITEMS);
  const values = sheet.getDataRange().getValues();
  let remain = takeQty;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowUserId = row[1];
    const rowItemName = row[2];
    const rowQty = Number(row[3]);
    const status = row[8];

    if (rowUserId !== userId || rowItemName !== itemName || status !== 'active' || rowQty <= 0) continue;

    if (rowQty > remain) {
      sheet.getRange(i + 1, 4).setValue(rowQty - remain);
      sheet.getRange(i + 1, 8).setValue(new Date());
      remain = 0;
      break;
    } else {
      remain -= rowQty;
      sheet.getRange(i + 1, 4).setValue(0);
      sheet.getRange(i + 1, 8).setValue(new Date());
      sheet.getRange(i + 1, 9).setValue('empty');
      if (remain === 0) break;
    }
  }

  const taken = takeQty - remain;
  if (taken <= 0) return `找不到可取出的 ${itemName}。`;

  addLog_(userId, 'TAKE', itemName, -taken, '', text);
  return remain > 0
    ? `已取出 ${itemName} ${taken}，但庫存不足，尚缺 ${remain}。`
    : `已取出 ${itemName} ${taken}。`;
}

function listItemsCommand_(userId) {
  const sheet = getSheet_(SHEET_NAMES.ITEMS);
  const values = sheet.getDataRange().getValues();
  const map = {};

  values.slice(1).forEach((row) => {
    if (row[1] !== userId || row[8] !== 'active') return;
    const key = `${row[2]}|${row[4]}`;
    const qty = Number(row[3]) || 0;
    if (!map[key]) {
      map[key] = { name: row[2], unit: row[4], qty: 0 };
    }
    map[key].qty += qty;
  });

  const items = Object.values(map);
  if (!items.length) return '目前沒有庫存。';

  const lines = items
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((i) => `- ${i.name}：${i.qty}${i.unit}`);

  return ['目前庫存：', ...lines].join('\n');
}

function expiringItemsCommand_(userId, text) {
  const user = getUser_(userId);
  const defaultDays = user ? Number(user.defaultReminderDays || 3) : 3;
  const customMatch = text.match(/^快過期\s+(\d+)$/);
  const days = customMatch ? Number(customMatch[1]) : defaultDays;

  const today = new Date();
  const limit = new Date();
  limit.setDate(today.getDate() + days);

  const sheet = getSheet_(SHEET_NAMES.ITEMS);
  const values = sheet.getDataRange().getValues();
  const lines = [];

  values.slice(1).forEach((row) => {
    if (row[1] !== userId || row[8] !== 'active' || !row[5]) return;
    const expiry = new Date(row[5]);
    if (expiry >= startOfDay_(today) && expiry <= endOfDay_(limit)) {
      const qty = Number(row[3]) || 0;
      lines.push(`- ${row[2]}：${qty}${row[4]}（到期：${formatDate_(expiry)}）`);
    }
  });

  if (!lines.length) return `${days} 天內沒有快過期品項。`;
  return [`${days} 天內快過期：`, ...lines].join('\n');
}

function setReminderCommand_(userId, text) {
  const match = text.match(/^設定提醒\s+(\d+)$/);
  if (!match) return '格式錯誤：設定提醒 天數';

  const days = Number(match[1]);
  if (days < 1 || days > 365) return '提醒天數需介於 1~365。';

  upsertUser_(userId, { defaultReminderDays: days });
  addLog_(userId, 'SET_REMINDER', '', 0, `days=${days}`, text);
  return `已設定預設快過期提醒為 ${days} 天。`;
}

function checkAndPushExpiryReminders() {
  const userSheet = getSheet_(SHEET_NAMES.USERS);
  const users = userSheet.getDataRange().getValues().slice(1);

  users.forEach((row) => {
    const userId = row[0];
    const days = Number(row[2] || 3);
    const message = expiringItemsCommand_(userId, `快過期 ${days}`);
    if (!/沒有快過期/.test(message)) {
      pushLineMessage_(userId, message);
    }
  });
}

function verifyLineSignature_(rawBody, signature) {
  if (!signature) return false;
  const secret = getRequiredProperty_('LINE_CHANNEL_SECRET');
  const hash = Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(rawBody, secret),
  );
  return hash === signature;
}

function replyLineMessage_(replyToken, text) {
  const token = getRequiredProperty_('LINE_CHANNEL_ACCESS_TOKEN');
  const payload = {
    replyToken,
    messages: [{ type: 'text', text }],
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

function pushLineMessage_(userId, text) {
  const token = getRequiredProperty_('LINE_CHANNEL_ACCESS_TOKEN');
  const payload = {
    to: userId,
    messages: [{ type: 'text', text }],
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

function upsertUser_(userId, options) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const opts = options || {};

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId) {
      if (opts.defaultReminderDays) sheet.getRange(i + 1, 3).setValue(opts.defaultReminderDays);
      sheet.getRange(i + 1, 5).setValue(now);
      return;
    }
  }

  sheet.appendRow([
    userId,
    opts.displayName || '',
    opts.defaultReminderDays || 3,
    now,
    now,
  ]);
}

function getUser_(userId) {
  const sheet = getSheet_(SHEET_NAMES.USERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId) {
      return {
        userId: values[i][0],
        displayName: values[i][1],
        defaultReminderDays: values[i][2],
      };
    }
  }
  return null;
}

function addLog_(userId, action, itemName, quantityChange, note, rawMessage) {
  const sheet = getSheet_(SHEET_NAMES.LOGS);
  sheet.appendRow([
    Utilities.getUuid(),
    new Date(),
    userId,
    action,
    itemName,
    quantityChange,
    note || '',
    rawMessage || '',
  ]);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = headers.some((h, i) => existing[i] !== h);

  if (needsHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function getSpreadsheet_() {
  const id = getRequiredProperty_('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error(`找不到分頁：${name}，請先執行 initializeSheets()`);
  return sheet;
}

function getRequiredProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error(`缺少 Script Property: ${key}`);
  return value;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
}

function endOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
}

/**
 * ===== 管理介面（web/index.html）呼叫函式 =====
 */
function getSettings() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID') || '';
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '';
  const secret = props.getProperty('LINE_CHANNEL_SECRET') || '';
  return {
    spreadsheetId: id,
    lineTokenSet: Boolean(token),
    lineSecretSet: Boolean(secret),
  };
}

function saveSettings(data) {
  const props = PropertiesService.getScriptProperties();
  if (data.spreadsheetId) props.setProperty('SPREADSHEET_ID', data.spreadsheetId.trim());
  if (data.lineChannelAccessToken) props.setProperty('LINE_CHANNEL_ACCESS_TOKEN', data.lineChannelAccessToken.trim());
  if (data.lineChannelSecret) props.setProperty('LINE_CHANNEL_SECRET', data.lineChannelSecret.trim());
  return '設定已儲存';
}
