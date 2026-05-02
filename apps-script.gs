/**
 * 荒野親子團 自然名查詢系統 — Google Apps Script 後端
 *
 * 部署步驟（請逐步操作，5 分鐘搞定）：
 *
 *  1. 開啟 https://sheets.google.com 建立一個新的 Google Sheet
 *  2. 上方選單：【擴充功能】→【Apps Script】
 *  3. 把整段這個檔案的內容，全部貼到打開的編輯器裡（覆蓋原本的 myFunction）
 *  4. 上方按💾儲存圖示，專案名稱可填「自然名 API」
 *  5. 右上角【部署】→【新增部署作業】
 *      - 類型：選【網頁應用程式】
 *      - 執行身分：【我】
 *      - 存取權：【任何人】
 *      - 按【部署】，第一次會要你授權
 *  6. 部署完成會給你一段網址，貼到網站的「☁️ 雲端同步設定」即可
 *
 * 修改後重新部署：【部署】→【管理部署作業】→ 鉛筆圖示 → 版本選【新版本】→【部署】
 *
 * --------------------------------------------------------
 * Sheet 欄位：groupId | groupName | name | addedAt | passwordHash
 * 密碼用 SHA-256 雜湊後儲存，不存明文。
 * 寫入操作（新增/刪除等）需通過密碼驗證；讀取（GET）開放任何人。
 * --------------------------------------------------------
 */

const SHEET_NAME = 'NatureNames';
const COLS = 5; // groupId, groupName, name, addedAt, passwordHash

// 超級使用者：SHA-256('泥鰍')。可繞過任何團體密碼，可清空整份資料。
// 修改此密碼：到 https://emn178.github.io/online-tools/sha256.html 算新雜湊後替換下面這串。
const SUPER_HASH = '7846d02b38c2a745a3d9bc7e24276224cda06b20cd0f6c09bc7c05768f95e9a9';

function _sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLS).setValues([['groupId', 'groupName', 'name', 'addedAt', 'passwordHash']]);
    sh.setFrozenRows(1);
  }
  // 自動補欄（從舊版 4 欄遷移）
  const headerRange = sh.getRange(1, 1, 1, Math.max(COLS, sh.getLastColumn() || 1));
  if ((sh.getLastColumn() || 0) < COLS) {
    sh.getRange(1, 5).setValue('passwordHash');
  }
  return sh;
}

function _readAll() {
  const sh = _sheet();
  const lr = sh.getLastRow();
  if (lr < 2) return { groups: {} };
  const data = sh.getRange(2, 1, lr - 1, COLS).getValues();
  const groups = {};
  for (const row of data) {
    const id = String(row[0] || '').trim();
    const gname = String(row[1] || '').trim();
    const item = String(row[2] || '').trim();
    const pwh = String(row[4] || '').trim();
    if (!id) continue;
    if (!groups[id]) groups[id] = { name: gname, names: [], passwordHash: '' };
    if (gname) groups[id].name = gname;
    if (pwh) groups[id].passwordHash = pwh;
    if (item) groups[id].names.push(item);
  }
  return { groups };
}

// 回傳時包含 passwordHash（雜湊值，不是明文密碼）
// 前端用它在本機驗證使用者輸入的密碼。實際寫入保護由 server-side _verifyAuth 把關。
function _publicView(all) {
  return all;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _verifyAuth(groupId, auth) {
  if (auth && String(auth) === SUPER_HASH) return; // 超級使用者繞過所有檢查
  const all = _readAll();
  const g = all.groups[groupId];
  if (!g) return; // 群組不存在，後續操作會自行處理
  if (!g.passwordHash) return; // 無密碼保護
  if (!auth || String(auth) !== String(g.passwordHash)) {
    throw new Error('密碼錯誤或未提供，無法執行此操作');
  }
}

function _verifySuper(auth) {
  if (!auth || String(auth) !== SUPER_HASH) {
    throw new Error('需要超級使用者權限');
  }
}

function doGet(e) {
  return _json({ ok: true, ..._publicView(_readAll()) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (err) { return _json({ ok: false, error: '伺服器忙碌中，請稍後再試' }); }
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    switch (action) {
      case 'createGroup':   _createGroup(body); break;
      case 'renameGroup':   _renameGroup(body); break;
      case 'deleteGroup':   _deleteGroup(body); break;
      case 'addNames':      _addNames(body); break;
      case 'deleteName':    _deleteName(body); break;
      case 'replaceNames':  _replaceNames(body); break;
      case 'clearGroup':    _clearGroup(body); break;
      case 'setPassword':   _setPassword(body); break;
      case 'wipeAll':       _wipeAll(body); break;
      case 'forceSetPassword': _forceSetPassword(body); break;
      default: throw new Error('未知的操作：' + action);
    }
    return _json({ ok: true, ..._publicView(_readAll()) });
  } catch (err) {
    return _json({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ================== 操作 ==================

function _createGroup(b) {
  const all = _readAll();
  if (all.groups[b.groupId]) {
    // 已存在 → 視為更新名稱/密碼，需驗證原密碼
    _verifyAuth(b.groupId, b.auth);
    if (b.groupName) _renameGroup({ groupId: b.groupId, newName: b.groupName, auth: b.auth });
    return;
  }
  // 新群組 → 直接建立（無需驗證，因為還沒密碼）
  _sheet().appendRow([b.groupId, b.groupName || '', '', new Date(), b.passwordHash || '']);
}

function _renameGroup(b) {
  _verifyAuth(b.groupId, b.auth);
  const sh = _sheet();
  const lr = sh.getLastRow();
  if (lr < 2) {
    sh.appendRow([b.groupId, b.newName, '', new Date(), '']);
    return;
  }
  const range = sh.getRange(2, 1, lr - 1, COLS);
  const data = range.getValues();
  let touched = false;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === b.groupId) {
      data[i][1] = b.newName;
      touched = true;
    }
  }
  if (touched) range.setValues(data);
  else sh.appendRow([b.groupId, b.newName, '', new Date(), '']);
}

function _deleteGroup(b) {
  _verifyAuth(b.groupId, b.auth);
  _removeRows(_sheet(), (row) => String(row[0]) === b.groupId);
}

function _addNames(b) {
  _verifyAuth(b.groupId, b.auth);
  const sh = _sheet();
  const all = _readAll();
  const g = all.groups[b.groupId];
  const existing = new Set(g ? g.names : []);
  const pwh = g ? g.passwordHash : (b.passwordHash || '');
  // 移除空白佔位列
  if (g) {
    _removeRows(sh, (row) => String(row[0]) === b.groupId && !String(row[2] || '').trim());
  }
  const toAdd = (b.names || []).map(s => String(s).trim()).filter(Boolean).filter(n => !existing.has(n));
  const now = new Date();
  const rows = toAdd.map(n => [b.groupId, b.groupName || (g ? g.name : ''), n, now, pwh]);
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, COLS).setValues(rows);
  }
  if (b.groupName && g && b.groupName !== g.name) {
    _renameGroup({ groupId: b.groupId, newName: b.groupName, auth: b.auth });
  }
}

function _deleteName(b) {
  _verifyAuth(b.groupId, b.auth);
  _removeRows(_sheet(), (row) => String(row[0]) === b.groupId && String(row[2]) === b.name, true);
  // 確保群組仍存在
  const all = _readAll();
  if (!all.groups[b.groupId]) {
    _sheet().appendRow([b.groupId, b.groupName || '', '', new Date(), '']);
  }
}

function _replaceNames(b) {
  _verifyAuth(b.groupId, b.auth);
  const all = _readAll();
  const g = all.groups[b.groupId];
  const pwh = g ? g.passwordHash : '';
  const gname = b.groupName || (g ? g.name : '');
  _removeRows(_sheet(), (row) => String(row[0]) === b.groupId);
  const sh = _sheet();
  const now = new Date();
  const list = (b.names || []).map(s => String(s).trim()).filter(Boolean);
  if (list.length === 0) {
    sh.appendRow([b.groupId, gname, '', now, pwh]);
  } else {
    const rows = list.map(n => [b.groupId, gname, n, now, pwh]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, COLS).setValues(rows);
  }
}

function _clearGroup(b) {
  _verifyAuth(b.groupId, b.auth);
  const all = _readAll();
  const g = all.groups[b.groupId];
  const pwh = g ? g.passwordHash : '';
  const gname = b.groupName || (g ? g.name : '');
  _removeRows(_sheet(), (row) => String(row[0]) === b.groupId);
  _sheet().appendRow([b.groupId, gname, '', new Date(), pwh]);
}

function _setPassword(b) {
  // 修改密碼：需提供舊密碼（如果原本有的話）
  _verifyAuth(b.groupId, b.auth);
  const sh = _sheet();
  const lr = sh.getLastRow();
  if (lr < 2) {
    sh.appendRow([b.groupId, b.groupName || '', '', new Date(), b.newPasswordHash || '']);
    return;
  }
  const range = sh.getRange(2, 1, lr - 1, COLS);
  const data = range.getValues();
  let touched = false;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === b.groupId) {
      data[i][4] = b.newPasswordHash || '';
      touched = true;
    }
  }
  if (touched) range.setValues(data);
}

// 超級使用者：清空所有資料（保留表頭）
function _wipeAll(b) {
  _verifySuper(b.auth);
  const sh = _sheet();
  const lr = sh.getLastRow();
  if (lr > 1) sh.deleteRows(2, lr - 1);
}

// 超級使用者：強制重設指定團體的密碼（不需要原密碼）
function _forceSetPassword(b) {
  _verifySuper(b.auth);
  const sh = _sheet();
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const range = sh.getRange(2, 1, lr - 1, COLS);
  const data = range.getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === b.groupId) data[i][4] = b.newPasswordHash || '';
  }
  range.setValues(data);
}

function _removeRows(sh, predicate, onlyOne) {
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const data = sh.getRange(2, 1, lr - 1, COLS).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (predicate(data[i])) {
      sh.deleteRow(i + 2);
      if (onlyOne) return;
    }
  }
}
