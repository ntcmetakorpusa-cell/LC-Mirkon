// ============================================================
//  ЛЦ МІРКОН — Google Apps Script Backend (Secure Version)
//  GitHub: зберігати цей файл як Code.gs
//  Паролі: зберігаються тільки в Google Sheets (хешовані)
// ============================================================

const SHEET_ORDERS   = 'Заявки';
const SHEET_USERS    = 'Користувачі';
const SHEET_SETTINGS = 'Налаштування';

// ── Точка входу ──────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('ЛЦ МІРКОН — Журнал заявок')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Хешування паролів ────────────────────────────────────────
function hashPwd(pwd) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    pwd,
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── Головний роутер ───────────────────────────────────────────
function handleRequest(action, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    switch (action) {
      case 'init':       return {ok:1};
      case 'getCfg':     return getCfg(ss);
      case 'setCfg':     return setCfg(ss, payload);
      case 'chgPwd':     return chgAdminPwd(ss, payload);
      case 'login':      return doLogin(ss, payload);
      case 'getUsers':   return getUsers(ss);
      case 'addUser':    return addUser(ss, payload);
      case 'updUser':    return updUser(ss, payload);
      case 'delUser':    return delUser(ss, payload);
      case 'getOrders':  return getOrders(ss);
      case 'addOrder':   return addOrder(ss, payload);
      case 'updOrder':   return updOrder(ss, payload);
      case 'delOrder':   return delOrder(ss, payload);
      default:           return {ok:0, err:'Unknown action: ' + action};
    }
  } catch(e) {
    return {ok:0, err: e.message};
  }
}

// ── Налаштування ──────────────────────────────────────────────
function getSheet(ss, name) {
  return ss.getSheetByName(name);
}

function getCfg(ss) {
  const sh = getSheet(ss, SHEET_SETTINGS);
  if (!sh) return {adminPwd: hashPwd('admin1234'), company: 'ЛЦ МІРКОН'};
  const data = sh.getDataRange().getValues();
  const cfg = {};
  data.forEach(r => { if (r[0]) cfg[r[0]] = r[1]; });
  return {adminPwd: cfg['adminPwd'] || hashPwd('admin1234'), company: cfg['company'] || 'ЛЦ МІРКОН'};
}

function setCfg(ss, payload) {
  const sh = getSheet(ss, SHEET_SETTINGS);
  if (!sh) return {ok:0, err:'Sheet not found'};
  const data = sh.getDataRange().getValues();
  if (payload.company !== undefined) {
    let found = false;
    for (let i=0; i<data.length; i++) {
      if (data[i][0] === 'company') { sh.getRange(i+1,2).setValue(payload.company); found=true; break; }
    }
    if (!found) sh.appendRow(['company', payload.company]);
  }
  return {ok:1};
}

function chgAdminPwd(ss, payload) {
  const cfg = getCfg(ss);
  const oldHash = hashPwd(payload.old);
  if (oldHash !== cfg.adminPwd) return {ok:0, err:'Невірний поточний пароль'};
  if (!payload.nw || payload.nw.length < 4) return {ok:0, err:'Мінімум 4 символи'};
  const sh = getSheet(ss, SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  let found = false;
  for (let i=0; i<data.length; i++) {
    if (data[i][0] === 'adminPwd') { sh.getRange(i+1,2).setValue(hashPwd(payload.nw)); found=true; break; }
  }
  if (!found) sh.appendRow(['adminPwd', hashPwd(payload.nw)]);
  return {ok:1};
}

// ── Логін ─────────────────────────────────────────────────────
function doLogin(ss, payload) {
  const cfg = getCfg(ss);
  const pwdHash = hashPwd(payload.pwd);
  
  // Admin login
  if (payload.login === 'admin') {
    if (pwdHash === cfg.adminPwd) {
      return {ok:1, user:{id:'admin', name:'Керуючий', login:'admin', role:'admin', perms:{}}};
    }
    return {ok:0, err:'Невірний пароль'};
  }
  
  // Regular users
  const sh = getSheet(ss, SHEET_USERS);
  if (!sh) return {ok:0, err:'Користувачів не знайдено'};
  const rows = sh.getDataRange().getValues().slice(1);
  for (const r of rows) {
    if (r[2] === payload.login && r[3] === pwdHash) {
      return {ok:1, user:{
        id: r[0], name: r[1], login: r[2], role: r[4],
        perms: JSON.parse(r[5] || '{}')
      }};
    }
  }
  return {ok:0, err:'Невірний логін або пароль'};
}

// ── Користувачі ───────────────────────────────────────────────
function getUsers(ss) {
  const sh = getSheet(ss, SHEET_USERS);
  if (!sh) return {ok:1, users:[]};
  const rows = sh.getDataRange().getValues().slice(1);
  const users = rows.filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], login: r[2],
    pwd: '••••••', // NEVER send real password to client
    role: r[4], perms: JSON.parse(r[5] || '{}')
  }));
  return {ok:1, users};
}

function mkUid() {
  return 'u' + new Date().getTime() + Math.floor(Math.random()*9999);
}

function addUser(ss, p) {
  const sh = getSheet(ss, SHEET_USERS);
  const rows = sh.getDataRange().getValues().slice(1);
  if (rows.some(r => r[2] === p.login)) return {ok:0, err:'Логін вже існує'};
  const id = mkUid();
  // Hash the password before saving
  sh.appendRow([id, p.name, p.login, hashPwd(p.pwd), p.role, JSON.stringify(p.perms)]);
  return {ok:1, id};
}

function updUser(ss, p) {
  const sh = getSheet(ss, SHEET_USERS);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === p.id) {
      sh.getRange(i+1,2).setValue(p.name);
      sh.getRange(i+1,3).setValue(p.login);
      // Only update password if it's not the placeholder
      if (p.pwd && p.pwd !== '••••••') {
        sh.getRange(i+1,4).setValue(hashPwd(p.pwd));
      }
      sh.getRange(i+1,5).setValue(p.role);
      sh.getRange(i+1,6).setValue(JSON.stringify(p.perms));
      return {ok:1};
    }
  }
  return {ok:0, err:'Не знайдено'};
}

function delUser(ss, p) {
  const sh = getSheet(ss, SHEET_USERS);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === p.id) { sh.deleteRow(i+1); return {ok:1}; }
  }
  return {ok:0, err:'Не знайдено'};
}

// ── Заявки ────────────────────────────────────────────────────
function ordRow2obj(r) {
  return {
    id: r[0], date: r[1], num: r[2], name: r[3], dec: r[4],
    doc: r[5], prod: r[6], client: r[7], mat: r[8], thick: r[9],
    qty: r[10], unit: r[11], done: r[12], made: r[13],
    details: r[14], comment: r[15], price: r[16], inv: r[17],
    vn: r[18], paid: r[19], ship: r[20], archived: r[21]||0,
    items: tryParse(r[22], []),
    files: tryParse(r[23], [])
  };
}
function tryParse(v, def) { try { return v ? JSON.parse(v) : def; } catch(e) { return def; } }

function getOrders(ss) {
  const sh = getSheet(ss, SHEET_ORDERS);
  if (!sh) return {ok:1, orders:[]};
  const rows = sh.getDataRange().getValues().slice(1);
  return {ok:1, orders: rows.filter(r=>r[0]).map(ordRow2obj)};
}

function mkOid() {
  return 'x' + new Date().getTime() + Math.floor(Math.random()*9999);
}

function addOrder(ss, p) {
  const sh = getSheet(ss, SHEET_ORDERS);
  const id = mkOid();
  sh.appendRow([
    id, p.date, p.num, p.name, p.dec, p.doc, p.prod, p.client,
    p.mat, p.thick, p.qty, p.unit||'шт', p.done, p.made,
    p.details, p.comment, p.price, p.inv, p.vn, p.paid, p.ship,
    p.archived||0, JSON.stringify(p.items||[]), JSON.stringify(p.files||[])
  ]);
  return {ok:1, id};
}

function updOrder(ss, p) {
  const sh = getSheet(ss, SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === p.id) {
      const cur = ordRow2obj(data[i]);
      const merged = Object.assign(cur, p);
      sh.getRange(i+1, 1, 1, 24).setValues([[
        merged.id, merged.date, merged.num, merged.name, merged.dec,
        merged.doc, merged.prod, merged.client, merged.mat, merged.thick,
        merged.qty, merged.unit||'шт', merged.done, merged.made,
        merged.details, merged.comment, merged.price, merged.inv,
        merged.vn, merged.paid, merged.ship, merged.archived||0,
        JSON.stringify(merged.items||[]), JSON.stringify(merged.files||[])
      ]]);
      return {ok:1};
    }
  }
  return {ok:0, err:'Не знайдено'};
}

function delOrder(ss, p) {
  const sh = getSheet(ss, SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === p.id) { sh.deleteRow(i+1); return {ok:1}; }
  }
  return {ok:0, err:'Не знайдено'};
}

// ── Ініціалізація таблиць ─────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Заявки
  let sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ORDERS);
    sh.appendRow([
      'ID','Дата','Номер','Назва','Десятковий №','Назва в документах',
      'До виробу','Замовник','Матеріал','Товщина мм','К-сть','Одиниця',
      'Факт','Виготовлено','Деталі','Коментар','Ціна з ПДВ','Рахунок',
      'ВН','Оплата','Відвантаження','Архів','Деталі JSON','Файли JSON'
    ]);
    sh.getRange(1,1,1,24).setFontWeight('bold').setBackground('#E8EAED');
    sh.setFrozenRows(1);
  }

  // Користувачі (паролі хешуються!)
  let ush = ss.getSheetByName(SHEET_USERS);
  if (!ush) {
    ush = ss.insertSheet(SHEET_USERS);
    ush.appendRow(['ID','Ім\'я','Логін','Пароль (MD5)','Роль','Права JSON']);
    ush.getRange(1,1,1,6).setFontWeight('bold').setBackground('#E8EAED');
    ush.setFrozenRows(1);
    // Демо-користувачі (паролі хешовані)
    const mgr = JSON.stringify({canView:1,canAdd:1,canEdit:1,canDelete:1,canFiles:1,canUpload:1,canPrice:1,canInv:1,canStatus:1,canAccounting:0,canArchive:1,canViewArchive:1});
    const usr = JSON.stringify({canView:1,canAdd:1,canEdit:1,canDelete:0,canFiles:1,canUpload:1,canPrice:0,canInv:0,canStatus:1,canAccounting:0,canArchive:0,canViewArchive:0});
    const acc = JSON.stringify({canView:1,canAdd:0,canEdit:0,canDelete:0,canFiles:1,canUpload:0,canPrice:1,canInv:1,canStatus:0,canAccounting:1,canArchive:0,canViewArchive:1});
    ush.appendRow(['u1','Іван Петренко','ivan',   hashPwd('ivan123'),  'manager',   mgr]);
    ush.appendRow(['u2','Марія Коваль', 'maria',  hashPwd('maria123'), 'user',      usr]);
    ush.appendRow(['u3','Тарас Бойко',  'taras',  hashPwd('taras123'),'viewer',    JSON.stringify({canView:1,canFiles:1})]);
    ush.appendRow(['u4','Оксана Бух',   'oksana', hashPwd('oksana123'),'accountant',acc]);
  }

  // Налаштування
  let cfg = ss.getSheetByName(SHEET_SETTINGS);
  if (!cfg) {
    cfg = ss.insertSheet(SHEET_SETTINGS);
    cfg.appendRow(['company',  'ЛЦ МІРКОН']);
    cfg.appendRow(['adminPwd', hashPwd('admin1234')]); // хеш, не пароль!
    cfg.getRange(1,1,2,1).setFontWeight('bold');
  }

  SpreadsheetApp.getUi().alert('✅ Таблиці створено!\n\nЛогіни:\n• admin / admin1234\n• ivan / ivan123\n• maria / maria123\n• oksana / oksana123\n\n⚠️ Змініть паролі після першого входу!');
}

// ── GET handler для GitHub Pages (fetch mode) ─────────────────
function doGet(e) {
  // If called from GitHub Pages with action param
  if (e && e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    const result = handleRequest(action, payload);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Normal GAS web app - serve HTML
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('ЛЦ МІРКОН — Журнал заявок')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
