/**
 * 焙香鍋貼｜線上訂單 → Google Sheet
 *
 * 使用方式：在要存放訂單的 Google 試算表開啟「擴充功能 → Apps Script」，
 * 貼上本檔案與 appsscript.json，執行一次 setup()，再部署為網頁應用程式。
 * 完整步驟請見 gas/README.md。
 */

var CONFIG = {
  OWNER_EMAIL: '',          // 新訂單通知信收件人；留空則不寄信
  SEND_CUSTOMER_EMAIL: true, // 顧客有填 Email 時寄出訂單確認信
  TIMEZONE: 'Asia/Taipei',
  SHIPPING_FEE: 60,
  FREE_SHIPPING: 500,
  SHEET_ORDERS: '訂單',
  SHEET_ITEMS: '訂單明細',
  SHEET_MENU: '菜單',
  STATUSES: ['新訂單', '製作中', '已完成', '已取消']
};

var STORES = { flagship: '旗艦店', xinyi: '信義店', banqiao: '板橋店' };
var STORE_ADDR = { flagship: '台北市中正區忠孝西路一段 100 號', xinyi: '台北市信義區松仁路 58 號', banqiao: '新北市板橋區文化路一段 88 號' };
var TYPES = { 'dine-in': '內用', 'takeout': '外帶', 'delivery': '宅配' };
var ETA = { 'dine-in': '入座後約 10 分鐘', 'takeout': '約 15 分鐘', 'delivery': '約 45–60 分鐘' };

var ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '狀態更新時間', '取餐方式', '門市／配送地址', '姓名', '手機', 'Email', '備註', '小計', '運費', '總金額', '請求ID'];
var ITEM_HEADERS = ['訂單編號', '品項代碼', '品名', '單價', '數量', '小計'];
var MENU_HEADERS = ['代碼', '品名', '單價', '供應中'];
var COL = { id: 1, created: 2, status: 3, statusAt: 4, type: 5, where: 6, name: 7, phone: 8, email: 9, note: 10, sub: 11, ship: 12, total: 13, rid: 14 };

var MENU_SEED = [
  ['g1', '招牌豬肉鍋貼', 85], ['g2', '韭香鮮肉鍋貼', 90], ['g3', '麻辣牛肉鍋貼', 110], ['g4', '起司玉米鍋貼', 100],
  ['g5', '蝦仁鮮蔬鍋貼', 120], ['g6', '酸辣泡菜鍋貼', 95],
  ['d1', '鮮肉高麗菜水餃', 90], ['d2', '玉米鮮蝦水餃', 110], ['d3', '紅油抄手', 100], ['d4', '素食三鮮水餃', 85],
  ['s1', '酸辣湯', 55], ['s2', '玉米濃湯', 50], ['s3', '古早味紅茶', 35], ['s4', '手工豆漿', 40],
  ['x1', '涼拌小黃瓜', 45], ['x2', '滷味拼盤', 75], ['x3', '麻辣口水雞', 95], ['x4', '蔥油拌麵', 60],
  ['gift1', '招牌招財鍋貼禮盒（50 入）', 680], ['gift2', '家庭號綜合冷凍組（80 入）', 980]
];

/* ============================== 初始化 ============================== */

/** 建立工作表、表頭、下拉選單與格式。可重複執行，不會覆蓋既有資料。 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  var orders = sheet_(ss, CONFIG.SHEET_ORDERS, ORDER_HEADERS);
  var items = sheet_(ss, CONFIG.SHEET_ITEMS, ITEM_HEADERS);
  var menu = sheet_(ss, CONFIG.SHEET_MENU, MENU_HEADERS);

  if (menu.getLastRow() < 2) {
    var rows = MENU_SEED.map(function (m) { return [m[0], m[1], m[2], true]; });
    menu.getRange(2, 1, rows.length, 4).setValues(rows);
    menu.getRange(2, 4, rows.length, 1).insertCheckboxes();
  }
  menu.getRange('A:A').setNumberFormat('@');

  orders.getRange(1, COL.phone, orders.getMaxRows(), 1).setNumberFormat('@');           // 保留手機開頭的 0
  orders.getRange(1, COL.created, orders.getMaxRows(), 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  orders.getRange(1, COL.statusAt, orders.getMaxRows(), 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  orders.getRange(2, COL.status, orders.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build());

  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('新訂單').setBackground('#fde7d3').setRanges([orders.getRange(2, COL.status, orders.getMaxRows() - 1, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('製作中').setBackground('#fff3b0').setRanges([orders.getRange(2, COL.status, orders.getMaxRows() - 1, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('已完成').setBackground('#d9ead3').setRanges([orders.getRange(2, COL.status, orders.getMaxRows() - 1, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('已取消').setBackground('#e6e6e6').setFontColor('#999999').setRanges([orders.getRange(2, COL.status, orders.getMaxRows() - 1, 1)]).build()
  ];
  orders.setConditionalFormatRules(rules);

  orders.setColumnWidths(1, 14, 110); orders.setColumnWidth(COL.where, 240); orders.setColumnWidth(COL.note, 200);
  items.setColumnWidths(1, 6, 110); items.setColumnWidth(3, 220);
  menu.setColumnWidths(1, 4, 110); menu.setColumnWidth(2, 220);
  ss.setActiveSheet(orders);
  toast_(ss, '初始化完成。請接著部署為網頁應用程式。');
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (cur.join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#3f2a1b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function toast_(ss, msg) { try { ss.toast(msg, '焙香鍋貼', 5); } catch (e) {} }

/* ============================== Web App ============================== */

function doGet() {
  return json_({ ok: true, service: 'baixiang-orders', time: new Date().toISOString() });
}

function doPost(e) {
  var lock = LockService.getScriptLock(), locked = false, result;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    lock.waitLock(20000); locked = true;
    result = createOrder_(body);
  } catch (err) {
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : '系統忙碌中，請稍後再試。' , code: err && err.code || 'server' });
  } finally {
    if (locked) lock.releaseLock();
  }
  if (result.isNew) { try { notify_(result.mail); } catch (mailErr) { console.error('mail failed: ' + mailErr); } }
  return json_({ ok: true, orderId: result.orderId, total: result.total, subtotal: result.subtotal, shipping: result.shipping, eta: result.eta, where: result.where, type: result.typeLabel });
}

function createOrder_(b) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = ss.getSheetByName(CONFIG.SHEET_ORDERS), itemsSh = ss.getSheetByName(CONFIG.SHEET_ITEMS), menuSh = ss.getSheetByName(CONFIG.SHEET_MENU);
  if (!orders || !itemsSh || !menuSh) throw fail_('server', '系統尚未完成初始化。');

  if (b.website) throw fail_('spam', '訂單無法送出。');                       // honeypot
  var rid = str_(b.requestId, 64);
  if (!rid) throw fail_('bad', '請求資料不完整。');

  // 重複送出（網路重試）：直接回傳已建立的訂單
  var dup = findByRid_(orders, rid);
  if (dup) return { isNew: false, orderId: dup.id, total: dup.total, subtotal: dup.sub, shipping: dup.ship, eta: ETA[dup.typeKey] || '', where: dup.where, typeLabel: dup.typeLabel };

  var typeKey = str_(b.orderType, 20);
  if (!TYPES[typeKey]) throw fail_('bad', '取餐方式不正確。');
  var name = str_(b.name, 50), phone = String(b.phone || '').replace(/[-\s]/g, ''), email = str_(b.email, 100), note = str_(b.note, 300);
  if (!name) throw fail_('bad', '請輸入姓名。');
  if (!/^09\d{8}$/.test(phone)) throw fail_('bad', '手機號碼格式不正確。');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw fail_('bad', 'Email 格式不正確。');
  var where;
  if (typeKey === 'delivery') { where = str_(b.address, 200); if (!where) throw fail_('bad', '請輸入配送地址。'); }
  else { var sid = str_(b.store, 20); if (!STORES[sid]) throw fail_('bad', '請選擇門市。'); where = STORES[sid] + '（' + STORE_ADDR[sid] + '）'; }

  // 以菜單工作表的價格為準，不信任前端送來的金額
  var menu = {}, mrows = menuSh.getLastRow() > 1 ? menuSh.getRange(2, 1, menuSh.getLastRow() - 1, 4).getValues() : [];
  mrows.forEach(function (r) { menu[String(r[0])] = { name: r[1], price: Number(r[2]), on: r[3] === true || String(r[3]).toUpperCase() === 'TRUE' }; });
  if (!Array.isArray(b.items) || !b.items.length || b.items.length > 30) throw fail_('bad', '購物車是空的。');
  var lines = [], sub = 0;
  b.items.forEach(function (it) {
    var m = menu[String(it.id)], qty = Math.floor(Number(it.qty));
    if (!m) throw fail_('bad', '找不到品項：' + String(it.id).slice(0, 20));
    if (!m.on) throw fail_('soldout', '「' + m.name + '」目前暫停供應，請從購物車移除後再送出。');
    if (!(qty >= 1 && qty <= 99)) throw fail_('bad', '品項數量不正確。');
    var line = m.price * qty; sub += line; lines.push([String(it.id), m.name, m.price, qty, line]);
  });
  var ship = (typeKey === 'delivery' && sub > 0 && sub < CONFIG.FREE_SHIPPING) ? CONFIG.SHIPPING_FEE : 0, total = sub + ship;

  var now = new Date(), orderId = nextId_(now);
  orders.appendRow([orderId, now, '新訂單', now, TYPES[typeKey], safe_(where), safe_(name), phone, safe_(email), safe_(note), sub, ship, total, rid]);
  var r0 = orders.getLastRow();
  orders.getRange(r0, COL.phone).setNumberFormat('@').setValue(phone);
  var start = itemsSh.getLastRow() + 1;
  itemsSh.getRange(start, 1, lines.length, 6).setValues(lines.map(function (l) { return [orderId, l[0], l[1], l[2], l[3], l[4]]; }));
  SpreadsheetApp.flush();

  return {
    isNew: true, orderId: orderId, total: total, subtotal: sub, shipping: ship, eta: ETA[typeKey], where: where, typeLabel: TYPES[typeKey],
    mail: { orderId: orderId, type: TYPES[typeKey], where: where, name: name, phone: phone, email: email, note: note, lines: lines, sub: sub, ship: ship, total: total, eta: ETA[typeKey] }
  };
}

/** 每日流水號：GTyyMMdd-001 */
function nextId_(now) {
  var props = PropertiesService.getScriptProperties();
  var day = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyMMdd'), n = 1;
  if (props.getProperty('SEQ_DAY') === day) n = Number(props.getProperty('SEQ_N') || 0) + 1;
  props.setProperties({ SEQ_DAY: day, SEQ_N: String(n) });
  return 'GT' + day + '-' + ('000' + n).slice(-3);
}

function findByRid_(orders, rid) {
  var last = orders.getLastRow(); if (last < 2) return null;
  var cell = orders.getRange(2, COL.rid, last - 1, 1).createTextFinder(rid).matchEntireCell(true).findNext();
  if (!cell) return null;
  var r = orders.getRange(cell.getRow(), 1, 1, ORDER_HEADERS.length).getValues()[0], label = r[COL.type - 1];
  var typeKey = Object.keys(TYPES).filter(function (k) { return TYPES[k] === label; })[0];
  return { id: r[COL.id - 1], sub: r[COL.sub - 1], ship: r[COL.ship - 1], total: r[COL.total - 1], where: r[COL.where - 1], typeLabel: label, typeKey: typeKey };
}

/* ============================== 通知信 ============================== */

function notify_(m) {
  var list = m.lines.map(function (l) { return '  ' + l[1] + ' × ' + l[3] + '　NT$ ' + l[4]; }).join('\n');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.OWNER_EMAIL) {
    MailApp.sendEmail(CONFIG.OWNER_EMAIL, '【新訂單】' + m.orderId + '｜' + m.type + '｜NT$ ' + m.total,
      '訂單編號：' + m.orderId + '\n取餐方式：' + m.type + '\n' + (m.type === '宅配' ? '配送地址：' : '門市：') + m.where +
      '\n姓名：' + m.name + '\n手機：' + m.phone + (m.email ? '\nEmail：' + m.email : '') + (m.note ? '\n備註：' + m.note : '') +
      '\n\n' + list + '\n\n小計 NT$ ' + m.sub + '　運費 NT$ ' + m.ship + '\n應付金額 NT$ ' + m.total + '\n\n試算表：' + ss.getUrl());
  }
  if (CONFIG.SEND_CUSTOMER_EMAIL && m.email) {
    MailApp.sendEmail({
      to: m.email, name: '焙香鍋貼', subject: '焙香鍋貼｜訂單確認 ' + m.orderId,
      body: m.name + ' 你好，\n\n我們已收到你的訂單，內容如下：\n\n訂單編號：' + m.orderId + '\n取餐方式：' + m.type + '\n' +
        (m.type === '宅配' ? '配送地址：' : '門市：') + m.where + '\n預計時間：' + m.eta + '\n\n' + list +
        '\n\n小計 NT$ ' + m.sub + '　運費 NT$ ' + m.ship + '\n應付金額 NT$ ' + m.total + '\n\n感謝你的訂購！'
    });
  }
}

/* ============================== 觸發條件 ============================== */

/** 簡易觸發：訂單狀態被修改時，自動更新「狀態更新時間」。 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.SHEET_ORDERS || e.range.getColumn() !== COL.status || e.range.getRow() < 2) return;
  sh.getRange(e.range.getRow(), COL.statusAt).setValue(new Date());
}

/** 執行一次，建立「每天早上 8 點寄出前一天營收摘要」的排程。 */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailySummary') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(8).inTimezone(CONFIG.TIMEZONE).create();
}

function dailySummary() {
  if (!CONFIG.OWNER_EMAIL) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
  if (!sh || sh.getLastRow() < 2) return;
  var y = new Date(); y.setDate(y.getDate() - 1);
  var key = Utilities.formatDate(y, CONFIG.TIMEZONE, 'yyyy/MM/dd'), cnt = 0, sum = 0, cancelled = 0;
  sh.getRange(2, 1, sh.getLastRow() - 1, ORDER_HEADERS.length).getValues().forEach(function (r) {
    if (Utilities.formatDate(new Date(r[COL.created - 1]), CONFIG.TIMEZONE, 'yyyy/MM/dd') !== key) return;
    if (r[COL.status - 1] === '已取消') { cancelled++; return; }
    cnt++; sum += Number(r[COL.total - 1]);
  });
  MailApp.sendEmail(CONFIG.OWNER_EMAIL, '焙香鍋貼｜' + key + ' 營收摘要', key + '\n有效訂單：' + cnt + ' 筆\n營收：NT$ ' + sum + '\n已取消：' + cancelled + ' 筆');
}

/* ============================== 工具 ============================== */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
/** 避免儲存格被當成公式執行（= + - @ 開頭）。 */
function safe_(v) { v = String(v || ''); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function fail_(code, msg) { var e = new Error(msg); e.code = code; e.userMessage = msg; return e; }
