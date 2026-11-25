/**
 * ค่าเช่าห้อง — Code.gs (Google Apps Script)
 * วางในโปรเจ็กต์ Apps Script ที่ผูกกับ Google Sheet ค่าเช่า (ส่วนขยาย → Apps Script)
 * ตั้งรหัสผ่าน: เมนูในชีต ค่าเช่าห้อง → ตั้งรหัสผ่านเข้าโปรแกรม
 * Deploy: เว็บแอป · Execute as: Me · Who has access: Anyone (ทุกคำสั่งต้องเข้าสู่ระบบด้วยรหัสผ่าน)
 */
var VERSION = '2026-11-25.1';

var T = {
  rooms: {
    name: 'ห้อง',
    keys: ['room', 'floor', 'tenant', 'phone', 'rent', 'other', 'otherName', 'status', 'moveIn', 'deposit', 'note', 'daily'],
    head: ['ห้อง', 'ชั้น', 'ผู้เช่า', 'เบอร์โทร', 'ค่าเช่า', 'ค่าบริการอื่น/เดือน', 'ชื่อค่าบริการอื่น', 'สถานะ', 'วันเข้าอยู่', 'เงินประกัน', 'หมายเหตุ', 'ห้องรายวัน (1 = ใช่, 0 = ไม่ใช่)']
  },
  readings: {
    name: 'มิเตอร์',
    keys: ['month', 'room', 'elec', 'water', 'elecPrev', 'waterPrev', 'elecRoll', 'waterRoll', 'at', 'note'],
    head: ['เดือน', 'ห้อง', 'ไฟ (เลขมิเตอร์)', 'น้ำ (เลขมิเตอร์)', 'ไฟ เลขก่อน (แก้เอง)', 'น้ำ เลขก่อน (แก้เอง)', 'ไฟ วนรอบ', 'น้ำ วนรอบ', 'จดเมื่อ', 'หมายเหตุ']
  },
  extras: {
    name: 'รายการเพิ่ม',
    keys: ['id', 'month', 'room', 'desc', 'amount'],
    head: ['รหัส', 'เดือน', 'ห้อง', 'รายการ', 'จำนวนเงิน']
  },
  bills: {
    name: 'บิล',
    keys: ['no', 'month', 'room', 'tenant', 'elecPrev', 'elecNow', 'elecUnits', 'elecAmt', 'waterPrev', 'waterNow', 'waterUnits', 'waterAmt',
      'rent', 'other', 'otherName', 'extras', 'extrasAmt', 'arrears', 'arrearsFrom', 'total', 'status', 'carriedTo',
      'paidAt', 'paidAmt', 'payMethod', 'issuedAt', 'dueDate', 'note'],
    head: ['เลขที่บิล', 'เดือน', 'ห้อง', 'ผู้เช่า', 'ไฟ เลขก่อน', 'ไฟ เลขหลัง', 'ไฟ หน่วย', 'ค่าไฟ', 'น้ำ เลขก่อน', 'น้ำ เลขหลัง', 'น้ำ หน่วย', 'ค่าน้ำ',
      'ค่าเช่า', 'ค่าบริการอื่น', 'ชื่อค่าบริการอื่น', 'รายการเพิ่ม', 'รวมรายการเพิ่ม', 'ยอดค้างยกมา', 'ยอดค้างจากบิล', 'ยอดรวม', 'สถานะ', 'ยกไปบิล',
      'วันที่ชำระ', 'ยอดที่ชำระ', 'วิธีชำระ', 'ออกบิลเมื่อ', 'ครบกำหนด', 'หมายเหตุ']
  },
  // การจองห้องรายวัน: วันออก = วันที่ลูกค้าออก (ไม่นับเป็นคืนที่พัก) จำนวนคืน = วันออก − วันเข้าพัก
  bookings: {
    name: 'จองรายวัน',
    keys: ['id', 'room', 'guest', 'phone', 'checkIn', 'checkOut', 'nights', 'rate', 'total', 'paid', 'status', 'note', 'createdAt'],
    head: ['รหัส', 'ห้อง', 'ชื่อผู้จอง', 'เบอร์โทร', 'วันเข้าพัก', 'วันออก', 'จำนวนคืน', 'ราคา/คืน', 'ยอดรวม', 'รับเงินแล้ว', 'สถานะ', 'หมายเหตุ', 'บันทึกเมื่อ']
  }
};

var SET_SHEET = 'ตั้งค่า';
// ค่าเริ่มต้นจากใบแจ้งหนี้เดิมของเปี่ยมสุขคอร์ท (ไฟล์ 10-26.xlsx)
var SETTINGS_DEF = [
  ['dormName', 'ชื่อหอพัก (พิมพ์บนบิล)', 'เปี่ยมสุขคอร์ท'],
  ['address', 'ที่อยู่ (พิมพ์บนบิล)', '3304 ซ.วัดบางนาใน แขวงบางนาใต้ เขตบางนา กทม. 10260'],
  ['phone', 'เบอร์ติดต่อ (พิมพ์บนบิล)', '082-559-5987'],
  ['elecRate', 'ค่าไฟ บาท/หน่วย', '7'],
  ['waterRate', 'ค่าน้ำ บาท/หน่วย', '18'],
  ['elecMin', 'ค่าไฟขั้นต่ำ บาท/เดือน (0 = ไม่มี)', '0'],
  ['waterMin', 'ค่าน้ำขั้นต่ำ บาท/เดือน (0 = ไม่มี)', '0'],
  ['dueDay', 'ชำระภายในวันที่ (ของเดือนถัดไป)', '10'],
  ['lateFee', 'ค่าปรับชำระล่าช้า บาท/วัน', '50'],
  ['payNote', 'วิธีชำระเงิน (พิมพ์บนบิล)', 'โอนเข้าบัญชีธนาคารกสิกรไทย เลขที่ 133-8-98331-6 นายจักรกฤษณ์ ณ ป้อมเพชร์ หรือสแกน QR CODE และส่งสลิปทางไลน์ @588rauwn'],
  ['footer', 'ข้อความท้ายบิล', 'กรุณาชำระเงินก่อนวันที่ 10 ของเดือนถัดไป หากชำระล่าช้า ปรับวันละ 50 บาท'],
  ['defaultRent', 'ค่าเช่าเริ่มต้น (ใช้ตอนเพิ่มห้องใหม่)', '3000'],
  ['dailyRate', 'ค่าห้องรายวัน บาท/คืน (ค่าเริ่มต้นตอนจอง)', '0']
];
var NUM_SETTINGS = ['elecRate', 'waterRate', 'elecMin', 'waterMin', 'dueDay', 'lateFee', 'defaultRent', 'dailyRate'];
// ห้องและเลขมิเตอร์เดือน ต.ค. 2569 จากไฟล์ 10-26.xlsx: [ห้อง, ชั้น, ค่าเช่า, ผู้เช่า/หมายเหตุ, ไฟก่อน, ไฟหลัง, น้ำก่อน, น้ำหลัง]
var START_MONTH = '2026-10';
var START_ROOMS = [
  ['101', '1', 2800, '', 909, 946, 3440, 3443], ['102', '1', 0, 'ร้านแม่', 8355, 8961, 0, 0], ['103', '1', 3000, '', 2840, 2995, 4339, 4349],
  ['105', '1', 0, '', 5049, 5093, 2370, 2371],
  ['201', '2', 3000, '', 3696, 3724, 2446, 2449], ['202', '2', 3000, '', 878, 907, 2172, 2173], ['203', '2', 3000, '', 3832, 3910, 3168, 3179],
  ['204', '2', 3000, '', 2820, 3043, 1820, 1834], ['205', '2', 3000, '', 4865, 4925, 1473, 1479], ['206', '2', 3000, '', 3952, 3994, 1135, 1146],
  ['301', '3', 3000, '', 4654, 4736, 2755, 2768], ['302', '3', 3000, '', 8635, 8695, 1341, 1346], ['303', '3', 3000, '', 1537, 1564, 532, 535],
  ['304', '3', 3000, '', 1321, 1336, 3466, 3469], ['305', '3', 3000, '', 1001, 1086, 1582, 1587], ['306', '3', 3000, '', 6027, 6180, 2870, 2879],
  ['401', '4', 3000, '', 7876, 7916, 1160, 1163], ['402', '4', 3000, '', 1507, 1559, 4054, 4069], ['403', '4', 3000, '', 4578, 4617, 817, 824],
  ['404', '4', 3000, '', 1823, 1862, 2813, 2822], ['405', '4', 3000, '', 3477, 3789, 1392, 1400], ['406', '4', 2700, '', 3517, 3556, 2177, 2179],
  ['ดาดฟ้า', 'ดาดฟ้า', 0, '', 6619, 6848, 1154, 1168]
];
var ST = { UNPAID: 'ค้างชำระ', PAID: 'ชำระแล้ว', CARRIED: 'ยกยอดไปบิลถัดไป', CANCEL: 'ยกเลิก' };
var VACANT = 'ว่าง', OCCUPIED = 'มีผู้เช่า';
var BK = { BOOKED: 'จองแล้ว', IN: 'เข้าพักแล้ว', OUT: 'ออกแล้ว', CANCEL: 'ยกเลิก' };

// ---------- หน้าเว็บ / เมนู ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ค่าเช่าห้อง')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ค่าเช่าห้อง')
    .addItem('ติดตั้ง / ตรวจชีต', 'menuSetup')
    .addItem('ตั้งรหัสผ่านเข้าโปรแกรม', 'menuSetPassword')
    .addToUi();
}
function menuSetup() {
  setup_();
  SpreadsheetApp.getUi().alert('พร้อมใช้งานแล้ว (เวอร์ชัน ' + VERSION + ')\n' +
    (hasPassword_() ? '' : 'ต่อไป: เมนู ค่าเช่าห้อง → ตั้งรหัสผ่านเข้าโปรแกรม\n') + 'แล้ว Deploy เป็นเว็บแอป');
}
function menuSetPassword() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('ตั้งรหัสผ่านเข้าโปรแกรม', 'พิมพ์รหัสผ่านที่ต้องการ (อย่างน้อย 4 ตัว)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var p = r.getResponseText();
  if (String(p).length < 4) { ui.alert('รหัสผ่านต้องมีอย่างน้อย 4 ตัว'); return; }
  setPassword_(p);
  ui.alert('ตั้งรหัสผ่านแล้ว — เครื่องที่เคยเข้าสู่ระบบไว้ต้องใส่รหัสใหม่');
}

// ---------- รหัสผ่าน / การเข้าใช้ ----------
function props_() { return PropertiesService.getScriptProperties(); }
function hash_(salt, pass) {
  var h = salt + '|' + pass;
  for (var i = 0; i < 300; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8)
      .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
  }
  return h;
}
function hasPassword_() { return !!props_().getProperty('APP_PASS'); }
function setPassword_(p) {
  var salt = Utilities.getUuid();
  var pr = props_(), all = pr.getProperties();
  Object.keys(all).forEach(function (k) { if (k.indexOf('S_') === 0) pr.deleteProperty(k); }); // ออกจากระบบทุกเครื่อง
  pr.setProperty('APP_PASS', salt + ':' + hash_(salt, String(p)));
}
function checkPassword_(p) {
  var v = props_().getProperty('APP_PASS');
  if (!v) return false;
  var i = v.indexOf(':');
  return hash_(v.slice(0, i), String(p || '')) === v.slice(i + 1);
}
function login(pass, remember) {
  if (!hasPassword_()) throw new Error('NO_PASSWORD: ยังไม่ได้ตั้งรหัสผ่าน — เปิด Google Sheet → เมนู ค่าเช่าห้อง → ตั้งรหัสผ่านเข้าโปรแกรม');
  var c = CacheService.getScriptCache(), n = +(c.get('LF') || 0);
  if (n >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง — รอ 10 นาทีแล้วลองใหม่');
  if (!checkPassword_(pass)) {
    c.put('LF', String(n + 1), 600);
    throw new Error('รหัสผ่านไม่ถูกต้อง' + (n + 1 >= 3 ? ' (เหลือ ' + (5 - n - 1) + ' ครั้ง)' : ''));
  }
  c.remove('LF');
  var tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var days = remember ? 30 : 1;
  props_().setProperty('S_' + tok, String(Date.now() + days * 86400000));
  return { token: tok, days: days };
}
function logout(token) { if (token) props_().deleteProperty('S_' + token); return true; }
function auth_(token) {
  var exp = token ? +(props_().getProperty('S_' + token) || 0) : 0;
  if (!exp || exp < Date.now()) {
    if (exp) props_().deleteProperty('S_' + token);
    throw new Error('AUTH: กรุณาเข้าสู่ระบบ');
  }
}
function changePassword_(oldP, newP) {
  if (!checkPassword_(oldP)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  if (String(newP || '').length < 4) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัว');
  setPassword_(newP);
  return true;
}
// ทุกคำสั่งจากหน้าเว็บผ่านตัวนี้ (ฟังก์ชันที่ลงท้าย _ เรียกจากหน้าเว็บตรงๆ ไม่ได้)
var API_ = {
  getMonth: getMonth_, getSummary: getSummary_, saveReading: saveReading_, deleteReading: deleteReading_,
  saveExtra: saveExtra_, deleteExtra: deleteExtra_, issueBills: issueBills_, cancelBill: cancelBill_,
  setPaid: setPaid_, setUnpaid: setUnpaid_, saveRoomForm: saveRoomForm_, setRoomList: setRoomList_,
  deleteRoom: deleteRoom_, saveSettings: saveSettings_, changePassword: changePassword_, roomHistory: roomHistory_, importHistory: importHistory_,
  getBookings: getBookings_, saveBooking: saveBooking_, deleteBooking: deleteBooking_
};
// เรียกจากหน้าเว็บที่โฮสต์ที่อื่น (GitHub Pages): POST text/plain {"fn":"login|logout|rpc|ping","args":[...]}
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}'), a = req.args || [];
    if (req.fn === 'login') out = { ok: true, result: login(a[0], a[1]) };
    else if (req.fn === 'logout') out = { ok: true, result: logout(a[0]) };
    else if (req.fn === 'rpc') out = { ok: true, result: rpc(a[0], a[1], a[2]) };
    else if (req.fn === 'ping') out = { ok: true, result: { version: VERSION } };
    else throw new Error('ไม่รู้จักคำสั่ง');
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function rpc(token, fn, args) {
  auth_(token);
  if (!Object.prototype.hasOwnProperty.call(API_, fn)) throw new Error('ไม่รู้จักคำสั่ง ' + fn);
  return API_[fn].apply(null, args || []);
}

// ---------- ชีต ----------
var MEMO_ = {};
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(t) {
  var ss = ss_(), sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1000, t.head.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(t) {
  if (MEMO_[t.name]) return MEMO_[t.name];
  var sh = sheet_(t), n = sh.getLastRow(), out = [];
  if (n > 1) {
    var v = sh.getRange(2, 1, n - 1, t.keys.length).getDisplayValues();
    v.forEach(function (r) {
      if (!r.some(function (x) { return String(x).trim() !== ''; })) return;
      var o = {};
      t.keys.forEach(function (k, i) { o[k] = String(r[i] == null ? '' : r[i]).trim(); });
      out.push(o);
    });
  }
  MEMO_[t.name] = out;
  return out;
}
function writeAll_(t, list) {
  var sh = sheet_(t), n = sh.getLastRow();
  var data = list.map(function (o) {
    return t.keys.map(function (k) { return o[k] == null ? '' : String(o[k]); });
  });
  if (data.length) sh.getRange(2, 1, data.length, t.keys.length).setNumberFormat('@').setValues(data);
  if (n - 1 > data.length) sh.getRange(2 + data.length, 1, n - 1 - data.length, t.keys.length).clearContent();
  MEMO_[t.name] = list;
}

// ---------- ตัวช่วย ----------
function num_(v) {
  if (v === null || v === undefined) return null;
  var s = String(v).replace(/,/g, '').trim();
  if (s === '') return null;
  var n = Number(s);
  return isFinite(n) ? n : null;
}
function r2_(n) { return Math.round(n * 100) / 100; }
function now_() {
  var d = new Date(Date.now() + 7 * 3600000).toISOString();
  return d.slice(0, 10) + ' ' + d.slice(11, 16);
}
function today_() { return now_().slice(0, 10); }
function monthOk_(m) {
  m = String(m || '');
  if (!/^\d{4}-\d{2}$/.test(m)) throw new Error('เดือนไม่ถูกต้อง');
  return m;
}
function addMonth_(m, k) {
  var y = +m.slice(0, 4), mo = +m.slice(5, 7) - 1 + k;
  y += Math.floor(mo / 12);
  mo = ((mo % 12) + 12) % 12;
  return y + '-' + ('0' + (mo + 1)).slice(-2);
}
function dueDate_(month, day) {
  var nm = addMonth_(month, 1), y = +nm.slice(0, 4), m = +nm.slice(5, 7);
  var dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  var d = Math.min(Math.max(Math.round(day) || 5, 1), dim);
  return nm + '-' + ('0' + d).slice(-2);
}
function lock_(fn) {
  var l = LockService.getScriptLock();
  if (!l.tryLock(20000)) throw new Error('ระบบกำลังบันทึกอยู่ ลองใหม่อีกครั้ง');
  try { return fn(); } finally { l.releaseLock(); }
}
function numStr_(v, label, allowNeg) {
  if (v === null || v === undefined || String(v).trim() === '') return '';
  var n = num_(v);
  if (n === null) throw new Error(label + ' ต้องเป็นตัวเลข');
  if (!allowNeg && n < 0) throw new Error(label + ' ติดลบไม่ได้');
  return String(n);
}

// ---------- ตั้งค่า ----------
function settingsSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SET_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SET_SHEET, 0);
    sh.getRange(1, 1, 100, 3).setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([['คีย์ (ห้ามแก้)', 'รายการ', 'ค่า']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function settings_() {
  if (MEMO_.settings) return MEMO_.settings;
  var sh = settingsSheet_(), n = sh.getLastRow(), have = {}, out = {};
  var v = n > 1 ? sh.getRange(2, 1, n - 1, 3).getDisplayValues() : [];
  v.forEach(function (r) { if (r[0]) have[r[0]] = r[2]; });
  var add = [];
  SETTINGS_DEF.forEach(function (d) {
    if (!(d[0] in have)) { add.push([d[0], d[1], d[2]]); have[d[0]] = d[2]; }
  });
  if (add.length) sh.getRange(n + 1, 1, add.length, 3).setNumberFormat('@').setValues(add);
  SETTINGS_DEF.forEach(function (d) { out[d[0]] = String(have[d[0]] == null ? '' : have[d[0]]).trim(); });
  NUM_SETTINGS.forEach(function (k) { out[k] = num_(out[k]) || 0; });
  MEMO_.settings = out;
  return out;
}
function saveSettings_(obj, viewMonth) {
  return lock_(function () {
    settings_();
    var sh = settingsSheet_(), n = sh.getLastRow();
    var v = sh.getRange(2, 1, n - 1, 3).getDisplayValues();
    v.forEach(function (r, i) {
      if (!obj || !Object.prototype.hasOwnProperty.call(obj, r[0])) return;
      var val = String(obj[r[0]] == null ? '' : obj[r[0]]).trim();
      if (NUM_SETTINGS.indexOf(r[0]) >= 0) val = numStr_(val, r[1]) || '0';
      sh.getRange(i + 2, 3).setNumberFormat('@').setValue(val);
    });
    MEMO_ = {};
    return getMonth_(viewMonth || today_().slice(0, 7));
  });
}

// ---------- ติดตั้ง ----------
// ชีตเดิมที่สร้างก่อนมีคอลัมน์ใหม่: เติมหัวคอลัมน์ที่ยังว่าง (ทำครั้งเดียวต่อเวอร์ชัน)
function ensureHeads_() {
  var pr = props_();
  if (pr.getProperty('HEAD_V') === VERSION) return;
  Object.keys(T).forEach(function (k) {
    var t = T[k], sh = sheet_(t), cur = sh.getRange(1, 1, 1, t.head.length).getDisplayValues()[0];
    t.head.forEach(function (h, i) {
      if (String(cur[i] || '').trim() === '') sh.getRange(1, i + 1).setNumberFormat('@').setValue(h).setFontWeight('bold');
    });
  });
  pr.setProperty('HEAD_V', VERSION);
}
function setup_() {
  var s = settings_();
  Object.keys(T).forEach(function (k) { sheet_(T[k]); });
  ensureHeads_();
  if (!rows_(T.rooms).length) {
    // ครั้งแรก: ห้องตามไฟล์เดิม + เลขมิเตอร์ ต.ค. 2569 (เลขหลังจด = เลขเดือนก่อนของ พ.ย.)
    writeAll_(T.rooms, START_ROOMS.map(function (x) {
      return { room: x[0], floor: x[1], tenant: x[3], phone: '', rent: String(x[2]), other: '', otherName: '',
        status: OCCUPIED, moveIn: '', deposit: '', note: '' };
    }));
    if (!rows_(T.readings).length) {
      writeAll_(T.readings, START_ROOMS.map(function (x) {
        return { month: START_MONTH, room: x[0], elec: String(x[5]), water: String(x[7]), elecPrev: String(x[4]), waterPrev: String(x[6]),
          elecRoll: '', waterRoll: '', at: 'นำเข้าจากไฟล์ 10-26.xlsx', note: '' };
      }));
    }
  }
}
// นำเข้าบิลย้อนหลังจากไฟล์ Excel ใบแจ้งหนี้เดิม (หน้าเว็บอ่านไฟล์แล้วส่งมาเป็นรายการ)
// item: {month, room, rent, fine, elecPrev, elecNow, elecUnits, elecAmt, waterPrev, waterNow, waterUnits, waterAmt, others:[{desc,amount}], total}
// ยอดเงินใช้ตามไฟล์ (ไม่คำนวณใหม่); opts.paid = ลงเป็นชำระแล้ว; บิลที่ชำระแล้วในระบบอยู่แล้วไม่ทับ
function importHistory_(items, opts) {
  opts = opts || {};
  return lock_(function () {
    var c = ctx_(), s = c.s, rooms = c.rooms, readings = c.readings, bills = c.bills;
    var res = { bills: 0, readings: 0, roomsAdded: [], skipped: [] };
    (items || []).forEach(function (it) {
      var month = monthOk_(it.month), room = String(it.room || '').trim();
      if (!room) return;
      var rm = rooms.filter(function (r) { return r.room === room; })[0];
      if (!rm) {
        rm = { room: room, floor: /^\d{3,}$/.test(room) ? room.charAt(0) : room, tenant: '', phone: '', rent: String(num_(it.rent) || 0),
          other: '', otherName: '', status: OCCUPIED, moveIn: '', deposit: '', note: '' };
        rooms.push(rm); res.roomsAdded.push(room);
      }
      var n = function (v) { var x = num_(v); return x === null ? '' : String(x); };
      if (n(it.elecNow) !== '' || n(it.waterNow) !== '') {
        var rec = { month: month, room: room, elec: n(it.elecNow), water: n(it.waterNow), elecPrev: n(it.elecPrev), waterPrev: n(it.waterPrev),
          elecRoll: '', waterRoll: '', at: 'นำเข้าจาก Excel ' + now_(), note: '' };
        var found = false;
        for (var i = 0; i < readings.length; i++) if (readings[i].month === month && readings[i].room === room) { readings[i] = rec; found = true; }
        if (!found) readings.push(rec);
        res.readings++;
      }
      var ex = billOf_(c, month, room);
      if (ex && (ex.status === ST.PAID || ex.status === ST.CARRIED)) { res.skipped.push(room + ' ' + month + ': มีบิล' + ex.status + 'อยู่แล้ว'); return; }
      var extras = [];
      if (num_(it.fine)) extras.push({ desc: 'ค่าปรับ', amount: num_(it.fine) });
      (it.others || []).forEach(function (o) { if (num_(o.amount)) extras.push({ desc: String(o.desc || 'อื่นๆ'), amount: num_(o.amount) }); });
      var total = num_(it.total);
      if (total === null) total = r2_((num_(it.rent) || 0) + (num_(it.elecAmt) || 0) + (num_(it.waterAmt) || 0) + extras.reduce(function (a, e) { return a + e.amount; }, 0));
      var paid = !!opts.paid;
      var bill = {
        no: ex ? ex.no : billNo_(c, month, room), month: month, room: room, tenant: rm.tenant,
        elecPrev: n(it.elecPrev), elecNow: n(it.elecNow), elecUnits: n(it.elecUnits), elecAmt: n(it.elecAmt) || '0',
        waterPrev: n(it.waterPrev), waterNow: n(it.waterNow), waterUnits: n(it.waterUnits), waterAmt: n(it.waterAmt) || '0',
        rent: n(it.rent) || '0', other: '', otherName: '',
        extras: extras.length ? JSON.stringify(extras) : '', extrasAmt: String(r2_(extras.reduce(function (a, e) { return a + e.amount; }, 0))),
        arrears: '0', arrearsFrom: '', total: String(total), status: paid ? ST.PAID : ST.UNPAID, carriedTo: '',
        paidAt: '', paidAmt: paid ? String(total) : '', payMethod: paid ? 'นำเข้าย้อนหลัง' : '',
        issuedAt: 'นำเข้าจาก Excel', dueDate: dueDate_(month, s.dueDay), note: 'นำเข้าจากไฟล์ใบแจ้งหนี้เดิม'
      };
      var idx = ex ? bills.indexOf(ex) : -1;
      if (idx >= 0) bills[idx] = bill; else bills.push(bill);
      res.bills++;
    });
    if (res.roomsAdded.length) writeAll_(T.rooms, rooms);
    writeAll_(T.readings, readings);
    writeAll_(T.bills, bills);
    return res;
  });
}
// ประวัติของห้อง: บิลและเลขมิเตอร์ 12 เดือนล่าสุด
function roomHistory_(room) {
  var c = ctx_();
  return {
    bills: c.bills.filter(function (b) { return b.room === room && b.status !== ST.CANCEL; })
      .sort(function (a, b) { return a.month < b.month ? 1 : -1; }).slice(0, 12).map(billOut_),
    readings: c.readings.filter(function (r) { return r.room === room; })
      .sort(function (a, b) { return a.month < b.month ? 1 : -1; }).slice(0, 12)
  };
}

// ---------- คำนวณ ----------
function ctx_() {
  return { s: settings_(), rooms: rows_(T.rooms), readings: rows_(T.readings), extras: rows_(T.extras), bills: rows_(T.bills) };
}
function readingOf_(c, month, room) {
  for (var i = 0; i < c.readings.length; i++) if (c.readings[i].month === month && c.readings[i].room === room) return c.readings[i];
  return null;
}
function prevOf_(c, month, room, f) {
  var best = null;
  c.readings.forEach(function (r) {
    if (r.room === room && r.month < month && num_(r[f]) !== null && (!best || r.month > best.month)) best = r;
  });
  return best ? { val: num_(best[f]), month: best.month } : null;
}
function billOf_(c, month, room) {
  for (var i = 0; i < c.bills.length; i++) {
    var b = c.bills[i];
    if (b.month === month && b.room === room && b.status !== ST.CANCEL) return b;
  }
  return null;
}
// หน่วยที่ใช้ = เลขใหม่ − เลขเก่า แล้วปัดขึ้นเป็นหน่วยเต็มเสมอ (เลขมิเตอร์มีทศนิยม 1 ตำแหน่ง: 125.1 → 126)
function ceilUnits_(x) {
  var t = Math.round(x * 10) / 10; // ตัดเศษจากการคำนวณทศนิยมของคอมพิวเตอร์
  return t < 0 ? t : Math.ceil(t);
}
function units_(prev, now, roll) {
  if (prev === null || now === null) return null;
  // มิเตอร์วนรอบ (เกิน 9,999 แล้วเริ่ม 0000) = ผลต่าง mod 10,000
  // ใช้ได้ทั้งกรณีเดือนก่อนจดเป็น 9,950 หรือจดเกินหลักเป็น 10,136 แล้วเดือนนี้มิเตอร์โชว์ 0,250
  if (roll) return ceilUnits_((((now - prev) % 10000) + 10000) % 10000);
  return ceilUnits_(now - prev);
}
function avgUnits_(c, month, room, f) {
  var list = c.bills.filter(function (b) { return b.room === room && b.month < month && b.status !== ST.CANCEL; })
    .sort(function (a, b) { return a.month < b.month ? 1 : -1; }).slice(0, 3)
    .map(function (b) { return num_(b[f]); }).filter(function (x) { return x !== null; });
  if (!list.length) return null;
  return list.reduce(function (a, b) { return a + b; }, 0) / list.length;
}
function billOut_(b) {
  var o = {};
  Object.keys(b).forEach(function (k) { o[k] = b[k]; });
  ['elecPrev', 'elecNow', 'elecUnits', 'elecAmt', 'waterPrev', 'waterNow', 'waterUnits', 'waterAmt', 'rent', 'other', 'extrasAmt', 'arrears', 'total', 'paidAmt']
    .forEach(function (k) { o[k] = num_(b[k]); });
  try { o.extras = b.extras ? JSON.parse(b.extras) : []; } catch (e) { o.extras = []; }
  return o;
}
function calc_(c, month, room) {
  var s = c.s, rd = readingOf_(c, month, room.room), errs = [], warns = [];
  var occupied = room.status !== VACANT;
  var res = {
    room: room.room, floor: room.floor, tenant: room.tenant, phone: room.phone, status: room.status, occupied: occupied, daily: isDaily_(room),
    reading: rd ? { elec: rd.elec, water: rd.water, elecPrev: rd.elecPrev, waterPrev: rd.waterPrev,
      elecRoll: rd.elecRoll === '1', waterRoll: rd.waterRoll === '1', at: rd.at, note: rd.note } : null
  };
  ['elec', 'water'].forEach(function (f) {
    var P = f + 'Prev', name = f === 'elec' ? 'ไฟ' : 'น้ำ';
    var auto = prevOf_(c, month, room.room, f);
    var manual = rd && num_(rd[P]) !== null;
    var prev = manual ? num_(rd[P]) : (auto ? auto.val : null);
    var now = rd ? num_(rd[f]) : null;
    var roll = !!(rd && rd[f + 'Roll'] === '1');
    var u = units_(prev, now, roll);
    var rate = f === 'elec' ? s.elecRate : s.waterRate, min = f === 'elec' ? s.elecMin : s.waterMin;
    var amt = (u === null || u < 0) ? null : r2_(Math.max(u * rate, occupied ? min : 0));
    var avg = avgUnits_(c, month, room.room, f + 'Units');
    res[f] = { prev: prev, prevMonth: manual ? 'แก้เอง' : (auto ? auto.month : ''), autoPrev: auto ? auto.val : null,
      now: now, roll: roll, units: u, rate: rate, min: min, amt: amt, avg: avg };
    if (now === null) { if (occupied) errs.push('ยังไม่ได้จด' + name); return; }
    if (prev === null) errs.push('ไม่มีเลข' + name + 'เดือนก่อน');
    else if (u < 0) errs.push('เลข' + name + 'น้อยกว่าเดือนก่อน (' + prev + ' → ' + now + ')');
    else {
      var extra = f === 'elec' ? 50 : 10;
      if (avg !== null && avg > 0 && u > Math.max(avg * 2, avg + extra)) warns.push(name + 'ใช้มากผิดปกติ ' + u + ' หน่วย (เฉลี่ย ' + Math.round(avg) + ')');
      if (u === 0 && occupied) warns.push(name + ' 0 หน่วย');
      if (roll) warns.push(name + 'วนรอบมิเตอร์ = ' + u + ' หน่วย');
    }
  });
  res.rent = num_(room.rent) || 0;
  res.other = num_(room.other) || 0;
  res.otherName = room.otherName || 'ค่าบริการอื่น';
  res.extras = c.extras.filter(function (e) { return e.month === month && e.room === room.room; })
    .map(function (e) { return { id: e.id, desc: e.desc, amount: num_(e.amount) || 0 }; });
  res.extrasAmt = r2_(res.extras.reduce(function (a, e) { return a + e.amount; }, 0));
  var bill = billOf_(c, month, room.room);
  res.bill = bill ? billOut_(bill) : null;
  var carried = c.bills.filter(function (b) {
    return b.room === room.room && b.month < month &&
      (b.status === ST.UNPAID || (bill && b.status === ST.CARRIED && b.carriedTo === bill.no));
  });
  res.arrearsBills = carried.map(function (b) {
    return { no: b.no, month: b.month, due: r2_((num_(b.total) || 0) - (num_(b.paidAmt) || 0)) };
  }).filter(function (x) { return x.due > 0; });
  res.arrears = r2_(res.arrearsBills.reduce(function (a, x) { return a + x.due; }, 0));
  res.errors = errs;
  res.warnings = warns;
  res.total = (errs.length || !rd) ? null : r2_(res.rent + res.other + (res.elec.amt || 0) + (res.water.amt || 0) + res.extrasAmt + res.arrears);
  res.changed = !!(bill && bill.status !== ST.PAID && res.total !== null && Math.abs(res.total - (num_(bill.total) || 0)) > 0.001);
  return res;
}

// ---------- API: อ่าน ----------
function getMonth_(month) {
  month = monthOk_(month);
  setup_();
  var c = ctx_();
  return {
    version: VERSION, month: month, today: today_(),
    settings: c.s, settingsDef: SETTINGS_DEF,
    rows: c.rooms.map(function (r) { return calc_(c, month, r); }),
    unpaid: c.bills.filter(function (b) { return b.status === ST.UNPAID; }).map(billOut_)
  };
}
function getSummary_() {
  setup_();
  var c = ctx_(), by = {};
  c.bills.forEach(function (b) {
    if (b.status === ST.CANCEL) return;
    var m = by[b.month] || (by[b.month] = { month: b.month, bills: 0, rent: 0, elec: 0, elecUnits: 0, water: 0, waterUnits: 0, other: 0, billed: 0, paid: 0, paidCount: 0, outstanding: 0 });
    var total = num_(b.total) || 0, arrears = num_(b.arrears) || 0, paid = num_(b.paidAmt) || 0;
    m.bills++;
    m.rent += num_(b.rent) || 0;
    m.elec += num_(b.elecAmt) || 0; m.elecUnits += num_(b.elecUnits) || 0;
    m.water += num_(b.waterAmt) || 0; m.waterUnits += num_(b.waterUnits) || 0;
    m.other += (num_(b.other) || 0) + (num_(b.extrasAmt) || 0);
    m.billed += total - arrears;
    m.paid += paid;
    if (b.status === ST.PAID) m.paidCount++;
    if (b.status === ST.UNPAID) m.outstanding += total - paid;
  });
  return Object.keys(by).sort().reverse().map(function (k) {
    var m = by[k];
    ['rent', 'elec', 'water', 'other', 'billed', 'paid', 'outstanding', 'elecUnits', 'waterUnits'].forEach(function (x) { m[x] = r2_(m[x]); });
    return m;
  });
}

// ---------- API: มิเตอร์ ----------
function assertEditable_(c, month, room) {
  var b = billOf_(c, month, room);
  if (b && b.status === ST.PAID) throw new Error('ห้อง ' + room + ' ชำระเงินเดือนนี้แล้ว — ยกเลิกการรับเงินก่อนจึงจะแก้ได้');
  if (b && b.status === ST.CARRIED) throw new Error('บิลห้อง ' + room + ' เดือนนี้ยกยอดไปบิล ' + b.carriedTo + ' แล้ว — ยกเลิกบิลนั้นก่อน');
}
function saveReading_(o) {
  return lock_(function () {
    var month = monthOk_(o.month), room = String(o.room || '');
    var c = ctx_();
    var roomObj = c.rooms.filter(function (r) { return r.room === room; })[0];
    if (!roomObj) throw new Error('ไม่พบห้อง ' + room);
    assertEditable_(c, month, room);
    var rec = {
      month: month, room: room,
      elec: numStr_(o.elec, 'เลขมิเตอร์ไฟ'), water: numStr_(o.water, 'เลขมิเตอร์น้ำ'),
      elecPrev: numStr_(o.elecPrev, 'เลขไฟเดือนก่อน'), waterPrev: numStr_(o.waterPrev, 'เลขน้ำเดือนก่อน'),
      elecRoll: o.elecRoll ? '1' : '', waterRoll: o.waterRoll ? '1' : '',
      at: now_(), note: String(o.note || '').trim()
    };
    if (rec.elec === '' && rec.water === '') throw new Error('ใส่เลขมิเตอร์อย่างน้อย 1 อย่าง');
    var list = c.readings.slice(), done = false;
    for (var i = 0; i < list.length; i++) if (list[i].month === month && list[i].room === room) { list[i] = rec; done = true; }
    if (!done) list.push(rec);
    writeAll_(T.readings, list);
    return calc_(ctx_(), month, roomObj);
  });
}
function deleteReading_(month, room) {
  return lock_(function () {
    month = monthOk_(month);
    var c = ctx_();
    assertEditable_(c, month, room);
    writeAll_(T.readings, c.readings.filter(function (r) { return !(r.month === month && r.room === room); }));
    return getMonth_(month);
  });
}

// ---------- API: รายการเพิ่ม ----------
function saveExtra_(o) {
  return lock_(function () {
    var month = monthOk_(o.month), room = String(o.room || '');
    var c = ctx_();
    if (!c.rooms.some(function (r) { return r.room === room; })) throw new Error('ไม่พบห้อง ' + room);
    assertEditable_(c, month, room);
    var desc = String(o.desc || '').trim();
    if (!desc) throw new Error('ใส่ชื่อรายการ');
    var amount = numStr_(o.amount, 'จำนวนเงิน', true);
    if (amount === '' || +amount === 0) throw new Error('ใส่จำนวนเงิน (ส่วนลดใส่ติดลบ เช่น -500)');
    var list = c.extras.slice();
    var id = o.id || ('X' + Date.now().toString(36) + Math.floor(Math.random() * 1e4));
    var rec = { id: id, month: month, room: room, desc: desc, amount: amount }, done = false;
    for (var i = 0; i < list.length; i++) if (list[i].id === id) { list[i] = rec; done = true; }
    if (!done) list.push(rec);
    writeAll_(T.extras, list);
    return getMonth_(month);
  });
}
function deleteExtra_(id, month) {
  return lock_(function () {
    var c = ctx_();
    var e = c.extras.filter(function (x) { return x.id === id; })[0];
    if (e) {
      assertEditable_(c, e.month, e.room);
      writeAll_(T.extras, c.extras.filter(function (x) { return x.id !== id; }));
    }
    return getMonth_(month || (e && e.month));
  });
}

// ---------- API: บิล ----------
function billNo_(c, month, room) {
  var yy = String(+month.slice(0, 4) + 543).slice(-2), base = yy + month.slice(5, 7) + '-' + room, no = base, k = 2;
  var used = {};
  c.bills.forEach(function (b) { used[b.no] = 1; });
  while (used[no]) no = base + '-' + (k++);
  return no;
}
function issueBills_(month, only) {
  return lock_(function () {
    month = monthOk_(month);
    var c = ctx_(), s = c.s, bills = c.bills, issued = 0, skipped = [];
    c.rooms.forEach(function (room) {
      if (only && only.length && only.indexOf(room.room) < 0) return;
      var r = calc_(c, month, room);
      if (!r.occupied) return;
      if (r.errors.length) { skipped.push({ room: room.room, reason: r.errors.join(', ') }); return; }
      var ex = billOf_(c, month, room.room);
      if (ex && ex.status === ST.PAID) { skipped.push({ room: room.room, reason: 'ชำระแล้ว ไม่ออกบิลใหม่' }); return; }
      if (ex && ex.status === ST.CARRIED) { skipped.push({ room: room.room, reason: 'ยกยอดไปบิล ' + ex.carriedTo + ' แล้ว' }); return; }
      if (ex && !r.changed) return;
      var no = ex ? ex.no : billNo_(c, month, room.room);
      bills.forEach(function (b) {
        if (b.status === ST.CARRIED && b.carriedTo === no) { b.status = ST.UNPAID; b.carriedTo = ''; }
      });
      r.arrearsBills.forEach(function (a) {
        bills.forEach(function (b) {
          if (b.no === a.no && b.status === ST.UNPAID) { b.status = ST.CARRIED; b.carriedTo = no; }
        });
      });
      var paidAmt = ex ? (num_(ex.paidAmt) || 0) : 0;
      var rec = {
        no: no, month: month, room: room.room, tenant: room.tenant,
        elecPrev: r.elec.prev, elecNow: r.elec.now, elecUnits: r.elec.units, elecAmt: r.elec.amt === null ? 0 : r.elec.amt,
        waterPrev: r.water.prev, waterNow: r.water.now, waterUnits: r.water.units, waterAmt: r.water.amt === null ? 0 : r.water.amt,
        rent: r.rent, other: r.other, otherName: r.other ? r.otherName : '',
        extras: r.extras.length ? JSON.stringify(r.extras.map(function (e) { return { desc: e.desc, amount: e.amount }; })) : '',
        extrasAmt: r.extrasAmt, arrears: r.arrears, arrearsFrom: r.arrearsBills.map(function (a) { return a.no; }).join(', '),
        total: r.total, status: (paidAmt && paidAmt >= r.total - 0.001) ? ST.PAID : ST.UNPAID, carriedTo: '',
        paidAt: ex ? ex.paidAt : '', paidAmt: paidAmt ? paidAmt : '', payMethod: ex ? ex.payMethod : '',
        issuedAt: now_(), dueDate: dueDate_(month, s.dueDay), note: ex ? ex.note : ''
      };
      ['elecPrev', 'elecNow', 'elecUnits', 'waterPrev', 'waterNow', 'waterUnits'].forEach(function (k) { if (rec[k] === null) rec[k] = ''; });
      var idx = ex ? bills.indexOf(ex) : -1;
      if (idx >= 0) bills[idx] = rec; else bills.push(rec);
      issued++;
    });
    writeAll_(T.bills, bills);
    return { issued: issued, skipped: skipped, data: getMonth_(month) };
  });
}
function findBill_(c, no) {
  var b = c.bills.filter(function (x) { return x.no === no && x.status !== ST.CANCEL; })[0];
  if (!b) throw new Error('ไม่พบบิล ' + no);
  return b;
}
function cancelBill_(no, viewMonth) {
  return lock_(function () {
    var c = ctx_(), b = findBill_(c, no);
    if (b.status === ST.PAID || num_(b.paidAmt)) throw new Error('บิลนี้มีการรับเงินแล้ว — ยกเลิกการรับเงินก่อน');
    if (b.status === ST.CARRIED) throw new Error('บิลนี้ยกยอดไปบิล ' + b.carriedTo + ' แล้ว — ยกเลิกบิลนั้นก่อน');
    c.bills.forEach(function (x) { if (x.status === ST.CARRIED && x.carriedTo === no) { x.status = ST.UNPAID; x.carriedTo = ''; } });
    b.status = ST.CANCEL;
    writeAll_(T.bills, c.bills);
    return getMonth_(viewMonth || b.month);
  });
}
function setPaid_(no, o, viewMonth) {
  return lock_(function () {
    o = o || {};
    var c = ctx_(), b = findBill_(c, no);
    if (b.status === ST.CARRIED) throw new Error('ยอดบิลนี้ยกไปรวมในบิล ' + b.carriedTo + ' แล้ว — รับเงินที่บิลนั้น');
    var total = num_(b.total) || 0;
    var amt = num_(o.paidAmt);
    if (amt === null) amt = total;
    if (amt <= 0) throw new Error('ยอดที่รับต้องมากกว่า 0');
    var at = String(o.paidAt || today_());
    if (!/^\d{4}-\d{2}-\d{2}$/.test(at)) throw new Error('วันที่ชำระไม่ถูกต้อง');
    b.paidAmt = String(r2_(amt));
    b.paidAt = at;
    b.payMethod = String(o.payMethod || 'เงินสด');
    b.status = amt >= total - 0.001 ? ST.PAID : ST.UNPAID;
    writeAll_(T.bills, c.bills);
    return getMonth_(viewMonth || b.month);
  });
}
function setUnpaid_(no, viewMonth) {
  return lock_(function () {
    var c = ctx_(), b = findBill_(c, no);
    if (b.status === ST.CARRIED) throw new Error('บิลนี้ยกยอดไปบิล ' + b.carriedTo + ' แล้ว');
    b.paidAmt = ''; b.paidAt = ''; b.payMethod = '';
    b.status = ST.UNPAID;
    writeAll_(T.bills, c.bills);
    return getMonth_(viewMonth || b.month);
  });
}

// ---------- API: ห้อง ----------
function roomHasData_(c, room) {
  return c.readings.some(function (r) { return r.room === room; }) || c.bills.some(function (b) { return b.room === room; }) ||
    c.extras.some(function (e) { return e.room === room; });
}
function saveRoom_(o, orig, viewMonth) {
  return lock_(function () {
    var c = ctx_(), list = c.rooms.slice();
    var room = String(o.room || '').trim();
    if (!room) throw new Error('ใส่เลขห้อง');
    if (orig && orig !== room && roomHasData_(c, orig)) throw new Error('เปลี่ยนเลขห้องไม่ได้ เพราะห้องนี้มีมิเตอร์/บิลแล้ว');
    if (list.some(function (r) { return r.room === room && r.room !== orig; })) throw new Error('มีห้อง ' + room + ' อยู่แล้ว');
    var rec = {
      room: room, floor: String(o.floor || '').trim(), tenant: String(o.tenant || '').trim(), phone: String(o.phone || '').trim(),
      rent: numStr_(o.rent, 'ค่าเช่า') || '0', other: numStr_(o.other, 'ค่าบริการอื่น'), otherName: String(o.otherName || '').trim(),
      status: o.status === VACANT ? VACANT : OCCUPIED, moveIn: String(o.moveIn || '').trim(),
      deposit: numStr_(o.deposit, 'เงินประกัน'), note: String(o.note || '').trim(),
      daily: o.daily === true || o.daily === '1' ? '1' : (o.daily === false || o.daily === '0' ? '0' : '')
    };
    var idx = -1;
    for (var i = 0; i < list.length; i++) if (list[i].room === (orig || room)) idx = i;
    if (idx >= 0) list[idx] = rec; else list.push(rec);
    writeAll_(T.rooms, list);
    return getMonth_(viewMonth || today_().slice(0, 7));
  });
}
// แก้จากหน้าเว็บ: คงช่องที่หน้าเว็บไม่ได้แสดง (วันเข้าอยู่ เงินประกัน หมายเหตุ / ห้องรายวัน ถ้าไม่ได้ส่งมา)
function saveRoomForm_(o, orig, viewMonth) {
  var old = rows_(T.rooms).filter(function (r) { return r.room === (orig || ''); })[0] || {};
  var x = {};
  ['moveIn', 'deposit', 'note', 'daily'].forEach(function (k) { x[k] = old[k] || ''; });
  Object.keys(o || {}).forEach(function (k) { x[k] = o[k]; });
  return saveRoom_(x, orig, viewMonth);
}
// ตั้งเลขห้องทั้งหมดครั้งเดียว: list = [{room, floor}] ตามลำดับ
// ห้องเดิมที่มีชื่อตรงกันเก็บข้อมูลไว้, ห้องเดิมที่ไม่อยู่ในรายการ: ไม่มีข้อมูล = ลบ / มีข้อมูล = เก็บไว้ท้ายรายการ
function setRoomList_(list, viewMonth) {
  return lock_(function () {
    var c = ctx_(), s = c.s, seen = {}, out = [], kept = [];
    (list || []).forEach(function (x) {
      var room = String(x.room || '').trim();
      if (!room || seen[room]) return;
      seen[room] = 1;
      var old = c.rooms.filter(function (r) { return r.room === room; })[0];
      if (old) { if (x.floor) old.floor = String(x.floor); out.push(old); }
      else out.push({ room: room, floor: String(x.floor || ''), tenant: '', phone: '', rent: String(s.defaultRent || 0), other: '', otherName: '',
        status: OCCUPIED, moveIn: '', deposit: '', note: '' });
    });
    if (!out.length) throw new Error('ไม่มีเลขห้อง');
    c.rooms.forEach(function (r) { if (!seen[r.room] && roomHasData_(c, r.room)) kept.push(r.room), out.push(r); });
    writeAll_(T.rooms, out);
    var d = getMonth_(viewMonth || today_().slice(0, 7));
    d.kept = kept;
    return d;
  });
}
function deleteRoom_(room, viewMonth) {
  return lock_(function () {
    var c = ctx_();
    if (roomHasData_(c, room)) throw new Error('ห้องนี้มีมิเตอร์/บิลแล้ว ลบไม่ได้ — เปลี่ยนสถานะเป็น "ว่าง" แทน');
    writeAll_(T.rooms, c.rooms.filter(function (r) { return r.room !== room; }));
    return getMonth_(viewMonth || today_().slice(0, 7));
  });
}

// ---------- API: จองห้องรายวัน ----------
// ห้องรายวัน = ติ๊ก "ห้องรายวัน" ในหน้าแก้ไขห้อง (คอลัมน์ ห้องรายวัน = 1)
// ห้องที่ยังไม่เคยตั้งค่า แต่ชื่อผู้เช่า/หมายเหตุมีคำว่า "รายวัน" ถือเป็นห้องรายวันด้วย (ตั้ง 0 เพื่อปิด)
function isDaily_(room) {
  if (room.daily === '1') return true;
  if (room.daily === '0') return false;
  return /รายวัน/.test(String(room.tenant || '') + ' ' + String(room.note || ''));
}
function dateOk_(s, label) {
  s = String(s || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || isNaN(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)))) throw new Error(label + ' ไม่ถูกต้อง');
  return s;
}
function dayNum_(s) { return Math.round(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) / 86400000); }
function bookOut_(b) {
  var o = {};
  Object.keys(b).forEach(function (k) { o[k] = b[k]; });
  ['nights', 'rate', 'total', 'paid'].forEach(function (k) { o[k] = num_(b[k]) || 0; });
  return o;
}
function getBookings_(month) {
  month = monthOk_(month);
  setup_();
  var c = ctx_(), s = c.s, list = rows_(T.bookings), today = today_();
  var from = month + '-01', to = addMonth_(month, 1) + '-01';
  return {
    month: month, today: today, dailyRate: s.dailyRate,
    rooms: c.rooms.filter(isDaily_).map(function (r) { return { room: r.room, floor: r.floor }; }),
    list: list.filter(function (b) { return b.checkIn < to && b.checkOut > from; })
      .sort(function (a, b) { return a.checkIn < b.checkIn ? -1 : a.checkIn > b.checkIn ? 1 : (a.room < b.room ? -1 : 1); }).map(bookOut_),
    // ที่จะเข้าพัก/กำลังพัก (ไม่ขึ้นกับเดือนที่เลือก)
    upcoming: list.filter(function (b) { return b.status !== BK.CANCEL && b.status !== BK.OUT && b.checkOut >= today; })
      .sort(function (a, b) { return a.checkIn < b.checkIn ? -1 : 1; }).slice(0, 30).map(bookOut_)
  };
}
function saveBooking_(o, viewMonth) {
  return lock_(function () {
    o = o || {};
    var c = ctx_(), list = rows_(T.bookings).slice();
    var old = o.id ? list.filter(function (b) { return b.id === o.id; })[0] : null;
    if (o.id && !old) throw new Error('ไม่พบการจองนี้ (อาจถูกลบไปแล้ว)');
    var room = String(o.room || '').trim();
    var rm = c.rooms.filter(function (r) { return r.room === room; })[0];
    if (!rm) throw new Error('เลือกห้อง');
    if (!isDaily_(rm) && !(old && old.room === room)) throw new Error('ห้อง ' + room + ' ไม่ได้ตั้งเป็นห้องรายวัน');
    var guest = String(o.guest || '').trim();
    if (!guest) throw new Error('ใส่ชื่อผู้จอง');
    var ci = dateOk_(o.checkIn, 'วันเข้าพัก'), co = dateOk_(o.checkOut, 'วันออก');
    var nights = dayNum_(co) - dayNum_(ci);
    if (nights < 1) throw new Error('วันออกต้องอยู่หลังวันเข้าพักอย่างน้อย 1 วัน');
    if (nights > 366) throw new Error('จองได้ไม่เกิน 366 คืน');
    var status = [BK.BOOKED, BK.IN, BK.OUT, BK.CANCEL].indexOf(o.status) >= 0 ? o.status : BK.BOOKED;
    if (status !== BK.CANCEL) {
      var clash = list.filter(function (b) {
        return b.room === room && b.id !== o.id && b.status !== BK.CANCEL && b.checkIn < co && b.checkOut > ci;
      })[0];
      if (clash) throw new Error('ห้อง ' + room + ' มีการจองซ้อนอยู่แล้ว: ' + clash.guest + ' (' + clash.checkIn + ' ถึง ' + clash.checkOut + ')');
    }
    var rate = numStr_(o.rate, 'ราคา/คืน') || '0';
    var total = numStr_(o.total, 'ยอดรวม');
    if (total === '') total = String(r2_(+rate * nights));
    var rec = {
      id: old ? old.id : ('B' + Date.now().toString(36) + Math.floor(Math.random() * 1e4)),
      room: room, guest: guest, phone: String(o.phone || '').trim(), checkIn: ci, checkOut: co, nights: String(nights),
      rate: rate, total: total, paid: numStr_(o.paid, 'รับเงินแล้ว') || '0', status: status,
      note: String(o.note || '').trim(), createdAt: old ? old.createdAt : now_()
    };
    if (old) list[list.indexOf(old)] = rec; else list.push(rec);
    writeAll_(T.bookings, list);
    var d = getBookings_(viewMonth ? monthOk_(viewMonth) : ci.slice(0, 7));
    d.saved = bookOut_(rec);
    return d;
  });
}
function deleteBooking_(id, viewMonth) {
  return lock_(function () {
    var list = rows_(T.bookings);
    writeAll_(T.bookings, list.filter(function (b) { return b.id !== id; }));
    return getBookings_(viewMonth || today_().slice(0, 7));
  });
}
