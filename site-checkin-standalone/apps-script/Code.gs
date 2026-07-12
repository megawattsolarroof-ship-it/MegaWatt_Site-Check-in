/**
 * ============================================================================
 *  Code.gs — Backend ระบบบันทึกเวลาทำงาน + เช็กอินหน้างาน (standalone)
 * ----------------------------------------------------------------------------
 *  วิธีติดตั้ง:
 *  1) สร้าง Google Sheet เปล่าใหม่ 1 ไฟล์ (ชื่ออะไรก็ได้)
 *  2) เปิดเมนู Extensions → Apps Script → ลบโค้ดเดิม แล้ววางไฟล์นี้ทั้งไฟล์
 *  3) (แนะนำ) Project Settings → Script Properties → เพิ่ม
 *       APP_SECRET = รหัสลับอะไรก็ได้ เช่น mycompany-2569-xyz
 *     ถ้าตั้งไว้ ทุกคำขอต้องแนบรหัสนี้ (กรอกในหน้า "ตั้งค่าระบบ" ของเว็บ)
 *     ถ้าเว้นว่าง = เปิดให้ทุกคนที่รู้ URL ใช้งานได้
 *
 *     ADMIN_SECRET = solarroof1 (ต้องตรงกับ ADMIN_SECRET_KEY ใน config.html)
 *     ใช้ป้องกันการ "แก้ตั้งค่า GPS" — คนที่ไม่รู้รหัสนี้แก้ค่าไม่ได้
 *     ถ้าเว้นว่าง = ใครก็แก้ตั้งค่าได้ (ไม่แนะนำ)
 *  4) Deploy → New deployment → Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  5) คัดลอก URL ที่ได้ (ลงท้าย /exec) ไปวางใน js/api-config.js ของเว็บ
 *     หรือกรอกในหน้า config.html
 *
 *  ชีตจะถูกสร้างให้อัตโนมัติเมื่อใช้งานครั้งแรก:
 *    - Faces        : ฐานข้อมูลใบหน้าพนักงาน
 *    - Attendance   : บันทึกเวลาเข้า-ออกงาน
 *    - Site_CheckIn : บันทึกเช็กอินหน้างาน
 *    - Config       : พิกัด GPS + รัศมีที่อนุญาต
 * ============================================================================
 */

var TZ = 'Asia/Bangkok';

// ---------------------------------------------------------------- entry points

function doGet(e) {
  if (!checkKey_(e, null)) return json_({ error: 'unauthorized', message: 'App Key ไม่ถูกต้อง — กรอกในหน้าตั้งค่าระบบ' });

  var action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'getConfig')            return json_(getConfig_());
    if (action === 'getKnownFaces')        return json_(getKnownFaces_());
    if (action === 'getTodayAttendance')   return json_(getTodayAttendance_());
    if (action === 'getTodaySiteCheckin')  return json_(getTodaySiteCheckin_());
    if (action === 'checkAdmin')           return json_({ ok: checkAdmin_(e.parameter.adminKey) });
    if (action === 'debugToday')           return json_(debugToday_());
    return json_({ error: 'unknown_action', message: 'ไม่รู้จัก action: ' + action });
  } catch (err) {
    return json_({ error: 'server_error', message: String(err) });
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}

  if (!checkKey_(e, body)) return json_({ error: 'unauthorized', message: 'App Key ไม่ถูกต้อง — กรอกในหน้าตั้งค่าระบบ' });

  try {
    if (body.action === 'registerUser')  return json_(registerUser_(body));
    if (body.action === 'logAttendance') return json_(logAttendance_(body));
    // saveConfig ต้องแนบรหัสหน้าตั้งค่า (adminKey) ให้ตรงกับ ADMIN_SECRET
    if (body.action === 'saveConfig') {
      if (!checkAdmin_(body.adminKey)) return json_({ error: 'admin_required', message: 'รหัสเข้าหน้าตั้งค่าไม่ถูกต้อง' });
      return json_(saveConfig_(body));
    }
    return json_({ error: 'unknown_action', message: 'ไม่รู้จัก action: ' + (body.action || '') });
  } catch (err) {
    return json_({ error: 'server_error', message: String(err) });
  }
}

// ---------------------------------------------------------------- security

function checkKey_(e, body) {
  var secret = PropertiesService.getScriptProperties().getProperty('APP_SECRET') || '';
  if (!secret) return true; // ไม่ได้ตั้งรหัส = เปิดใช้ได้เลย
  var key = (e && e.parameter && e.parameter.key) || (body && body.key) || '';
  return String(key) === secret;
}

function checkAdmin_(key) {
  var admin = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || '';
  if (!admin) return true; // ไม่ได้ตั้งรหัสผู้ดูแล = ไม่บังคับ (ไม่แนะนำ)
  return String(key || '') === admin;
}

// ---------------------------------------------------------------- sheets

function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function facesSheet_()      { return ensureSheet_('Faces',        ['ชื่อ', 'Descriptor', 'วันที่ลงทะเบียน']); }
function attendanceSheet_() { return ensureSheet_('Attendance',   ['วันที่', 'ชื่อ', 'เวลาเข้า', 'เวลาออก', 'Lat', 'Lng', 'Google Map Link', 'หมายเหตุ']); }
function siteSheet_()       { return ensureSheet_('Site_CheckIn', ['วันที่', 'เวลา', 'ชื่อ', 'Lat', 'Lng', 'Google Map Link', 'หมายเหตุ']); }
function configSheet_()     { return ensureSheet_('Config',       ['key', 'value']); }

function today_()   { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowTime_() { return Utilities.formatDate(new Date(), TZ, 'HH:mm:ss'); }
function mapLink_(lat, lng) {
  if (!lat || !lng) return '';
  return 'https://www.google.com/maps?q=' + lat + ',' + lng;
}

// เช็กว่าเป็นค่าแบบวันที่หรือไม่ — ห้ามใช้ instanceof Date เพราะใน runtime บางแบบให้ผลผิด
function isDate_(v) {
  return v && typeof v.getFullYear === 'function' && typeof v.getTime === 'function';
}

// อ่านค่าเซลล์ "วันที่" ให้เป็น yyyy-MM-dd (ค.ศ.) เสมอ
// รองรับทั้ง Date object, ข้อความ วัน/เดือน/ปี, ปี พ.ศ. และข้อความวันที่รูปแบบอื่น ๆ
function cellDate_(v) {
  if (isDate_(v)) {
    var y = v.getFullYear();
    if (y > 2400) { // ปี พ.ศ. → แปลงเป็น ค.ศ.
      v = new Date(v.getTime());
      v.setFullYear(y - 543);
    }
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // ข้อความ วัน/เดือน/ปี (มีเวลาต่อท้ายก็ได้)
  if (m) {
    var yr = Number(m[3]);
    if (yr > 2400) yr -= 543; // ปี พ.ศ. → ค.ศ.
    return yr + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  var t = new Date(s); // เผื่อเป็นข้อความรูปแบบอื่น เช่น "Sun Jul 12 2026 11:52:02 GMT+0700"
  if (!isNaN(t.getTime())) return Utilities.formatDate(t, TZ, 'yyyy-MM-dd');
  return s;
}

// เพิ่มแถวใหม่โดยเขียน "วันที่" เป็น Date จริง + บังคับรูปแบบ วัน/เดือน/ปี ให้เหมือนกันทุกแถว
function appendRowWithDate_(sheet, values) {
  sheet.appendRow(values);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('d/M/yyyy');
}

// อ่านค่าเซลล์ "เวลา" ให้เป็น HH:mm:ss เสมอ
function cellTime_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'HH:mm:ss');
  return String(v || '');
}

// ---------------------------------------------------------------- actions: GET

function getConfig_() {
  var sheet = configSheet_();
  var data = sheet.getDataRange().getValues();
  var config = { lat: '', lng: '', radius: 0 };
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0]);
    if (k === 'lat')    config.lat    = Number(data[i][1]) || '';
    if (k === 'lng')    config.lng    = Number(data[i][1]) || '';
    if (k === 'radius') config.radius = Number(data[i][1]) || 0;
  }
  return config;
}

function getKnownFaces_() {
  var sheet = facesSheet_();
  var data = sheet.getDataRange().getValues();
  var faces = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (!name) continue;
    try {
      var descriptor = JSON.parse(data[i][1]);
      if (descriptor && descriptor.length) faces.push({ label: name, descriptor: descriptor });
    } catch (err) {}
  }
  return faces;
}

function getTodayAttendance_() {
  var sheet = attendanceSheet_();
  var data = sheet.getDataRange().getValues();
  var today = today_();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (cellDate_(data[i][0]) !== today) continue;
    rows.push({
      'ชื่อ':            String(data[i][1] || ''),
      'เวลาเข้า':        cellTime_(data[i][2]),
      'เวลาออก':         cellTime_(data[i][3]) || '-',
      'Google Map Link': String(data[i][6] || ''),
      'หมายเหตุ':        String(data[i][7] || '')
    });
  }
  return rows;
}

function getTodaySiteCheckin_() {
  var sheet = siteSheet_();
  var data = sheet.getDataRange().getValues();
  var today = today_();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (cellDate_(data[i][0]) !== today) continue;
    rows.push({
      'เวลา':            cellTime_(data[i][1]),
      'ชื่อ':            String(data[i][2] || ''),
      'Google Map Link': String(data[i][5] || ''),
      'หมายเหตุ':        String(data[i][6] || '')
    });
  }
  return rows;
}

// ตัวช่วยไล่ปัญหา "เช็กอินแล้วไม่ขึ้นในเว็บ" — เปิด <URL>/exec?action=debugToday&key=<APP_KEY>
// จะโชว์ว่าระบบมองว่า "วันนี้" คืออะไร และอ่านวันที่ของ 5 แถวล่าสุดในชีตได้เป็นอะไร
function debugToday_() {
  var sheet = siteSheet_();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = Math.max(1, data.length - 5); i < data.length; i++) {
    rows.push({
      row: i + 1,
      rawDate: String(data[i][0]),
      parsedDate: cellDate_(data[i][0]),
      name: String(data[i][2] || '')
    });
  }
  return { today: today_(), lastRows: rows };
}

// ---------------------------------------------------------------- actions: POST

function registerUser_(body) {
  var name = String(body.name || '').trim();
  var descriptor = body.descriptor;
  if (!name) return { status: 'error', message: 'ไม่พบชื่อพนักงาน' };
  if (!descriptor || !descriptor.length) return { status: 'error', message: 'ไม่พบข้อมูลใบหน้า' };

  facesSheet_().appendRow([name, JSON.stringify(descriptor), Utilities.formatDate(new Date(), TZ, 'd/M/yyyy HH:mm:ss')]);
  return { status: 'success', message: 'บันทึกใบหน้าของ ' + name + ' สำเร็จ' };
}

function logAttendance_(body) {
  var name = String(body.name || '').trim();
  if (!name) return { error: 'bad_request', message: 'ไม่พบชื่อพนักงาน' };

  var lat  = body.lat  || '';
  var lng  = body.lng  || '';
  var note = String(body.note || '').trim();
  var link = mapLink_(lat, lng);

  // เช็กอินหน้างาน → เพิ่มแถวใหม่ทุกครั้ง (เช็กอินได้หลายรอบต่อวัน)
  if (body.sheetTarget === 'Site_CheckIn') {
    appendRowWithDate_(siteSheet_(), [new Date(), nowTime_(), name, lat, lng, link, note]);
    return { message: 'เช็กอินหน้างานสำเร็จ (' + name + ' เวลา ' + nowTime_().substring(0, 5) + ' น.)' };
  }

  // สแกนเข้า-ออกงาน → ครั้งแรกของวัน = เวลาเข้า, ครั้งถัดไป = เวลาออก
  var sheet = attendanceSheet_();
  var data = sheet.getDataRange().getValues();
  var today = today_();

  for (var i = data.length - 1; i >= 1; i--) {
    if (cellDate_(data[i][0]) === today && String(data[i][1]) === name) {
      var row = i + 1;
      sheet.getRange(row, 4).setValue(nowTime_()); // เวลาออก (อัปเดตทับได้ ถ้าสแกนออกซ้ำ)
      if (note) {
        var oldNote = String(data[i][7] || '');
        sheet.getRange(row, 8).setValue(oldNote ? oldNote + ' | ' + note : note);
      }
      return { message: 'บันทึกเวลาออกงานสำเร็จ (' + name + ' เวลา ' + nowTime_().substring(0, 5) + ' น.)' };
    }
  }

  appendRowWithDate_(sheet, [new Date(), name, nowTime_(), '', lat, lng, link, note]);
  return { message: 'บันทึกเวลาเข้างานสำเร็จ (' + name + ' เวลา ' + nowTime_().substring(0, 5) + ' น.)' };
}

function saveConfig_(body) {
  var sheet = configSheet_();
  var values = { lat: body.lat, lng: body.lng, radius: body.radius };
  var data = sheet.getDataRange().getValues();

  Object.keys(values).forEach(function (k) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === k) {
        sheet.getRange(i + 1, 2).setValue(values[k]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([k, values[k]]);
  });

  return { message: 'บันทึกการตั้งค่า GPS สำเร็จ' };
}

// ---------------------------------------------------------------- helpers

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
