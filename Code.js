// ฟังก์ชันสำหรับ Render หน้าเว็บ HTML เมื่อมีคนเปิดลิงก์ Web App
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('OSCE QuickScore & Executive Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// PIN สำหรับยืนยันสิทธิ์กรรมการ — เก็บใน Script Properties เท่านั้น ห้าม hardcode ในไฟล์นี้
// วิธีตั้งค่า (ทำครั้งเดียวต่อโปรเจกต์ ไม่ถูก commit ขึ้น git):
//   Apps Script Editor > ไอคอนเฟือง "Project Settings" > Script Properties > Add script property
//   Property: ADMIN_PIN   Value: <รหัสที่ต้องการ>
function getAdminPin_() {
  var pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!pin) {
    throw new Error('ยังไม่ได้ตั้งค่า ADMIN_PIN ใน Script Properties (Project Settings > Script Properties) กรุณาตั้งค่าก่อนใช้งาน');
  }
  return pin;
}

// ป้องกันการลองสุ่ม PIN (brute force) — ล็อกชั่วคราวหลังพิมพ์ผิดติดกันเกินจำนวนที่กำหนด
var PIN_MAX_ATTEMPTS = 5;
var PIN_LOCKOUT_SECONDS = 300; // 5 นาที
var PIN_ATTEMPT_CACHE_KEY = 'pin_fail_count';
var PIN_LOCKOUT_CACHE_KEY = 'pin_locked_until';

function requireAdminPin_(pin) {
  var cache = CacheService.getScriptCache();
  var now = Date.now();
  var lockedUntil = Number(cache.get(PIN_LOCKOUT_CACHE_KEY) || 0);

  if (lockedUntil && now < lockedUntil) {
    var waitSec = Math.ceil((lockedUntil - now) / 1000);
    throw new Error('กรอกรหัส PIN ผิดหลายครั้งเกินไป กรุณารออีก ' + waitSec + ' วินาทีแล้วลองใหม่');
  }

  if (pin !== getAdminPin_()) {
    var attempts = Number(cache.get(PIN_ATTEMPT_CACHE_KEY) || 0) + 1;
    if (attempts >= PIN_MAX_ATTEMPTS) {
      cache.put(PIN_LOCKOUT_CACHE_KEY, String(now + PIN_LOCKOUT_SECONDS * 1000), PIN_LOCKOUT_SECONDS);
      cache.remove(PIN_ATTEMPT_CACHE_KEY);
      throw new Error('กรอกรหัส PIN ผิดเกิน ' + PIN_MAX_ATTEMPTS + ' ครั้ง ถูกล็อกชั่วคราว ' + PIN_LOCKOUT_SECONDS + ' วินาที');
    }
    cache.put(PIN_ATTEMPT_CACHE_KEY, String(attempts), PIN_LOCKOUT_SECONDS);
    throw new Error('รหัส PIN ไม่ถูกต้อง (ผิดไปแล้ว ' + attempts + '/' + PIN_MAX_ATTEMPTS + ' ครั้ง)');
  }

  cache.remove(PIN_ATTEMPT_CACHE_KEY);
  cache.remove(PIN_LOCKOUT_CACHE_KEY);
}

// ตรวจสอบ PIN อย่างเดียวโดยไม่แก้ข้อมูล — ใช้ตอนกรรมการปลดล็อกโหมดบันทึกคะแนนบนหน้าเว็บ
function verifyAdminPin(pin) {
  requireAdminPin_(pin);
  return { ok: true };
}

function clampScore_(value) {
  var n = Number(value) || 0;
  return Math.max(0, Math.min(25, n));
}

var HEADER_ROW = [
  "เวลาบันทึก",
  "รหัสนักศึกษา",
  "ชื่อ-นามสกุล",
  "ฐาน1 (ซักประวัติ)",
  "ฐาน2 (ตรวจร่างกาย)",
  "ฐาน3 (หัตถการ)",
  "ฐาน4 (สื่อสาร)",
  "คะแนนรวม (100)",
  "ผลการสอบ",
  "หมายเหตุ"
];

// คืน sheet ปัจจุบัน พร้อมสร้าง Header อัตโนมัติหากยังไม่มี
function getSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
  }
  return sheet;
}

// อ่านข้อมูลคะแนนทั้งหมดจาก Sheet เพื่อแสดงในแดชบอร์ด (เรียกผ่าน google.script.run)
function getAllRecords() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  var records = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[1] && !row[2]) continue; // ข้ามแถวว่าง

    records.push({
      id: i + 2, // เลขแถวจริงใน sheet ใช้อ้างอิงสำหรับลบ/แก้ไข
      timestamp: row[0] instanceof Date ? row[0].toLocaleString('th-TH') : row[0],
      studentId: row[1],
      studentName: row[2],
      s1: Number(row[3]) || 0,
      s2: Number(row[4]) || 0,
      s3: Number(row[5]) || 0,
      s4: Number(row[6]) || 0,
      total: Number(row[7]) || 0,
      passed: row[8] === "PASS",
      note: row[9] || "",
      synced: true
    });
  }

  return records.reverse(); // แสดงรายการล่าสุดขึ้นก่อน
}

// บันทึกคะแนนสอบ 1 รายการลง Sheet (เรียกผ่าน google.script.run)
function addRecord(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('ระบบมีผู้บันทึกข้อมูลพร้อมกันจำนวนมาก กรุณาลองบันทึกใหม่อีกครั้ง');
  }

  try {
    requireAdminPin_(data && data.pin);

    if (!data || !data.studentId || !data.studentName) {
      throw new Error('กรุณากรอกรหัสนักศึกษาและชื่อ-นามสกุลให้ครบถ้วน');
    }

    var s1 = clampScore_(data.s1);
    var s2 = clampScore_(data.s2);
    var s3 = clampScore_(data.s3);
    var s4 = clampScore_(data.s4);
    var total = s1 + s2 + s3 + s4;
    var passed = total >= 60;
    var now = new Date();

    var sheet = getSheet_();
    sheet.appendRow([
      now,
      data.studentId,
      data.studentName,
      s1, s2, s3, s4,
      total,
      passed ? "PASS" : "FAIL",
      data.note || ""
    ]);

    return {
      id: sheet.getLastRow(),
      timestamp: now.toLocaleString('th-TH'),
      studentId: data.studentId,
      studentName: data.studentName,
      s1: s1, s2: s2, s3: s3, s4: s4,
      total: total,
      passed: passed,
      note: data.note || "",
      synced: true
    };
  } finally {
    lock.releaseLock();
  }
}

// ลบข้อมูล 1 แถวออกจาก Sheet ตามเลขแถว (id ที่ได้จาก getAllRecords)
function deleteRecordByRow(rowId, pin) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('ระบบมีผู้แก้ไขข้อมูลพร้อมกันจำนวนมาก กรุณาลองใหม่อีกครั้ง');
  }

  try {
    requireAdminPin_(pin);

    var sheet = getSheet_();
    var row = Number(rowId);
    if (!row || row < 2 || row > sheet.getLastRow()) {
      throw new Error('ไม่พบรายการที่ต้องการลบ (ข้อมูลอาจถูกแก้ไขไปแล้ว กรุณารีเฟรช)');
    }
    sheet.deleteRow(row);
    return { status: 'success' };
  } finally {
    lock.releaseLock();
  }
}

// รับข้อมูลผ่าน HTTP POST (สำหรับเรียกจากภายนอก Apps Script เช่น curl/Postman)
// การบันทึกจากตัวหน้าเว็บเองใช้ addRecord() ผ่าน google.script.run เป็นหลัก เพราะรู้ผลสำเร็จ/ล้มเหลวจริง
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var record = addRecord(data);
    return ContentService.createTextOutput(JSON.stringify({ status: "success", record: record }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
