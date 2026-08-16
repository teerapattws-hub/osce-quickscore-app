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

function clampScore_(value) {
  var n = Number(value) || 0;
  return Math.max(0, Math.min(25, n));
}

// ป้องกัน Formula Injection ในตัว Sheet เอง — ถ้าปล่อยให้ค่าที่ขึ้นต้นด้วย =, +, -, @ เข้าไปตรง ๆ
// Google Sheets จะตีความเป็นสูตรทันทีเหมือนมีคนพิมพ์ใส่เซลล์เอง (เช่น ดึงข้อมูลออกไปเซิร์ฟเวอร์ภายนอกผ่าน IMPORTXML)
// เติม ' นำหน้าเพื่อบังคับให้เก็บเป็นข้อความเสมอ (Sheets จะตัด ' ออกเองตอนอ่านค่ากลับ ไม่กระทบข้อมูลที่แสดง)
function sanitizeCellText_(value) {
  var str = value == null ? '' : String(value);
  if (/^[=+\-@\t\r]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

var HEADER_ROW = [
  "เวลาบันทึก",
  "ชื่อ-นามสกุล",
  "ฐาน1 (หัตถการ)",
  "ฐาน2 (ข้อห้ามและข้อควรระวัง)",
  "ฐาน3 (การให้คำแนะนำ)",
  "ฐาน4 (การช่วยฟื้นคืนชีพ)",
  "คะแนนรวม (100)",
  "ผลการสอบ",
  "หมายเหตุ"
];

// คืน sheet ปัจจุบัน พร้อมสร้าง Header อัตโนมัติหากยังไม่มี
// ระบุ Spreadsheet ตรง ๆ ด้วย ID แทนการพึ่ง SpreadsheetApp.getActiveSpreadsheet()
// เพราะเมื่อรันเป็น Web App (ไม่ได้เปิดจากหน้า Sheet) จะไม่มี "active spreadsheet" ให้อ้างอิง
// ID นี้เอามาจาก URL ของชีท: https://docs.google.com/spreadsheets/d/<ID>/edit
var SPREADSHEET_ID = '1iDBcZBrZCw3eUk4Lg35TUdo780zzWhVbhK5cPqKJBqI';

function getSheet_() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
  }
  return sheet;
}

// อ่านข้อมูลคะแนนทั้งหมดจาก Sheet (รวมชื่อนักศึกษา) — เป็น private helper เท่านั้น
// ห้าม expose ตรง ๆ ผ่าน google.script.run เพราะมีข้อมูลระบุตัวตน ต้องผ่าน getFullRecords() ที่เช็ค PIN ก่อนเสมอ
// โครงสร้างคอลัมน์: A=เวลาบันทึก, B=ชื่อ-นามสกุล, C-F=ฐาน1-4, G=คะแนนรวม, H=ผลการสอบ, I=หมายเหตุ
function getAllRecords_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  var records = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[1]) continue; // ข้ามแถวว่าง (ไม่มีชื่อ)

    records.push({
      id: i + 2, // เลขแถวจริงใน sheet ใช้อ้างอิงสำหรับลบ/แก้ไข
      timestamp: row[0] instanceof Date ? row[0].toLocaleString('th-TH') : row[0],
      studentName: String(row[1]),
      s1: Number(row[2]) || 0,
      s2: Number(row[3]) || 0,
      s3: Number(row[4]) || 0,
      s4: Number(row[5]) || 0,
      total: Number(row[6]) || 0,
      passed: row[7] === "PASS",
      note: row[8] || "",
      synced: true
    });
  }

  return records.reverse(); // แสดงรายการล่าสุดขึ้นก่อน
}

// ข้อมูลสรุปแบบไม่ระบุตัวตน (ไม่มีชื่อ/รหัสนักศึกษา/หมายเหตุ) — ใช้กับแท็บภาพรวมที่เปิดดูได้ทุกคน ไม่ต้องมี PIN
function getSummaryRecords() {
  return getAllRecords_().map(function(r) {
    return {
      id: r.id,
      timestamp: r.timestamp,
      s1: r.s1, s2: r.s2, s3: r.s3, s4: r.s4,
      total: r.total,
      passed: r.passed
    };
  });
}

// ข้อมูลเต็มพร้อมชื่อนักศึกษา — ต้องยืนยัน PIN ก่อนทุกครั้ง ใช้กับแท็บ "รายชื่อนักศึกษา & ผลสอบ"
function getFullRecords(pin) {
  requireAdminPin_(pin);
  return getAllRecords_();
}

// Top N คะแนนสูงสุดพร้อมชื่อ — ตั้งใจเปิดให้ทุกคนดูได้โดยไม่ต้องมี PIN (ตามที่ผู้ใช้เลือกให้ Top 3 แสดงชื่อสาธารณะ)
// ต่างจาก getSummaryRecords เพราะเปิดเผยเฉพาะกลุ่มบนสุดเท่านั้น ไม่ใช่รายชื่อทั้งหมด
function getTopScorers(n) {
  var count = Number(n) || 3;
  return getAllRecords_()
    .slice()
    .sort(function(a, b) { return b.total - a.total; })
    .slice(0, count)
    .map(function(r) {
      return { studentName: r.studentName, total: r.total, passed: r.passed };
    });
}

// บันทึกคะแนนสอบ 1 รายการลง Sheet (เรียกผ่าน google.script.run)
function addRecord(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('ระบบมีผู้บันทึกข้อมูลพร้อมกันจำนวนมาก กรุณาลองบันทึกใหม่อีกครั้ง');
  }

  try {
    requireAdminPin_(data && data.pin);

    if (!data || !data.studentName) {
      throw new Error('กรุณากรอกชื่อ-นามสกุลนักศึกษา');
    }

    var s1 = clampScore_(data.s1);
    var s2 = clampScore_(data.s2);
    var s3 = clampScore_(data.s3);
    var s4 = clampScore_(data.s4);
    var total = s1 + s2 + s3 + s4;
    var passed = total >= 60;
    var now = new Date();

    var studentName = sanitizeCellText_(data.studentName);
    var note = sanitizeCellText_(data.note || "");

    var sheet = getSheet_();
    sheet.appendRow([
      now,
      studentName,
      s1, s2, s3, s4,
      total,
      passed ? "PASS" : "FAIL",
      note
    ]);

    return {
      id: sheet.getLastRow(),
      timestamp: now.toLocaleString('th-TH'),
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
