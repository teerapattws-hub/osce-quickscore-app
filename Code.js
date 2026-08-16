// ฟังก์ชันสำหรับ Render หน้าเว็บ HTML เมื่อมีคนเปิดลิงก์ Web App
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('OSCE QuickScore & Executive Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันสำหรับรับข้อมูลการบันทึกคะแนนจากหน้าเว็บลง Google Sheet
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // ป้องกันปัญหากรรมการบันทึกพร้อมกันในวินาทีเดียวกัน
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    // สร้าง Header อัตโนมัติหากเป็นบรรทัดแรก
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
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
      ]);
      // จัดรูปแบบ Header ให้สวยงาม
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    }
    
    // บันทึกข้อมูลแถวใหม่
    sheet.appendRow([
      new Date(),
      data.studentId,
      data.studentName,
      Number(data.s1),
      Number(data.s2),
      Number(data.s3),
      Number(data.s4),
      Number(data.total),
      data.passed ? "PASS" : "FAIL",
      data.note || ""
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "บันทึกเรียบร้อย" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}