/**
 * ระบบลงเวลาเข้า-ออกงานผ่าน Google Apps Script Web App
 * เชื่อมโยงฐานข้อมูล Google Sheets: Users, Branches, Logs, Leaves
 * รองรับ:
 * 1. สแกนเข้ามากกว่า 1 ครั้งต่อวัน คิดเวลาแรกสุดเท่านั้น (ไม่บันทึกซ้ำ)
 * 2. สแกนออกมากกว่า 1 ครั้งต่อวัน คิดเวลาสุดท้ายเท่านั้น (อัปเดตแถวเดิม)
 * 3. เวลาเกินจากปกติคิดเป็นโอที บันทึกเวลาโอทีลงตาราง Logs
 * 4. สแกนออกก่อนเวลาคิดเป็นออกก่อน คำนวณเวลาออกก่อนลงตาราง Logs
 */

// 1. ฟังก์ชันโหลดหน้าเว็บ Web App เมื่อเปิดผ่าน Apps Script URL โดยตรง
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var userEmail = Session.getActiveUser().getEmail();
  template.userEmail = userEmail;
  
  var htmlOutput = template.evaluate()
    .setTitle('ระบบลงเวลาเข้า-ออกงานและจัดการการลา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
    
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return htmlOutput;
}

// 2. ฟังก์ชัน API สำหรับรับคำขอจากหน้าเว็บภายนอก (GitHub Pages / Vercel / Local Server)
function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : "{}";
    var req = JSON.parse(contents);
    var action = req.action;
    var result = { success: false, message: "Action not recognized" };

    if (action === "getInitialData") {
      result = getInitialData(req.email || "");
      result.success = true;
    } else if (action === "recordAttendance") {
      result = recordAttendance(req);
    } else if (action === "getRecentLogs") {
      var logs = getRecentLogs(req.userId || "");
      result = { success: true, logs: logs };
    } else if (action === "registerUser") {
      result = registerUser(req);
    } else if (action === "submitLeave") {
      result = submitLeaveRequest(req);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// รวมไฟล์ย่อย
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 3. ดึงข้อมูลพนักงานและสาขาตาม Email
function getInitialData(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users");
  var branchSheet = ss.getSheetByName("Branches");
  
  var users = userSheet.getDataRange().getValues();
  var branches = branchSheet.getDataRange().getValues();
  
  var currentUser = null;
  for (var i = 1; i < users.length; i++) {
    if (users[i][4] && users[i][4].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
      currentUser = {
        userId: users[i][0],
        name: users[i][1],
        role: users[i][2],
        branchId: users[i][3],
        email: users[i][4],
        status: users[i][5]
      };
      break;
    }
  }
  
  var branchList = [];
  for (var j = 1; j < branches.length; j++) {
    if (branches[j][0]) {
      branchList.push({
        branchId: branches[j][0],
        branchName: branches[j][1],
        lat: Number(branches[j][2]),
        lng: Number(branches[j][3]),
        radius: Number(branches[j][4]),
        workStart: branches[j][5] ? Utilities.formatDate(new Date(branches[j][5]), Session.getScriptTimeZone(), "HH:mm") : "08:30",
        workEnd: branches[j][6] ? Utilities.formatDate(new Date(branches[j][6]), Session.getScriptTimeZone(), "HH:mm") : "17:30"
      });
    }
  }
  
  return {
    user: currentUser,
    branches: branchList
  };
}

// 4. ลงทะเบียนพนักงานใหม่
function registerUser(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var newId = "U" + ("000" + sheet.getLastRow()).slice(-3);
    
    sheet.appendRow([
      newId,
      data.name,
      "user",
      data.branchId,
      data.email,
      "ACTIVE"
    ]);
    
    return { success: true, message: "ลงทะเบียนสำเร็จ กรุณารอผู้ดูแลระบบตรวจสอบสิทธิ์" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 5. คำนวณระยะห่างพิกัด GPS
function calculateDistance(lat1, lon1, lat2, lon2) {
  var R = 6371e3;
  var phi1 = lat1 * Math.PI / 180;
  var phi2 = lat2 * Math.PI / 180;
  var dPhi = (lat2 - lat1) * Math.PI / 180;
  var dLambda = (lon2 - lon1) * Math.PI / 180;

  var a = Math.sin(dPhi / 2) * Math.sin(dPhi / 2) +
          Math.cos(phi1) * Math.cos(phi2) *
          Math.sin(dLambda / 2) * Math.sin(dLambda / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c);
}

// 6. บันทึกรูปถ่ายลง Google Drive
function saveImageToDrive(base64Data, userId) {
  try {
    var folderName = "Attendance_Photos";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var split = base64Data.split('base64,');
    var contentType = split[0].split(':')[1].split(';')[0];
    var bytes = Utilities.base64Decode(split[1]);
    var blob = Utilities.newBlob(bytes, contentType, "Scan_" + userId + "_" + new Date().getTime() + ".jpg");
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Drive_Error: " + e.message;
  }
}

// ตรวจสอบและสร้างหัวตาราง Logs ให้รองรับคอลัมน์ OT และ ออกก่อน
function ensureLogHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Log ID", "User ID", "Timestamp", "Type", "Branch ID", 
      "Distance (m)", "Photo URL", "Status", "Note", "OT (นาที)", "ออกก่อน (นาที)"
    ]);
  } else {
    var lastCol = Math.max(sheet.getLastColumn(), 11);
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headerRow[9] || headerRow[9].toString().trim() === "") {
      sheet.getRange(1, 10).setValue("OT (นาที)");
    }
    if (!headerRow[10] || headerRow[10].toString().trim() === "") {
      sheet.getRange(1, 11).setValue("ออกก่อน (นาที)");
    }
  }
}

// ฟังก์ชันแปลงจำนวนนาทีเป็นข้อความ ชั่วโมง/นาที ที่อ่านง่าย
function formatMinutesToText(minutes) {
  minutes = Math.round(minutes);
  if (minutes < 60) {
    return minutes + " นาที";
  }
  var h = Math.floor(minutes / 60);
  var m = minutes % 60;
  if (m === 0) {
    return h + " ชม.";
  }
  return h + " ชม. " + m + " นาที";
}

// 7. ฟังก์ชันบันทึกเวลาสแกนเข้า-ออกงาน (พร้อมกฎ 4 ข้อ)
function recordAttendance(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Logs");
    var branchSheet = ss.getSheetByName("Branches");
    
    // ตรวจสอบข้อมูลสาขา
    var branches = branchSheet.getDataRange().getValues();
    var targetBranch = null;
    
    for (var i = 1; i < branches.length; i++) {
      if (branches[i][0] === payload.branchId) {
        targetBranch = {
          lat: Number(branches[i][2]),
          lng: Number(branches[i][3]),
          radius: Number(branches[i][4]),
          workStart: branches[i][5],
          workEnd: branches[i][6]
        };
        break;
      }
    }
    
    if (!targetBranch) {
      return { success: false, message: "ไม่พบข้อมูลสาขาที่กำหนด" };
    }
    
    // ตรวจสอบระยะห่างด้วย Haversine
    var distance = calculateDistance(payload.lat, payload.lng, targetBranch.lat, targetBranch.lng);
    if (distance > targetBranch.radius) {
      return { 
        success: false, 
        message: "อยู่นอกพื้นที่ที่กำหนด! ระยะห่างของคุณคือ " + distance + " เมตร (อนุญาตไม่เกิน " + targetBranch.radius + " เมตร)" 
      };
    }

    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var todayDateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var currentTimeStr = Utilities.formatDate(now, tz, "HH:mm");
    var timestampStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");

    // ตรวจสอบหัวตาราง Logs
    ensureLogHeaders(logSheet);

    var logData = logSheet.getDataRange().getValues();

    // ตรวจหาประวัติการสแกนของพนักงานในวันนี้
    var existingCheckInRow = -1;
    var existingCheckInTime = "";
    var existingCheckOutRow = -1;
    var existingCheckOutTime = "";

    for (var r = 1; r < logData.length; r++) {
      var rowUserId = logData[r][1];
      var rowTimestamp = logData[r][2];
      var rowType = logData[r][3];

      if (rowUserId === payload.userId && rowTimestamp) {
        var rowDateStr = "";
        if (rowTimestamp instanceof Date) {
          rowDateStr = Utilities.formatDate(rowTimestamp, tz, "yyyy-MM-dd");
        } else {
          rowDateStr = rowTimestamp.toString().substring(0, 10);
        }

        if (rowDateStr === todayDateStr) {
          if (rowType === "CHECK_IN") {
            existingCheckInRow = r + 1; // 1-based index
            existingCheckInTime = (rowTimestamp instanceof Date) 
              ? Utilities.formatDate(rowTimestamp, tz, "HH:mm") 
              : rowTimestamp.toString().substring(11, 16);
          } else if (rowType === "CHECK_OUT") {
            existingCheckOutRow = r + 1;
            existingCheckOutTime = (rowTimestamp instanceof Date) 
              ? Utilities.formatDate(rowTimestamp, tz, "HH:mm") 
              : rowTimestamp.toString().substring(11, 16);
          }
        }
      }
    }

    // ==========================================
    // กฎข้อ 1: สแกนเข้ามากกว่า 1 ครั้งต่อวัน ให้คิดเวลาแรกสุดเท่านั้น
    // (หากมีเวลาเข้าอยู่แล้วในวันนี้ จะไม่นับและไม่บันทึกซ้ำ)
    // ==========================================
    if (payload.type === "CHECK_IN") {
      if (existingCheckInRow > 0) {
        return {
          success: false,
          message: "วันนี้คุณได้ลงเวลาเข้างานไปแล้วเมื่อเวลา " + existingCheckInTime + " น. (ระบบบันทึกเวลาแรกสุดเท่านั้น ไม่บันทึกซ้ำ)"
        };
      }

      // คำนวณเวลาเข้าสาย
      var workStartTimeStr = "08:30";
      if (targetBranch.workStart) {
        workStartTimeStr = (targetBranch.workStart instanceof Date)
          ? Utilities.formatDate(targetBranch.workStart, tz, "HH:mm")
          : targetBranch.workStart.toString();
      }

      var currentMin = now.getHours() * 60 + now.getMinutes();
      var startParts = workStartTimeStr.split(":");
      var startMin = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);

      var status = "SUCCESS";
      var note = "เข้างานตรงเวลา";

      if (currentMin > startMin) {
        var lateMinutes = currentMin - startMin;
        var lateText = formatMinutesToText(lateMinutes);
        status = "LATE";
        note = "เข้างานสาย " + lateText + " (เริ่มงาน " + workStartTimeStr + " สแกนเข้า " + currentTimeStr + ")";
      }

      var photoUrl = saveImageToDrive(payload.imageBase64, payload.userId);
      var logId = "LOG" + Utilities.formatDate(now, tz, "yyyyMMddHHmmss");

      logSheet.appendRow([
        logId,
        payload.userId,
        timestampStr,
        "CHECK_IN",
        payload.branchId,
        distance,
        photoUrl,
        status,
        note,
        0,  // OT
        0   // ออกก่อน
      ]);

      var checkInMsg = "ลงเวลาเข้างานเรียบร้อยแล้ว (" + currentTimeStr + " น.)";
      if (status === "LATE") {
        checkInMsg += " (สาย " + formatMinutesToText(currentMin - startMin) + ")";
      }

      return {
        success: true,
        message: checkInMsg,
        status: status
      };
    }

    // ==========================================
    // กฎข้อ 2, 3, 4: สแกนออก
    // 2. สแกนออกมากกว่า 1 ครั้งต่อวัน คิดเวลาสุดท้ายเท่านั้น (อัปเดตแถวเดิม)
    // 3. เวลาเกินจากปกติคิดเป็นโอที บันทึกเวลาโอทีลงตารางด้วย
    // 4. สแกนออกก่อนเวลาคิดเป็นออกก่อน คำนวณเวลาออกก่อนด้วย
    // ==========================================
    if (payload.type === "CHECK_OUT") {
      var workEndTimeStr = "17:30";
      if (targetBranch.workEnd) {
        workEndTimeStr = (targetBranch.workEnd instanceof Date)
          ? Utilities.formatDate(targetBranch.workEnd, tz, "HH:mm")
          : targetBranch.workEnd.toString();
      }

      var currentMin = now.getHours() * 60 + now.getMinutes();
      var endParts = workEndTimeStr.split(":");
      var endMin = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);

      var status = "SUCCESS";
      var note = "ออกงานตามเวลาปกติ";
      var otMinutes = 0;
      var earlyMinutes = 0;

      if (currentMin > endMin) {
        // กฎข้อ 3: เวลาเกินจากปกติ คิดเป็นโอที
        otMinutes = currentMin - endMin;
        var otText = formatMinutesToText(otMinutes);
        status = "OVERTIME";
        note = "ทำโอที " + otText + " (เลิกงาน " + workEndTimeStr + " สแกนออก " + currentTimeStr + ")";
      } else if (currentMin < endMin) {
        // กฎข้อ 4: สแกนออกก่อนเวลา คิดเป็นออกก่อน
        earlyMinutes = endMin - currentMin;
        var earlyText = formatMinutesToText(earlyMinutes);
        status = "EARLY_LEAVE";
        note = "ออกก่อนเวลา " + earlyText + " (เลิกงาน " + workEndTimeStr + " สแกนออก " + currentTimeStr + ")";
      }

      var photoUrl = saveImageToDrive(payload.imageBase64, payload.userId);

      // กฎข้อ 2: ถ้ามีสแกนออกในวันนี้แล้ว ให้อัปเดตแถวเดิมเป็นเวลาล่าสุดเท่านั้น
      if (existingCheckOutRow > 0) {
        logSheet.getRange(existingCheckOutRow, 3).setValue(timestampStr); // Timestamp ล่าสุด
        logSheet.getRange(existingCheckOutRow, 6).setValue(distance);     // Distance
        logSheet.getRange(existingCheckOutRow, 7).setValue(photoUrl);     // PhotoUrl ล่าสุด
        logSheet.getRange(existingCheckOutRow, 8).setValue(status);       // Status
        logSheet.getRange(existingCheckOutRow, 9).setValue(note);         // Note
        logSheet.getRange(existingCheckOutRow, 10).setValue(otMinutes);   // OT (นาที)
        logSheet.getRange(existingCheckOutRow, 11).setValue(earlyMinutes);// ออกก่อน (นาที)

        var updateMsg = "อัปเดตเวลาออกงานเป็นเวลาล่าสุด (" + currentTimeStr + " น.) เรียบร้อยแล้ว";
        if (otMinutes > 0) updateMsg += " (โอที " + formatMinutesToText(otMinutes) + ")";
        if (earlyMinutes > 0) updateMsg += " (ออกก่อน " + formatMinutesToText(earlyMinutes) + ")";

        return {
          success: true,
          message: updateMsg,
          status: status
        };
      } else {
        // สแกนออกครั้งแรกของวันนี้ -> บันทึกแถวใหม่
        var logId = "LOG" + Utilities.formatDate(now, tz, "yyyyMMddHHmmss");
        logSheet.appendRow([
          logId,
          payload.userId,
          timestampStr,
          "CHECK_OUT",
          payload.branchId,
          distance,
          photoUrl,
          status,
          note,
          otMinutes,
          earlyMinutes
        ]);

        var checkOutMsg = "ลงเวลาออกงานเรียบร้อยแล้ว (" + currentTimeStr + " น.)";
        if (otMinutes > 0) checkOutMsg += " (โอที " + formatMinutesToText(otMinutes) + ")";
        if (earlyMinutes > 0) checkOutMsg += " (ออกก่อน " + formatMinutesToText(earlyMinutes) + ")";

        return {
          success: true,
          message: checkOutMsg,
          status: status
        };
      }
    }

    return { success: false, message: "ประเภทการลงเวลาไม่ถูกต้อง" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.toString() };
  }
}

// 8. ฟังก์ชันยื่นคำขอลางาน
function submitLeaveRequest(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leaveSheet = ss.getSheetByName("Leaves");
    var leaveId = "L" + ("000" + leaveSheet.getLastRow()).slice(-3);
    
    leaveSheet.appendRow([
      leaveId,
      data.userId,
      data.leaveType,
      data.startDate,
      data.endDate,
      data.reason,
      "PENDING"
    ]);
    
    return { success: true, message: "ส่งคำขอลาเรียบร้อยแล้ว รอผู้ดูแลระบบอนุมัติ" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// 9. ดึงประวัติการลงเวลาล่าสุด
function getRecentLogs(userId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Logs");
  var data = logSheet.getDataRange().getValues();
  var logs = [];
  var tz = Session.getScriptTimeZone();
  
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === userId) {
      var ts = data[i][2];
      var tsStr = (ts instanceof Date) 
        ? Utilities.formatDate(ts, tz, "yyyy-MM-dd HH:mm:ss") 
        : ts.toString();

      var otVal = Number(data[i][9]) || 0;
      var earlyVal = Number(data[i][10]) || 0;

      logs.push({
        timestamp: tsStr,
        type: data[i][3],
        distance: data[i][5],
        status: data[i][7],
        note: data[i][8] || "",
        ot: otVal > 0 ? formatMinutesToText(otVal) : "",
        early: earlyVal > 0 ? formatMinutesToText(earlyVal) : ""
      });
      if (logs.length >= 5) break;
    }
  }
  return logs;
}
