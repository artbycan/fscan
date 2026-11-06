/**
 * ระบบลงเวลาเข้า-ออกงานผ่าน Google Apps Script Web App
 * เชื่อมโยงฐานข้อมูล Google Sheets: Users, Branches, Logs, Leaves
 * รองรับ:
 * 1. สแกนเข้ามากกว่า 1 ครั้งต่อวัน คิดเวลาแรกสุดเท่านั้น (ไม่บันทึกซ้ำ)
 * 2. สแกนออกมากกว่า 1 ครั้งต่อวัน คิดเวลาสุดท้ายเท่านั้น (อัปเดตแถวเดิม)
 * 3. เวลาเกินจากปกติคิดเป็นโอที บันทึกเวลาโอทีลงตาราง Logs
 * 4. สแกนออกก่อนเวลาคิดเป็นออกก่อน คำนวณเวลาออกก่อนลงตาราง Logs
 * 5. คอลัมน์ เข้าสาย (นาที) ในตาราง Logs
 * 6. คอลัมน์ ค่าโอทีต่อชั่วโมง และ โอทีต่อนาที (บาท) ในตาราง Users
 * 7. แดชบอร์ดสรุปรายเดือน (ขาด ลา มา สาย เวลาโอที เงินค่าโอที) สำหรับแอดมิน + ส่งออกไฟล์ Excel
 * 8. พนักงานสมัครใหม่มีสถานะ INACTIVE จนกว่าแอดมินจะอนุมัติ พร้อมหน้าอนุมัติและจัดการสิทธิ์ในแอป
 * 9. หน้าจัดการสาขา (เพิ่ม/แก้ไข/ลบสาขา, ตำแหน่ง GPS, รัศมี, เวลาเข้างาน-ออกงาน) สำหรับแอดมิน
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
    } else if (action === "getAdminDashboard") {
      result = getAdminMonthlySummary(req.monthStr || "", req.adminEmail || "");
    } else if (action === "updateUserOtRate") {
      result = updateUserOtRate(req);
    } else if (action === "getAllUsersForAdmin") {
      result = getAllUsersForAdmin(req.adminEmail || "");
    } else if (action === "updateUserPermission") {
      result = updateUserPermission(req);
    } else if (action === "getBranchesForAdmin") {
      result = getBranchesForAdmin(req.adminEmail || "");
    } else if (action === "saveBranch") {
      result = saveBranch(req);
    } else if (action === "deleteBranch") {
      result = deleteBranch(req);
    } else if (action === "getUserDailySummary") {
      result = getUserDailySummary(req.userId || "", req.monthStr || "");
    } else if (action === "getAdminDailySummary") {
      result = getAdminDailySummary(req.dateStr || "", req.adminEmail || "", req.branchId || "");
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

// ตรวจสอบและสร้างหัวตาราง Users ให้รองรับคอลัมน์ ค่าโอทีต่อชั่วโมง และ โอทีต่อนาที (บาท)
function ensureUserHeaders(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "User ID", "Name", "Role", "Branch ID", "Email", "Status",
      "ค่าโอทีต่อชั่วโมง", "โอทีต่อนาที (บาท)"
    ]);
  } else {
    var lastCol = Math.max(sheet.getLastColumn(), 8);
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headerRow[6] || headerRow[6].toString().trim() === "") {
      sheet.getRange(1, 7).setValue("ค่าโอทีต่อชั่วโมง");
    }
    if (!headerRow[7] || headerRow[7].toString().trim() === "") {
      sheet.getRange(1, 8).setValue("โอทีต่อนาที (บาท)");
    }

    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rateRange = sheet.getRange(2, 7, lastRow - 1, 2);
      var rateVals = rateRange.getValues();
      var rateFormulas = rateRange.getFormulas();
      for (var i = 0; i < rateVals.length; i++) {
        var rowNum = i + 2;
        var hrRate = Number(rateVals[i][0]) || 0;
        var minCellVal = rateVals[i][1];
        var minFormula = rateFormulas[i][1];
        if (!minFormula && (minCellVal === "" || minCellVal === null || (hrRate > 0 && Number(minCellVal) === 0))) {
          sheet.getRange(rowNum, 8).setFormula("=IF(ISNUMBER(G" + rowNum + "), G" + rowNum + "/60, 0)");
        }
      }
    }
  }
}

// ตรวจสอบและสร้างหัวตาราง Branches
function ensureBranchHeaders(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "branch_id", "branch_name", "lat", "lng", "radius_meters", "work_start", "work_end"
    ]);
  }
}

// แปลงค่าเวลาในชีต Branches (รองรับทั้ง Date object และ String "HH:mm")
function formatSheetTime(val, defaultTime) {
  if (!val && val !== 0) return defaultTime || "08:30";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  var str = val.toString().trim();
  var match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    return ("0" + parseInt(match[1], 10)).slice(-2) + ":" + match[2];
  }
  return defaultTime || "08:30";
}

// ตรวจสอบว่าอีเมลที่เรียกใช้งานเป็น Admin ที่ ACTIVE หรือไม่
function verifyAdminAccess(usersData, adminEmail) {
  var callerEmail = (adminEmail || Session.getActiveUser().getEmail() || "").toString().trim().toLowerCase();
  if (!callerEmail) return false;
  for (var i = 1; i < usersData.length; i++) {
    var uEmail = (usersData[i][4] || "").toString().trim().toLowerCase();
    var uRole = (usersData[i][2] || "").toString().trim().toLowerCase();
    var uStatus = (usersData[i][5] || "ACTIVE").toString().trim().toUpperCase();
    if (uEmail === callerEmail && uRole === "admin" && uStatus === "ACTIVE") {
      return true;
    }
  }
  return false;
}

// 3. ดึงข้อมูลพนักงานและสาขาตาม Email
function getInitialData(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users");
  var branchSheet = ss.getSheetByName("Branches");
  var logSheet = ss.getSheetByName("Logs");
  
  ensureUserHeaders(userSheet);
  ensureBranchHeaders(branchSheet);
  ensureLogHeaders(logSheet);

  var users = userSheet.getDataRange().getValues();
  var branches = branchSheet.getDataRange().getValues();
  
  var currentUser = null;
  for (var i = 1; i < users.length; i++) {
    if (users[i][4] && users[i][4].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
      var otRateHour = Number(users[i][6]) || 0;
      var otRateMin = Number(users[i][7]) || (otRateHour / 60);
      currentUser = {
        userId: users[i][0],
        name: users[i][1],
        role: (users[i][2] || "user").toString().trim().toLowerCase(),
        branchId: users[i][3],
        email: users[i][4],
        status: (users[i][5] || "INACTIVE").toString().trim().toUpperCase(),
        otRatePerHour: otRateHour,
        otRatePerMinute: Number(otRateMin.toFixed(4))
      };
      break;
    }
  }
  
  var branchList = [];
  for (var j = 1; j < branches.length; j++) {
    if (branches[j][0]) {
      branchList.push({
        branchId: branches[j][0].toString().trim(),
        branchName: branches[j][1] ? branches[j][1].toString().trim() : "",
        lat: Number(branches[j][2]) || 0,
        lng: Number(branches[j][3]) || 0,
        radius: Number(branches[j][4]) || 200,
        workStart: formatSheetTime(branches[j][5], "08:30"),
        workEnd: formatSheetTime(branches[j][6], "17:30")
      });
    }
  }
  
  return {
    user: currentUser,
    branches: branchList
  };
}

// 4. ลงทะเบียนพนักงานใหม่ (กำหนด Status เริ่มต้นเป็น INACTIVE รอแอดมินอนุมัติ)
function registerUser(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    ensureUserHeaders(sheet);

    var users = sheet.getDataRange().getValues();
    var cleanEmail = (data.email || "").toString().trim().toLowerCase();
    for (var i = 1; i < users.length; i++) {
      if ((users[i][4] || "").toString().trim().toLowerCase() === cleanEmail) {
        return { success: false, message: "อีเมลนี้ถูกลงทะเบียนในระบบแล้ว" };
      }
    }

    var nextRow = sheet.getLastRow() + 1;
    var maxNum = 0;
    for (var k = 1; k < users.length; k++) {
      var uidStr = (users[k][0] || "").toString();
      var m = uidStr.match(/\d+/);
      if (m) {
        var n = parseInt(m[0], 10);
        if (n > maxNum) maxNum = n;
      }
    }
    var newId = "U" + ("000" + (maxNum + 1)).slice(-3);
    var defaultOtHour = Number(data.otRatePerHour) || 0;
    
    sheet.appendRow([
      newId,
      data.name,
      "user",
      data.branchId,
      data.email,
      "INACTIVE", // สมัครใหม่เป็น INACTIVE จนกว่าแอดมินจะอนุมัติ
      defaultOtHour,
      defaultOtHour / 60
    ]);
    sheet.getRange(nextRow, 8).setFormula("=IF(ISNUMBER(G" + nextRow + "), G" + nextRow + "/60, 0)");
    
    return { success: true, message: "ลงทะเบียนสำเร็จ บัญชีของคุณมีสถานะ INACTIVE กรุณารอผู้ดูแลระบบ (Admin) อนุมัติสิทธิ์เข้าใช้งาน" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ดึงข้อมูลสาขาทั้งหมดสำหรับหน้าจัดการสาขา (เฉพาะ Admin)
function getBranchesForAdmin(adminEmail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    var branchSheet = ss.getSheetByName("Branches");
    ensureUserHeaders(userSheet);
    ensureBranchHeaders(branchSheet);

    var users = userSheet.getDataRange().getValues();
    if (!verifyAdminAccess(users, adminEmail)) {
      return { success: false, message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้น" };
    }

    var branches = branchSheet.getDataRange().getValues();
    var branchList = [];
    for (var j = 1; j < branches.length; j++) {
      if (branches[j][0]) {
        branchList.push({
          branchId: branches[j][0].toString().trim(),
          branchName: branches[j][1] ? branches[j][1].toString().trim() : "",
          lat: Number(branches[j][2]) || 0,
          lng: Number(branches[j][3]) || 0,
          radius: Number(branches[j][4]) || 200,
          workStart: formatSheetTime(branches[j][5], "08:30"),
          workEnd: formatSheetTime(branches[j][6], "17:30")
        });
      }
    }

    return {
      success: true,
      branches: branchList
    };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาดในการดึงข้อมูลสาขา: " + err.toString() };
  }
}

// เพิ่มหรือแก้ไขข้อมูลสาขา ตำแหน่งที่ตั้ง รัศมี เวลาเข้างาน-ออกงาน (เฉพาะ Admin)
function saveBranch(req) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    var branchSheet = ss.getSheetByName("Branches");
    ensureUserHeaders(userSheet);
    ensureBranchHeaders(branchSheet);

    var users = userSheet.getDataRange().getValues();
    if (!verifyAdminAccess(users, req.adminEmail)) {
      return { success: false, message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่จัดการสาขาได้" };
    }

    var branchName = (req.branchName || "").toString().trim();
    var lat = Number(req.lat);
    var lng = Number(req.lng);
    var radius = Number(req.radius) || 200;
    var workStart = formatSheetTime(req.workStart, "08:30");
    var workEnd = formatSheetTime(req.workEnd, "17:30");

    if (!branchName) {
      return { success: false, message: "กรุณาระบุชื่อสาขา" };
    }
    if (isNaN(lat) || isNaN(lng)) {
      return { success: false, message: "พิกัดละติจูด (Lat) หรือลองจิจูด (Lng) ไม่ถูกต้อง" };
    }

    var branches = branchSheet.getDataRange().getValues();
    var branchId = (req.branchId || "").toString().trim();

    // กรณีแก้ไขสาขาเดิม
    if (branchId) {
      for (var i = 1; i < branches.length; i++) {
        if (branches[i][0] && branches[i][0].toString().trim() === branchId) {
          var rowIdx = i + 1;
          branchSheet.getRange(rowIdx, 2).setValue(branchName);
          branchSheet.getRange(rowIdx, 3).setValue(lat);
          branchSheet.getRange(rowIdx, 4).setValue(lng);
          branchSheet.getRange(rowIdx, 5).setValue(radius);
          branchSheet.getRange(rowIdx, 6).setNumberFormat("@").setValue(workStart);
          branchSheet.getRange(rowIdx, 7).setNumberFormat("@").setValue(workEnd);
          return {
            success: true,
            message: "อัปเดตข้อมูลสาขา " + branchName + " (" + branchId + ") เรียบร้อยแล้ว"
          };
        }
      }
    }

    // กรณีสร้างสาขาใหม่ (สร้างรหัส BR01, BR02, ... อัตโนมัติหากไม่ได้ระบุ)
    if (!branchId) {
      var maxBr = 0;
      for (var k = 1; k < branches.length; k++) {
        var bStr = (branches[k][0] || "").toString();
        var m = bStr.match(/\d+/);
        if (m) {
          var num = parseInt(m[0], 10);
          if (num > maxBr) maxBr = num;
        }
      }
      branchId = "BR" + ("0" + (maxBr + 1)).slice(-2);
    }

    var nextRow = branchSheet.getLastRow() + 1;
    branchSheet.appendRow([
      branchId,
      branchName,
      lat,
      lng,
      radius,
      workStart,
      workEnd
    ]);
    branchSheet.getRange(nextRow, 6, 1, 2).setNumberFormat("@");
    branchSheet.getRange(nextRow, 6).setValue(workStart);
    branchSheet.getRange(nextRow, 7).setValue(workEnd);

    return {
      success: true,
      message: "เพิ่มสาขาใหม่ " + branchName + " (" + branchId + ") เรียบร้อยแล้ว"
    };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาดในการบันทึกสาขา: " + err.toString() };
  }
}

// ลบสาขา (เฉพาะ Admin)
function deleteBranch(req) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    var branchSheet = ss.getSheetByName("Branches");

    var users = userSheet.getDataRange().getValues();
    if (!verifyAdminAccess(users, req.adminEmail)) {
      return { success: false, message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่ลบสาขาได้" };
    }

    var targetBranchId = (req.branchId || "").toString().trim();
    if (!targetBranchId) {
      return { success: false, message: "ไม่พบรหัสสาขาที่ต้องการลบ" };
    }

    var branches = branchSheet.getDataRange().getValues();
    for (var i = 1; i < branches.length; i++) {
      if (branches[i][0] && branches[i][0].toString().trim() === targetBranchId) {
        var branchName = branches[i][1];
        branchSheet.deleteRow(i + 1);
        return {
          success: true,
          message: "ลบสาขา " + branchName + " (" + targetBranchId + ") เรียบร้อยแล้ว"
        };
      }
    }

    return { success: false, message: "ไม่พบรหัสสาขาในระบบ" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาดในการลบสาขา: " + err.toString() };
  }
}

// ดึงรายชื่อพนักงานทั้งหมดสำหรับหน้าอนุมัติและจัดการสิทธิ์ (เฉพาะ Admin)
function getAllUsersForAdmin(adminEmail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    var branchSheet = ss.getSheetByName("Branches");
    ensureUserHeaders(userSheet);

    var users = userSheet.getDataRange().getValues();
    if (!verifyAdminAccess(users, adminEmail)) {
      return { success: false, message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้น" };
    }

    var branchMap = {};
    if (branchSheet) {
      var branches = branchSheet.getDataRange().getValues();
      for (var b = 1; b < branches.length; b++) {
        if (branches[b][0]) {
          branchMap[branches[b][0]] = branches[b][1] || branches[b][0];
        }
      }
    }

    var userList = [];
    var pendingCount = 0;
    var activeCount = 0;

    for (var i = 1; i < users.length; i++) {
      if (!users[i][0]) continue;
      var status = (users[i][5] || "INACTIVE").toString().trim().toUpperCase();
      var role = (users[i][2] || "user").toString().trim().toLowerCase();
      var otHour = Number(users[i][6]) || 0;
      var otMin = Number(users[i][7]) || (otHour / 60);

      if (status === "INACTIVE") {
        pendingCount++;
      } else if (status === "ACTIVE") {
        activeCount++;
      }

      userList.push({
        userId: users[i][0],
        name: users[i][1] || "",
        role: role,
        branchId: users[i][3] || "",
        branchName: branchMap[users[i][3]] || users[i][3] || "-",
        email: users[i][4] || "",
        status: status,
        otRatePerHour: Number(otHour.toFixed(2)),
        otRatePerMinute: Number(otMin.toFixed(4))
      });
    }

    return {
      success: true,
      pendingCount: pendingCount,
      activeCount: activeCount,
      totalCount: userList.length,
      users: userList
    };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.toString() };
  }
}

// อนุมัติผู้ใช้งาน / จัดการสิทธิ์ Role, Status, สาขา และค่าโอที (เฉพาะ Admin)
function updateUserPermission(req) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    ensureUserHeaders(userSheet);

    var users = userSheet.getDataRange().getValues();
    if (!verifyAdminAccess(users, req.adminEmail)) {
      return { success: false, message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่จัดการสิทธิ์ได้" };
    }

    var targetUserId = req.targetUserId;
    if (!targetUserId) {
      return { success: false, message: "ไม่พบรหัสพนักงานที่ต้องการแก้ไข" };
    }

    for (var r = 1; r < users.length; r++) {
      if (users[r][0] === targetUserId) {
        var rowIndex = r + 1;
        var currentName = users[r][1];
        var currentRole = (users[r][2] || "user").toString().trim().toLowerCase();
        var currentBranch = users[r][3];
        var currentStatus = (users[r][5] || "INACTIVE").toString().trim().toUpperCase();
        var currentOtHour = Number(users[r][6]) || 0;

        var newName = req.name !== undefined ? req.name.toString().trim() : currentName;
        var newRole = req.role !== undefined ? req.role.toString().trim().toLowerCase() : currentRole;
        var newBranch = req.branchId !== undefined ? req.branchId.toString().trim() : currentBranch;
        var newStatus = req.status !== undefined ? req.status.toString().trim().toUpperCase() : currentStatus;
        var newOtHour = req.otRatePerHour !== undefined ? Number(req.otRatePerHour) : currentOtHour;

        if (isNaN(newOtHour) || newOtHour < 0) newOtHour = currentOtHour;

        userSheet.getRange(rowIndex, 2).setValue(newName);
        userSheet.getRange(rowIndex, 3).setValue(newRole);
        userSheet.getRange(rowIndex, 4).setValue(newBranch);
        userSheet.getRange(rowIndex, 6).setValue(newStatus);
        userSheet.getRange(rowIndex, 7).setValue(newOtHour);
        userSheet.getRange(rowIndex, 8).setFormula("=IF(ISNUMBER(G" + rowIndex + "), G" + rowIndex + "/60, 0)");

        return {
          success: true,
          message: "อัปเดตข้อมูลของ " + newName + " (" + newStatus + " | " + newRole.toUpperCase() + " | OT " + newOtHour + " บาท/ชม.) เรียบร้อยแล้ว"
        };
      }
    }

    return { success: false, message: "ไม่พบรหัสพนักงานที่ระบุในระบบ" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// อัปเดตค่าโอทีต่อชั่วโมงของพนักงาน (เฉพาะ Admin)
function updateUserOtRate(req) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    ensureUserHeaders(userSheet);

    var users = userSheet.getDataRange().getValues();
    if (!verifyAdminAccess(users, req.adminEmail)) {
      return { success: false, message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่แก้ไขค่าโอทีได้" };
    }

    var targetUserId = req.targetUserId;
    var otRatePerHour = Number(req.otRatePerHour);
    if (isNaN(otRatePerHour) || otRatePerHour < 0) {
      return { success: false, message: "ค่าโอทีต่อชั่วโมงไม่ถูกต้อง" };
    }

    for (var r = 1; r < users.length; r++) {
      if (users[r][0] === targetUserId) {
        var rowIndex = r + 1;
        userSheet.getRange(rowIndex, 7).setValue(otRatePerHour);
        userSheet.getRange(rowIndex, 8).setFormula("=IF(ISNUMBER(G" + rowIndex + "), G" + rowIndex + "/60, 0)");
        return {
          success: true,
          message: "อัปเดตค่าโอทีของ " + users[r][1] + " เป็น " + otRatePerHour + " บาท/ชม. (" + (otRatePerHour / 60).toFixed(2) + " บาท/นาที) เรียบร้อยแล้ว"
        };
      }
    }

    return { success: false, message: "ไม่พบรหัสพนักงานที่ระบุ" };
  } catch (err) {
    return { success: false, message: err.toString() };
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

// ตรวจสอบและสร้างหัวตาราง Logs ให้รองรับคอลัมน์ OT, ออกก่อน และ เข้าสาย (นาที)
function ensureLogHeaders(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Log ID", "User ID", "Timestamp", "Type", "Branch ID", 
      "Distance (m)", "Photo URL", "Status", "Note", "OT (นาที)", "ออกก่อน (นาที)", "เข้าสาย (นาที)"
    ]);
  } else {
    var lastCol = Math.max(sheet.getLastColumn(), 12);
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headerRow[9] || headerRow[9].toString().trim() === "") {
      sheet.getRange(1, 10).setValue("OT (นาที)");
    }
    if (!headerRow[10] || headerRow[10].toString().trim() === "") {
      sheet.getRange(1, 11).setValue("ออกก่อน (นาที)");
    }
    if (!headerRow[11] || headerRow[11].toString().trim() === "") {
      sheet.getRange(1, 12).setValue("เข้าสาย (นาที)");
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

// ฟังก์ชันดึงค่านาทีเข้าสายจากแถวในตาราง Logs (รองรับทั้งข้อมูลใหม่ในคอลัมน์ L และข้อมูลเก่าที่อยู่ใน Note)
function extractLateMinutesFromRow(row) {
  var colLate = Number(row[11]);
  if (!isNaN(colLate) && colLate > 0) {
    return colLate;
  }
  if (row[7] === "LATE" && row[8]) {
    var noteStr = row[8].toString();
    var hmMatch = noteStr.match(/สาย\s*(?:(\d+)\s*ชม\.\s*)?(?:(\d+)\s*นาที)?/);
    if (hmMatch && (hmMatch[1] || hmMatch[2])) {
      var h = hmMatch[1] ? parseInt(hmMatch[1], 10) : 0;
      var m = hmMatch[2] ? parseInt(hmMatch[2], 10) : 0;
      if (h > 0 || m > 0) return h * 60 + m;
    }
    var timeMatch = noteStr.match(/\((\d{1,2}):(\d{2})\s*>\s*(\d{1,2}):(\d{2})\)/);
    if (timeMatch) {
      var actualMin = parseInt(timeMatch[1], 10) * 60 + parseInt(timeMatch[2], 10);
      var startMin = parseInt(timeMatch[3], 10) * 60 + parseInt(timeMatch[4], 10);
      if (actualMin > startMin) return actualMin - startMin;
    }
  }
  return 0;
}

// 7. ฟังก์ชันบันทึกเวลาสแกนเข้า-ออกงาน (ตรวจสอบสถานะ ACTIVE + กฎ 4 ข้อ + บันทึกนาทีเข้าสาย)
function recordAttendance(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("Users");
    var logSheet = ss.getSheetByName("Logs");
    var branchSheet = ss.getSheetByName("Branches");
    
    // ตรวจสอบสถานะพนักงานว่าได้รับอนุมัติ (ACTIVE) แล้วหรือไม่
    var users = userSheet.getDataRange().getValues();
    var userStatus = "INACTIVE";
    for (var u = 1; u < users.length; u++) {
      if (users[u][0] === payload.userId) {
        userStatus = (users[u][5] || "INACTIVE").toString().trim().toUpperCase();
        break;
      }
    }
    if (userStatus !== "ACTIVE") {
      return {
        success: false,
        message: "บัญชีของคุณยังไม่ได้รับการอนุมัติ (สถานะ: " + userStatus + ") กรุณารอผู้ดูแลระบบอนุมัติก่อนลงเวลา"
      };
    }

    // ตรวจสอบข้อมูลสาขา
    var branches = branchSheet.getDataRange().getValues();
    var targetBranch = null;
    
    for (var i = 1; i < branches.length; i++) {
      if (branches[i][0] && branches[i][0].toString().trim() === payload.branchId) {
        targetBranch = {
          lat: Number(branches[i][2]),
          lng: Number(branches[i][3]),
          radius: Number(branches[i][4]),
          workStart: formatSheetTime(branches[i][5], "08:30"),
          workEnd: formatSheetTime(branches[i][6], "17:30")
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

    ensureLogHeaders(logSheet);

    var logData = logSheet.getDataRange().getValues();

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
            existingCheckInRow = r + 1;
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

    // กฎข้อ 1: สแกนเข้ามากกว่า 1 ครั้งต่อวัน ให้คิดเวลาแรกสุดเท่านั้น
    if (payload.type === "CHECK_IN") {
      if (existingCheckInRow > 0) {
        return {
          success: false,
          message: "วันนี้คุณได้ลงเวลาเข้างานไปแล้วเมื่อเวลา " + existingCheckInTime + " น. (ระบบบันทึกเวลาแรกสุดเท่านั้น ไม่บันทึกซ้ำ)"
        };
      }

      var workStartTimeStr = targetBranch.workStart || "08:30";
      var currentMin = now.getHours() * 60 + now.getMinutes();
      var startParts = workStartTimeStr.split(":");
      var startMin = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);

      var status = "SUCCESS";
      var note = "เข้างานตรงเวลา";
      var lateMinutes = 0;

      if (currentMin > startMin) {
        lateMinutes = currentMin - startMin;
        var lateText = formatMinutesToText(lateMinutes);
        status = "LATE";
        note = "เข้างานสาย " + lateMinutes + " นาที (" + lateText + " | เริ่มงาน " + workStartTimeStr + " สแกนเข้า " + currentTimeStr + ")";
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
        0,           // OT (นาที)
        0,           // ออกก่อน (นาที)
        lateMinutes  // เข้าสาย (นาที)
      ]);

      var checkInMsg = "ลงเวลาเข้างานเรียบร้อยแล้ว (" + currentTimeStr + " น.)";
      if (status === "LATE") {
        checkInMsg += " (เข้าสาย " + lateMinutes + " นาที)";
      }

      return {
        success: true,
        message: checkInMsg,
        status: status,
        lateMinutes: lateMinutes
      };
    }

    // กฎข้อ 2, 3, 4: สแกนออก
    if (payload.type === "CHECK_OUT") {
      var workEndTimeStr = targetBranch.workEnd || "17:30";
      var currentMin = now.getHours() * 60 + now.getMinutes();
      var endParts = workEndTimeStr.split(":");
      var endMin = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);

      var status = "SUCCESS";
      var note = "ออกงานตามเวลาปกติ";
      var otMinutes = 0;
      var earlyMinutes = 0;

      if (currentMin > endMin) {
        otMinutes = currentMin - endMin;
        var otText = formatMinutesToText(otMinutes);
        status = "OVERTIME";
        note = "ทำโอที " + otMinutes + " นาที (" + otText + " | เลิกงาน " + workEndTimeStr + " สแกนออก " + currentTimeStr + ")";
      } else if (currentMin < endMin) {
        earlyMinutes = endMin - currentMin;
        var earlyText = formatMinutesToText(earlyMinutes);
        status = "EARLY_LEAVE";
        note = "ออกก่อนเวลา " + earlyMinutes + " นาที (" + earlyText + " | เลิกงาน " + workEndTimeStr + " สแกนออก " + currentTimeStr + ")";
      }

      var photoUrl = saveImageToDrive(payload.imageBase64, payload.userId);

      if (existingCheckOutRow > 0) {
        logSheet.getRange(existingCheckOutRow, 3).setValue(timestampStr);
        logSheet.getRange(existingCheckOutRow, 6).setValue(distance);
        logSheet.getRange(existingCheckOutRow, 7).setValue(photoUrl);
        logSheet.getRange(existingCheckOutRow, 8).setValue(status);
        logSheet.getRange(existingCheckOutRow, 9).setValue(note);
        logSheet.getRange(existingCheckOutRow, 10).setValue(otMinutes);
        logSheet.getRange(existingCheckOutRow, 11).setValue(earlyMinutes);
        logSheet.getRange(existingCheckOutRow, 12).setValue(0);

        var updateMsg = "อัปเดตเวลาออกงานเป็นเวลาล่าสุด (" + currentTimeStr + " น.) เรียบร้อยแล้ว";
        if (otMinutes > 0) updateMsg += " (โอที " + otMinutes + " นาที)";
        if (earlyMinutes > 0) updateMsg += " (ออกก่อน " + earlyMinutes + " นาที)";

        return {
          success: true,
          message: updateMsg,
          status: status
        };
      } else {
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
          earlyMinutes,
          0
        ]);

        var checkOutMsg = "ลงเวลาออกงานเรียบร้อยแล้ว (" + currentTimeStr + " น.)";
        if (otMinutes > 0) checkOutMsg += " (โอที " + otMinutes + " นาที)";
        if (earlyMinutes > 0) checkOutMsg += " (ออกก่อน " + earlyMinutes + " นาที)";

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
    var userSheet = ss.getSheetByName("Users");
    var users = userSheet.getDataRange().getValues();
    var userStatus = "INACTIVE";
    for (var u = 1; u < users.length; u++) {
      if (users[u][0] === data.userId) {
        userStatus = (users[u][5] || "INACTIVE").toString().trim().toUpperCase();
        break;
      }
    }
    if (userStatus !== "ACTIVE") {
      return { success: false, message: "บัญชีของคุณยังไม่ได้รับอนุมัติ ไม่สามารถยื่นใบลาได้" };
    }

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
  ensureLogHeaders(logSheet);
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
      var lateVal = extractLateMinutesFromRow(data[i]);

      logs.push({
        timestamp: tsStr,
        type: data[i][3],
        distance: data[i][5],
        status: data[i][7],
        note: data[i][8] || "",
        ot: otVal > 0 ? otVal + " นาที" : "",
        early: earlyVal > 0 ? earlyVal + " นาที" : "",
        late: lateVal > 0 ? lateVal + " นาที" : "",
        otMinutes: otVal,
        earlyMinutes: earlyVal,
        lateMinutes: lateVal
      });
      if (logs.length >= 5) break;
    }
  }
  return logs;
}

// 10. ดึงข้อมูลสรุปแดชบอร์ดรายเดือนสำหรับแอดมิน (สรุป ขาด, ลา, มา, สาย, เวลาโอที, ค่าโอที)
function getAdminMonthlySummary(monthStr, adminEmail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();

    if (!monthStr || monthStr.trim() === "") {
      monthStr = Utilities.formatDate(new Date(), tz, "yyyy-MM");
    }

    var userSheet = ss.getSheetByName("Users");
    var logSheet = ss.getSheetByName("Logs");
    var leaveSheet = ss.getSheetByName("Leaves");
    var branchSheet = ss.getSheetByName("Branches");

    ensureUserHeaders(userSheet);
    ensureLogHeaders(logSheet);

    var usersData = userSheet ? userSheet.getDataRange().getValues() : [];
    var logsData = logSheet ? logSheet.getDataRange().getValues() : [];
    var leavesData = leaveSheet ? leaveSheet.getDataRange().getValues() : [];
    var branchesData = branchSheet ? branchSheet.getDataRange().getValues() : [];

    // ตรวจสอบสิทธิ์ Admin
    if (!verifyAdminAccess(usersData, adminEmail)) {
      return {
        success: false,
        message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่สามารถดูแดชบอร์ดสรุปรายเดือนได้"
      };
    }

    var branchMap = {};
    for (var b = 1; b < branchesData.length; b++) {
      if (branchesData[b][0]) {
        branchMap[branchesData[b][0]] = branchesData[b][1] || branchesData[b][0];
      }
    }

    var parts = monthStr.split("-");
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    var daysInMonth = new Date(year, month, 0).getDate();
    var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    var workingDatesUpToToday = [];
    var totalWorkingDaysInMonth = 0;

    for (var d = 1; d <= daysInMonth; d++) {
      var dateObj = new Date(year, month - 1, d);
      var dayOfWeek = dateObj.getDay();
      var dStr = year + "-" + ("0" + month).slice(-2) + "-" + ("0" + d).slice(-2);
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        totalWorkingDaysInMonth++;
        if (dStr <= todayStr) {
          workingDatesUpToToday.push(dStr);
        }
      }
    }

    var staffMap = {};
    var staffOrder = [];

    for (var i = 1; i < usersData.length; i++) {
      var uid = usersData[i][0];
      if (!uid) continue;
      var otHour = Number(usersData[i][6]) || 0;
      var otMin = Number(usersData[i][7]);
      if (isNaN(otMin) || (otMin === 0 && otHour > 0)) {
        otMin = otHour / 60;
      }

      staffMap[uid] = {
        userId: uid,
        name: usersData[i][1] || uid,
        role: (usersData[i][2] || "user").toString(),
        branchId: usersData[i][3] || "",
        branchName: branchMap[usersData[i][3]] || usersData[i][3] || "-",
        email: usersData[i][4] || "",
        status: (usersData[i][5] || "INACTIVE").toString().trim().toUpperCase(),
        otRatePerHour: Number(otHour.toFixed(2)),
        otRatePerMinute: Number(otMin.toFixed(4)),
        presentDatesMap: {},
        leaveDatesMap: {},
        presentDays: 0,
        leaveDays: 0,
        absentDays: 0,
        lateCount: 0,
        lateMinutes: 0,
        earlyLeaveCount: 0,
        earlyLeaveMinutes: 0,
        otMinutes: 0,
        otPayBaht: 0
      };
      staffOrder.push(uid);
    }

    for (var l = 1; l < logsData.length; l++) {
      var row = logsData[l];
      var lUid = row[1];
      var lTs = row[2];
      var lType = row[3];
      var lStatus = row[7];

      if (!lUid || !staffMap[lUid] || !lTs) continue;

      var dateStr = (lTs instanceof Date)
        ? Utilities.formatDate(lTs, tz, "yyyy-MM-dd")
        : lTs.toString().substring(0, 10);

      if (dateStr.substring(0, 7) !== monthStr) continue;

      var st = staffMap[lUid];

      if (lType === "CHECK_IN") {
        st.presentDatesMap[dateStr] = true;
        var lateMin = extractLateMinutesFromRow(row);
        if (lStatus === "LATE" || lateMin > 0) {
          st.lateCount++;
          st.lateMinutes += lateMin;
        }
      } else if (lType === "CHECK_OUT") {
        st.presentDatesMap[dateStr] = true;
        var otMinVal = Number(row[9]) || 0;
        var earlyMinVal = Number(row[10]) || 0;
        if (otMinVal > 0) {
          st.otMinutes += otMinVal;
        }
        if (earlyMinVal > 0 || lStatus === "EARLY_LEAVE") {
          st.earlyLeaveCount++;
          st.earlyLeaveMinutes += earlyMinVal;
        }
      }
    }

    for (var lv = 1; lv < leavesData.length; lv++) {
      var lvRow = leavesData[lv];
      var lvUid = lvRow[1];
      var lvStart = lvRow[3];
      var lvEnd = lvRow[4];
      var lvStatus = (lvRow[6] || "PENDING").toString().toUpperCase();

      if (!lvUid || !staffMap[lvUid] || !lvStart || !lvEnd) continue;
      if (lvStatus === "REJECTED" || lvStatus === "CANCELLED") continue;

      var startStr = (lvStart instanceof Date)
        ? Utilities.formatDate(lvStart, tz, "yyyy-MM-dd")
        : lvStart.toString().substring(0, 10);
      var endStr = (lvEnd instanceof Date)
        ? Utilities.formatDate(lvEnd, tz, "yyyy-MM-dd")
        : lvEnd.toString().substring(0, 10);

      var sParts = startStr.split("-");
      var eParts = endStr.split("-");
      if (sParts.length === 3 && eParts.length === 3) {
        var curDate = new Date(parseInt(sParts[0], 10), parseInt(sParts[1], 10) - 1, parseInt(sParts[2], 10));
        var endDate = new Date(parseInt(eParts[0], 10), parseInt(eParts[1], 10) - 1, parseInt(eParts[2], 10));

        var safetyCounter = 0;
        while (curDate <= endDate && safetyCounter < 366) {
          var curDateStr = Utilities.formatDate(curDate, tz, "yyyy-MM-dd");
          if (curDateStr.substring(0, 7) === monthStr) {
            staffMap[lvUid].leaveDatesMap[curDateStr] = true;
          }
          curDate.setDate(curDate.getDate() + 1);
          safetyCounter++;
        }
      }
    }

    var summaryRows = [];
    var totals = {
      totalStaff: staffOrder.length,
      workingDaysElapsed: workingDatesUpToToday.length,
      totalWorkingDaysInMonth: totalWorkingDaysInMonth,
      presentDays: 0,
      leaveDays: 0,
      absentDays: 0,
      lateCount: 0,
      lateMinutes: 0,
      otMinutes: 0,
      otPayBaht: 0
    };

    for (var k = 0; k < staffOrder.length; k++) {
      var id = staffOrder[k];
      var item = staffMap[id];

      item.presentDays = Object.keys(item.presentDatesMap).length;
      item.leaveDays = Object.keys(item.leaveDatesMap).length;

      var absentCount = 0;
      if (item.status === "ACTIVE") {
        for (var w = 0; w < workingDatesUpToToday.length; w++) {
          var wDate = workingDatesUpToToday[w];
          if (!item.presentDatesMap[wDate] && !item.leaveDatesMap[wDate]) {
            absentCount++;
          }
        }
      }
      item.absentDays = absentCount;

      item.otPayBaht = Number((item.otMinutes * item.otRatePerMinute).toFixed(2));
      item.otHoursText = formatMinutesToText(item.otMinutes);
      item.lateHoursText = formatMinutesToText(item.lateMinutes);

      totals.presentDays += item.presentDays;
      totals.leaveDays += item.leaveDays;
      totals.absentDays += item.absentDays;
      totals.lateCount += item.lateCount;
      totals.lateMinutes += item.lateMinutes;
      totals.otMinutes += item.otMinutes;
      totals.otPayBaht += item.otPayBaht;

      summaryRows.push({
        userId: item.userId,
        name: item.name,
        email: item.email,
        role: item.role,
        status: item.status,
        branchId: item.branchId,
        branchName: item.branchName,
        presentDays: item.presentDays,
        leaveDays: item.leaveDays,
        absentDays: item.absentDays,
        lateCount: item.lateCount,
        lateMinutes: item.lateMinutes,
        lateHoursText: item.lateHoursText,
        earlyLeaveCount: item.earlyLeaveCount,
        earlyLeaveMinutes: item.earlyLeaveMinutes,
        otRatePerHour: item.otRatePerHour,
        otRatePerMinute: item.otRatePerMinute,
        otMinutes: item.otMinutes,
        otHoursText: item.otHoursText,
        otPayBaht: item.otPayBaht
      });
    }

    totals.otPayBaht = Number(totals.otPayBaht.toFixed(2));
    totals.otHoursText = formatMinutesToText(totals.otMinutes);
    totals.lateHoursText = formatMinutesToText(totals.lateMinutes);

    return {
      success: true,
      monthStr: monthStr,
      totals: totals,
      staffSummary: summaryRows
    };
  } catch (err) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาดในการคำนวณแดชบอร์ด: " + err.toString()
    };
  }
}

// 11. ตารางสรุปรายวันสำหรับพนักงาน (User Daily Summary) เลือกดูตามเดือน
function getUserDailySummary(userId, monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();

    if (!userId) {
      return { success: false, message: "ไม่พบรหัสพนักงาน" };
    }
    if (!monthStr || monthStr.trim() === "") {
      monthStr = Utilities.formatDate(new Date(), tz, "yyyy-MM");
    }

    var userSheet = ss.getSheetByName("Users");
    var logSheet = ss.getSheetByName("Logs");
    var leaveSheet = ss.getSheetByName("Leaves");
    var branchSheet = ss.getSheetByName("Branches");

    ensureUserHeaders(userSheet);
    ensureLogHeaders(logSheet);

    var usersData = userSheet ? userSheet.getDataRange().getValues() : [];
    var logsData = logSheet ? logSheet.getDataRange().getValues() : [];
    var leavesData = leaveSheet ? leaveSheet.getDataRange().getValues() : [];
    var branchesData = branchSheet ? branchSheet.getDataRange().getValues() : [];

    var branchMap = {};
    for (var b = 1; b < branchesData.length; b++) {
      if (branchesData[b][0]) {
        branchMap[branchesData[b][0]] = {
          name: branchesData[b][1] || branchesData[b][0],
          workStart: formatSheetTime(branchesData[b][5], "08:30"),
          workEnd: formatSheetTime(branchesData[b][6], "17:30")
        };
      }
    }

    var targetUser = null;
    for (var u = 1; u < usersData.length; u++) {
      if (usersData[u][0] === userId) {
        var otHour = Number(usersData[u][6]) || 0;
        var otMin = Number(usersData[u][7]);
        if (isNaN(otMin) || (otMin === 0 && otHour > 0)) {
          otMin = otHour / 60;
        }
        var bInfo = branchMap[usersData[u][3]] || { name: usersData[u][3] || "-", workStart: "08:30", workEnd: "17:30" };
        targetUser = {
          userId: usersData[u][0],
          name: usersData[u][1] || userId,
          branchId: usersData[u][3] || "",
          branchName: bInfo.name,
          workStart: bInfo.workStart,
          workEnd: bInfo.workEnd,
          otRatePerHour: Number(otHour.toFixed(2)),
          otRatePerMinute: Number(otMin.toFixed(4))
        };
        break;
      }
    }

    if (!targetUser) {
      return { success: false, message: "ไม่พบข้อมูลพนักงานในระบบ" };
    }

    // เก็บข้อมูลการสแกนรายวันของพนักงานในเดือนที่เลือก
    var dailyLogs = {};
    for (var l = 1; l < logsData.length; l++) {
      var row = logsData[l];
      if (row[1] !== userId || !row[2]) continue;

      var ts = row[2];
      var dateStr = (ts instanceof Date)
        ? Utilities.formatDate(ts, tz, "yyyy-MM-dd")
        : ts.toString().substring(0, 10);

      if (dateStr.substring(0, 7) !== monthStr) continue;

      var timeStr = (ts instanceof Date)
        ? Utilities.formatDate(ts, tz, "HH:mm")
        : ts.toString().substring(11, 16);

      if (!dailyLogs[dateStr]) {
        dailyLogs[dateStr] = {
          checkInTime: "",
          checkOutTime: "",
          checkInStatus: "",
          checkOutStatus: "",
          lateMinutes: 0,
          otMinutes: 0,
          earlyMinutes: 0,
          notes: []
        };
      }

      var dLog = dailyLogs[dateStr];
      var lType = row[3];
      var lStatus = row[7];
      var lNote = row[8] ? row[8].toString() : "";

      if (lType === "CHECK_IN") {
        if (!dLog.checkInTime) {
          dLog.checkInTime = timeStr;
          dLog.checkInStatus = lStatus;
          dLog.lateMinutes = extractLateMinutesFromRow(row);
          if (lNote) dLog.notes.push(lNote);
        }
      } else if (lType === "CHECK_OUT") {
        dLog.checkOutTime = timeStr;
        dLog.checkOutStatus = lStatus;
        dLog.otMinutes = Number(row[9]) || 0;
        dLog.earlyMinutes = Number(row[10]) || 0;
        if (lNote) dLog.notes.push(lNote);
      }
    }

    // เก็บข้อมูลวันลาของพนักงานในเดือนที่เลือก
    var leaveDates = {};
    var leaveTypeLabels = { SICK: "ลาป่วย", BUSINESS: "ลากิจ", VACATION: "ลาพักร้อน" };
    for (var lv = 1; lv < leavesData.length; lv++) {
      var lvRow = leavesData[lv];
      if (lvRow[1] !== userId || !lvRow[3] || !lvRow[4]) continue;
      var lvStatus = (lvRow[6] || "PENDING").toString().toUpperCase();
      if (lvStatus === "REJECTED" || lvStatus === "CANCELLED") continue;

      var startStr = (lvRow[3] instanceof Date)
        ? Utilities.formatDate(lvRow[3], tz, "yyyy-MM-dd")
        : lvRow[3].toString().substring(0, 10);
      var endStr = (lvRow[4] instanceof Date)
        ? Utilities.formatDate(lvRow[4], tz, "yyyy-MM-dd")
        : lvRow[4].toString().substring(0, 10);

      var sParts = startStr.split("-");
      var eParts = endStr.split("-");
      if (sParts.length === 3 && eParts.length === 3) {
        var curDate = new Date(parseInt(sParts[0], 10), parseInt(sParts[1], 10) - 1, parseInt(sParts[2], 10));
        var endDate = new Date(parseInt(eParts[0], 10), parseInt(eParts[1], 10) - 1, parseInt(eParts[2], 10));
        var safety = 0;
        while (curDate <= endDate && safety < 366) {
          var cStr = Utilities.formatDate(curDate, tz, "yyyy-MM-dd");
          if (cStr.substring(0, 7) === monthStr) {
            var lTypeKey = (lvRow[2] || "").toString().toUpperCase();
            leaveDates[cStr] = {
              leaveType: leaveTypeLabels[lTypeKey] || lTypeKey || "ลางาน",
              reason: lvRow[5] || "",
              status: lvStatus
            };
          }
          curDate.setDate(curDate.getDate() + 1);
          safety++;
        }
      }
    }

    // สร้างรายการสรุปรายวันในเดือนนั้น (แสดงตั้งแต่วันที่ 1 ถึงวันปัจจุบัน หรือสิ้นเดือนถ้าเป็นเดือนในอดีต)
    var parts = monthStr.split("-");
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    var daysInMonth = new Date(year, month, 0).getDate();
    var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var thaiDays = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];

    var dailyRows = [];
    var summary = {
      presentDays: 0,
      leaveDays: 0,
      absentDays: 0,
      lateCount: 0,
      lateMinutes: 0,
      earlyMinutes: 0,
      otMinutes: 0,
      otPayBaht: 0
    };

    for (var d = 1; d <= daysInMonth; d++) {
      var dStr = year + "-" + ("0" + month).slice(-2) + "-" + ("0" + d).slice(-2);
      var dateObj = new Date(year, month - 1, d);
      var dayOfWeek = dateObj.getDay();
      var isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

      var logItem = dailyLogs[dStr];
      var leaveItem = leaveDates[dStr];

      // ข้ามวันในอนาคตที่ยังมาไม่ถึง เว้นแต่มีการยื่นใบลาล่วงหน้าไว้
      if (dStr > todayStr && !logItem && !leaveItem) {
        continue;
      }

      var statusCode = "WEEKEND";
      var statusLabel = "วันหยุดสุดสัปดาห์";
      var checkInTime = "-";
      var checkOutTime = "-";
      var lateMin = 0;
      var earlyMin = 0;
      var otMinVal = 0;
      var otPayVal = 0;
      var noteText = "";

      if (logItem) {
        summary.presentDays++;
        checkInTime = logItem.checkInTime || "-";
        checkOutTime = logItem.checkOutTime || "-";
        lateMin = logItem.lateMinutes || 0;
        earlyMin = logItem.earlyMinutes || 0;
        otMinVal = logItem.otMinutes || 0;
        otPayVal = Number((otMinVal * targetUser.otRatePerMinute).toFixed(2));

        if (lateMin > 0) {
          summary.lateCount++;
          summary.lateMinutes += lateMin;
          statusCode = "LATE";
          statusLabel = "มาสาย (" + lateMin + " นาที)";
        } else {
          statusCode = "PRESENT";
          statusLabel = "มาทำงานปกติ";
        }

        if (otMinVal > 0) {
          summary.otMinutes += otMinVal;
          summary.otPayBaht += otPayVal;
        }
        if (earlyMin > 0) {
          summary.earlyMinutes += earlyMin;
        }
        noteText = logItem.notes.join(" | ");
      } else if (leaveItem) {
        summary.leaveDays++;
        statusCode = "LEAVE";
        statusLabel = leaveItem.leaveType + (leaveItem.status === "PENDING" ? " (รออนุมัติ)" : "");
        noteText = leaveItem.reason ? "เหตุผล: " + leaveItem.reason : "";
      } else if (!isWeekend && dStr <= todayStr) {
        summary.absentDays++;
        statusCode = "ABSENT";
        statusLabel = "ขาดงาน";
        noteText = "ไม่พบประวัติสแกนเข้า-ออกงาน";
      } else {
        // วันหยุดเสาร์-อาทิตย์ที่ไม่ได้มาทำโอที ไม่ต้องแสดงก็ได้หรือแสดงเป็นวันหยุด
        continue;
      }

      dailyRows.push({
        dateStr: dStr,
        dayName: thaiDays[dayOfWeek],
        statusCode: statusCode,
        statusLabel: statusLabel,
        checkInTime: checkInTime,
        checkOutTime: checkOutTime,
        lateMinutes: lateMin,
        earlyMinutes: earlyMin,
        otMinutes: otMinVal,
        otPayBaht: otPayVal,
        note: noteText
      });
    }

    // เรียงลำดับวันที่ล่าสุดขึ้นก่อน
    dailyRows.reverse();
    summary.otPayBaht = Number(summary.otPayBaht.toFixed(2));

    return {
      success: true,
      monthStr: monthStr,
      user: targetUser,
      summary: summary,
      dailyRows: dailyRows
    };
  } catch (err) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาดในการดึงตารางสรุปรายวัน: " + err.toString()
    };
  }
}

// 12. ตารางสรุปรายวันของพนักงานทุกคนสำหรับแอดมิน (Admin Daily Summary) เลือกดูตามวันที่และสาขา
function getAdminDailySummary(dateStr, adminEmail, filterBranchId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();

    if (!dateStr || dateStr.trim() === "") {
      dateStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    }

    var userSheet = ss.getSheetByName("Users");
    var logSheet = ss.getSheetByName("Logs");
    var leaveSheet = ss.getSheetByName("Leaves");
    var branchSheet = ss.getSheetByName("Branches");

    ensureUserHeaders(userSheet);
    ensureLogHeaders(logSheet);

    var usersData = userSheet ? userSheet.getDataRange().getValues() : [];
    var logsData = logSheet ? logSheet.getDataRange().getValues() : [];
    var leavesData = leaveSheet ? leaveSheet.getDataRange().getValues() : [];
    var branchesData = branchSheet ? branchSheet.getDataRange().getValues() : [];

    if (!verifyAdminAccess(usersData, adminEmail)) {
      return {
        success: false,
        message: "ปฏิเสธการเข้าถึง: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่สามารถดูตารางสรุปรายวันทั้งหมดได้"
      };
    }

    var branchMap = {};
    for (var b = 1; b < branchesData.length; b++) {
      if (branchesData[b][0]) {
        branchMap[branchesData[b][0]] = {
          name: branchesData[b][1] || branchesData[b][0],
          workStart: formatSheetTime(branchesData[b][5], "08:30"),
          workEnd: formatSheetTime(branchesData[b][6], "17:30")
        };
      }
    }

    // ดึงข้อมูลสแกนของทุกคนในวันที่เลือก
    var logsByUser = {};
    for (var l = 1; l < logsData.length; l++) {
      var row = logsData[l];
      var lUid = row[1];
      var lTs = row[2];
      if (!lUid || !lTs) continue;

      var rowDateStr = (lTs instanceof Date)
        ? Utilities.formatDate(lTs, tz, "yyyy-MM-dd")
        : lTs.toString().substring(0, 10);

      if (rowDateStr !== dateStr) continue;

      var timeStr = (lTs instanceof Date)
        ? Utilities.formatDate(lTs, tz, "HH:mm")
        : lTs.toString().substring(11, 16);

      if (!logsByUser[lUid]) {
        logsByUser[lUid] = {
          checkInTime: "",
          checkOutTime: "",
          checkInPhoto: "",
          checkOutPhoto: "",
          lateMinutes: 0,
          otMinutes: 0,
          earlyMinutes: 0,
          notes: []
        };
      }

      var uLog = logsByUser[lUid];
      var lType = row[3];
      var photoUrl = row[6] ? row[6].toString() : "";
      var noteStr = row[8] ? row[8].toString() : "";

      if (lType === "CHECK_IN") {
        if (!uLog.checkInTime) {
          uLog.checkInTime = timeStr;
          uLog.checkInPhoto = photoUrl;
          uLog.lateMinutes = extractLateMinutesFromRow(row);
          if (noteStr) uLog.notes.push(noteStr);
        }
      } else if (lType === "CHECK_OUT") {
        uLog.checkOutTime = timeStr;
        uLog.checkOutPhoto = photoUrl;
        uLog.otMinutes = Number(row[9]) || 0;
        uLog.earlyMinutes = Number(row[10]) || 0;
        if (noteStr) uLog.notes.push(noteStr);
      }
    }

    // ดึงข้อมูลการลาที่มีผลในวันที่เลือก
    var leaveByUser = {};
    var leaveTypeLabels = { SICK: "ลาป่วย", BUSINESS: "ลากิจ", VACATION: "ลาพักร้อน" };
    for (var lv = 1; lv < leavesData.length; lv++) {
      var lvRow = leavesData[lv];
      var lvUid = lvRow[1];
      if (!lvUid || !lvRow[3] || !lvRow[4]) continue;
      var lvStatus = (lvRow[6] || "PENDING").toString().toUpperCase();
      if (lvStatus === "REJECTED" || lvStatus === "CANCELLED") continue;

      var startStr = (lvRow[3] instanceof Date)
        ? Utilities.formatDate(lvRow[3], tz, "yyyy-MM-dd")
        : lvRow[3].toString().substring(0, 10);
      var endStr = (lvRow[4] instanceof Date)
        ? Utilities.formatDate(lvRow[4], tz, "yyyy-MM-dd")
        : lvRow[4].toString().substring(0, 10);

      if (dateStr >= startStr && dateStr <= endStr) {
        var lTypeKey = (lvRow[2] || "").toString().toUpperCase();
        leaveByUser[lvUid] = {
          leaveType: leaveTypeLabels[lTypeKey] || lTypeKey || "ลางาน",
          reason: lvRow[5] || "",
          status: lvStatus
        };
      }
    }

    var totals = {
      totalStaff: 0,
      presentCount: 0,
      onTimeCount: 0,
      lateCount: 0,
      leaveCount: 0,
      absentCount: 0,
      totalLateMinutes: 0,
      totalOtMinutes: 0,
      totalOtPayBaht: 0
    };

    var staffDailyRows = [];

    for (var i = 1; i < usersData.length; i++) {
      var uid = usersData[i][0];
      if (!uid) continue;

      var uBranchId = (usersData[i][3] || "").toString().trim();
      if (filterBranchId && filterBranchId !== "ALL" && uBranchId !== filterBranchId) {
        continue;
      }

      var uStatus = (usersData[i][5] || "INACTIVE").toString().trim().toUpperCase();
      var uName = usersData[i][1] || uid;
      var uRole = (usersData[i][2] || "user").toString();
      var otHour = Number(usersData[i][6]) || 0;
      var otMinRate = Number(usersData[i][7]);
      if (isNaN(otMinRate) || (otMinRate === 0 && otHour > 0)) {
        otMinRate = otHour / 60;
      }

      var bInfo = branchMap[uBranchId] || { name: uBranchId || "-", workStart: "08:30", workEnd: "17:30" };
      var uLog = logsByUser[uid];
      var uLeave = leaveByUser[uid];

      totals.totalStaff++;

      var dayStatusCode = "ABSENT";
      var dayStatusLabel = "ขาดงาน / ยังไม่สแกน";
      var checkInTime = "-";
      var checkOutTime = "-";
      var lateMin = 0;
      var earlyMin = 0;
      var otMin = 0;
      var otPay = 0;
      var checkInPhoto = "";
      var checkOutPhoto = "";
      var noteText = "";

      if (uLog) {
        totals.presentCount++;
        checkInTime = uLog.checkInTime || "-";
        checkOutTime = uLog.checkOutTime || "-";
        checkInPhoto = uLog.checkInPhoto || "";
        checkOutPhoto = uLog.checkOutPhoto || "";
        lateMin = uLog.lateMinutes || 0;
        earlyMin = uLog.earlyMinutes || 0;
        otMin = uLog.otMinutes || 0;
        otPay = Number((otMin * otMinRate).toFixed(2));

        if (lateMin > 0) {
          totals.lateCount++;
          totals.totalLateMinutes += lateMin;
          dayStatusCode = "LATE";
          dayStatusLabel = "สาย (" + lateMin + " นาที)";
        } else {
          totals.onTimeCount++;
          dayStatusCode = "ONTIME";
          dayStatusLabel = "มาตรงเวลา";
        }

        if (otMin > 0) {
          totals.totalOtMinutes += otMin;
          totals.totalOtPayBaht += otPay;
        }
        noteText = uLog.notes.join(" | ");
      } else if (uLeave) {
        totals.leaveCount++;
        dayStatusCode = "LEAVE";
        dayStatusLabel = uLeave.leaveType + (uLeave.status === "PENDING" ? " (รออนุมัติ)" : "");
        noteText = uLeave.reason ? "เหตุผล: " + uLeave.reason : "";
      } else if (uStatus === "INACTIVE") {
        dayStatusCode = "INACTIVE";
        dayStatusLabel = "รออนุมัติสิทธิ์ (INACTIVE)";
      } else {
        totals.absentCount++;
      }

      staffDailyRows.push({
        userId: uid,
        name: uName,
        role: uRole,
        userStatus: uStatus,
        branchId: uBranchId,
        branchName: bInfo.name,
        workSchedule: bInfo.workStart + "-" + bInfo.workEnd,
        dayStatusCode: dayStatusCode,
        dayStatusLabel: dayStatusLabel,
        checkInTime: checkInTime,
        checkOutTime: checkOutTime,
        checkInPhoto: checkInPhoto,
        checkOutPhoto: checkOutPhoto,
        lateMinutes: lateMin,
        earlyMinutes: earlyMin,
        otRatePerHour: Number(otHour.toFixed(2)),
        otRatePerMinute: Number(otMinRate.toFixed(4)),
        otMinutes: otMin,
        otPayBaht: otPay,
        note: noteText
      });
    }

    totals.totalOtPayBaht = Number(totals.totalOtPayBaht.toFixed(2));

    return {
      success: true,
      dateStr: dateStr,
      totals: totals,
      staffDailyRows: staffDailyRows
    };
  } catch (err) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาดในการดึงตารางสรุปรายวันแอดมิน: " + err.toString()
    };
  }
}
