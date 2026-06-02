// =========================================================================
// 1. KHU VỰC CẤU HÌNH (SỬA CÁC THÔNG SỐ NÀY KHI ÁP DỤNG CHO KHU VỰC KHÁC)
// =========================================================================
const CONFIG = {
  // Tên khu vực sẽ hiện trên thông báo Chatbot (VD: "KHU VỰC PKG", "KHU VỰC MIXING"...)
  AREA_NAME: "KHU VỰC MIXING", 
  
  // Link Webhook của Google Chat (Tạo riêng cho từng nhóm/khu vực nếu cần)
  WEBHOOK_URL: "https://chat.googleapis.com/v1/spaces/AAAA_BQmvKU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=WGp6pMLfJOl33DRKLvCYbBkl4svl1faI4LrrRR4_iBs",
  
  // Các khung giờ muốn Bot réo tên báo cáo (0-23)
  ALERT_HOURS: [8, 13, 15, 21, 0, 5], 

  // Bản đồ số thứ tự Cột dữ liệu trong Sheet "DATA" (Cột A = 0, Cột B = 1,...)
  COLUMNS: {
    JOB_ID: 0,            // Cột jobId
    STATUS: 6,           // Cột trạng thái (completed/paused)
    CREATED_AT: 7,       // Cột Thời gian tạo
    MACHINE: 20,         // Cột Máy
    BUSINESS_DATE: 21,   // Cột data.Ngày thực hiện
    EMAIL: 22,           // Cột Người thực hiện
    GROUP: 23,           // Cột Nhóm
    MACHINE_STATUS: 24   // Cột Tình trạng máy (Có kế hoạch...)
  }
};

// =========================================================================
// -------------------- KẾT THÚC CẤU HÌNH ----------------------------------
// =========================================================================

// Hàm phụ trợ giúp ép mọi loại định dạng ngày về chuẩn duy nhất yyyy-MM-dd
function getSafeDateString(rawDate, timeZone) {
  if (!rawDate) return "";
  if (rawDate instanceof Date) {
    try {
      timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch(e) {}
    return Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd");
  }

  var text = rawDate.toString().trim();
  if (!text) return "";

  var isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    return isoMatch[1] + "-" + isoMatch[2].padStart(2, "0") + "-" + isoMatch[3].padStart(2, "0");
  }

  var dmyMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmyMatch) {
    return dmyMatch[3] + "-" + dmyMatch[2].padStart(2, "0") + "-" + dmyMatch[1].padStart(2, "0");
  }

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    try {
      timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch(e) {}
    return Utilities.formatDate(parsed, timeZone, "yyyy-MM-dd");
  }

  return text;
}

// 1. Hàm cài đặt Trigger (Chạy mỗi giờ)
function triggerSyncKPI() {
  var hour = new Date().getHours();
  
  // Gọi hàm update() trước để lấy dữ liệu mới nhất
  try {
    if (typeof update === "function") {
      update();
      // Lệnh cực kỳ quan trọng: Ép hệ thống cập nhật hết mọi thay đổi từ hàm update()
      // xuống mặt Sheet thực tế trước khi code chạy tiếp để không bị đọc sót data.
      SpreadsheetApp.flush(); 
    }
  } catch (e) {
    Logger.log("Lỗi khi chạy hàm update(), hoặc hàm không tồn tại: " + e.toString());
  }
  
  // Sau khi chắc chắn dữ liệu đã được cập nhật, tiến hành đồng bộ và báo cáo
  if (CONFIG.ALERT_HOURS.indexOf(hour) !== -1) {
    runSyncAndNotify(); 
  } else {
    syncHistoricalKPI_Fast(); 
  }
}

// 2. Hàm Gửi tin nhắn Bot
function sendGoogleChat(message) {
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({ "text": message })
  };
  UrlFetchApp.fetch(CONFIG.WEBHOOK_URL, options);
}

// 3. Hàm tổng hợp và Kiểm tra Máy thiếu & Máy chưa duyệt
function runSyncAndNotify() {
  // Chạy đồng bộ dữ liệu trước khi quét báo cáo
  syncHistoricalKPI_Fast();
  SpreadsheetApp.flush(); 
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kpiSheet = ss.getSheetByName("KPI_New");
  var kpiData = kpiSheet.getDataRange().getValues(); 
  
  var shiftInfo = getCurrentShift(); 
  var missingMachines = [];          // Máy chưa làm gì cả
  var inProgressMachines = [];       // Máy đang thực hiện (created/in-progress)
  var pendingApprovalMachines = [];  // Máy chờ duyệt (paused)
  var totalMachinesInShift = 0; 
  
  for (var i = 1; i < kpiData.length; i++) {
    var row = kpiData[i];
    
    var rowDate = getSafeDateString(row[0], "Asia/Ho_Chi_Minh");
    var rowShift = row[1];
    var rowMachine = row[2];
    var rowStatus = row[6]; 
    var rowApproval = row[7]; 

    if (rowDate === shiftInfo.date && rowShift === shiftInfo.shift) {
      totalMachinesInShift++; 
      
      // Phân loại trạng thái để Bot báo cáo chính xác
      if (rowStatus === "Chưa hoàn thành") {
        if (missingMachines.indexOf(rowMachine) === -1) missingMachines.push(rowMachine);
      } else if (rowStatus === "Đang thực hiện") {
        if (inProgressMachines.indexOf(rowMachine) === -1) inProgressMachines.push(rowMachine);
      } else if (rowStatus === "Hoàn thành" && rowApproval === "Chưa duyệt") {
        if (pendingApprovalMachines.indexOf(rowMachine) === -1) pendingApprovalMachines.push(rowMachine);
      }
    }
  }

  // Nhận diện ngày nghỉ
  if (totalMachinesInShift > 0 && missingMachines.length === totalMachinesInShift) {
    Logger.log(CONFIG.AREA_NAME + " hôm nay có vẻ là ngày nghỉ (0/" + totalMachinesInShift + " máy). Bot im lặng.");
    return; 
  }

  // Nếu có máy chưa làm, đang làm dở HOẶC chưa duyệt thì đều gửi thông báo
  if (missingMachines.length > 0 || pendingApprovalMachines.length > 0 || inProgressMachines.length > 0) {
    var msg = "<users/all>\n⚠️ *TÌNH TRẠNG CHECKLIST - " + CONFIG.AREA_NAME + "* ⚠️\n" +
              "📅 Ngày: " + shiftInfo.date + "\n" +
              "⏰ Ca: " + shiftInfo.shift + "\n" +
              "----------------------------------\n";
    
    // Báo cáo mảng chưa đánh Checklist
    if (missingMachines.length > 0) {
      msg += "❌ *CÁC MÁY CHƯA ĐƯỢC KIỂM TRA:*\n👉 " + missingMachines.join(", ") + "\n\n";
    }

    // Báo cáo mảng Đang thực hiện
    if (inProgressMachines.length > 0) {
      msg += "🔄 *CÁC MÁY ĐANG KIỂM TRA (CHƯA XONG):*\n👉 " + inProgressMachines.join(", ") + "\n\n";
    }

    // Báo cáo mảng chờ Leader duyệt
    if (pendingApprovalMachines.length > 0) {
      msg += "⏳ *CÁC MÁY ĐÃ KIỂM TRA NHƯNG CHỜ LEADER DUYỆT:*\n👉 " + pendingApprovalMachines.join(", ") + "\n\n";
    }

    msg += "📢 Anh em và Leader lưu ý hoàn thành đúng tiến độ nhé!";
    sendGoogleChat(msg);
  } else if (totalMachinesInShift > 0) {
    // Nếu tất cả các mảng trên đều trống nghĩa là 100% hoàn thành và đã duyệt
    var msgSuccess = "<users/all>\n🎉 *TIN VUI CHECKLIST - " + CONFIG.AREA_NAME + "* 🎉\n" +
                     "📅 Ngày: " + shiftInfo.date + "\n" +
                     "⏰ Ca: " + shiftInfo.shift + "\n" +
                     "----------------------------------\n" +
                     "✅ Tuyệt vời! Toàn bộ " + totalMachinesInShift + " máy trong ca đều đã hoàn thành và được duyệt 100%.\n" +
                     "💪 Cảm ơn anh em đã làm việc chăm chỉ. Chúc mọi người hoàn thành ca làm việc an toàn và hiệu quả! 🚀";
    sendGoogleChat(msgSuccess);
  }
}

// 4. Hàm đồng bộ dữ liệu phiên bản Toàn diện
function syncHistoricalKPI_Fast() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master_data");
  var dataSheet = ss.getSheetByName("DATA");
  var kpiSheet = ss.getSheetByName("KPI_New");

  var lastRowData = dataSheet.getLastRow();
  if (lastRowData < 2) return; 

  // ĐỌC TOÀN BỘ DỮ LIỆU ĐỂ KHÔNG BAO GIỜ SÓT (Thay vì chỉ lấy 500 dòng)
  var dataValues = dataSheet.getRange(2, 1, lastRowData - 1, dataSheet.getLastColumn()).getValues();

  var masterData = masterSheet.getRange("A2:A" + masterSheet.getLastRow()).getValues().flat().filter(String);
  var masterSet = new Set(masterData);

  var completedRecords = {};
  var affectedDates = new Set(); 
  var timeZone = "Asia/Ho_Chi_Minh";

  var now = new Date();
  var nowTimeStr = Utilities.formatDate(now, timeZone, "HH:mm:ss");
  var nowDateStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  if (nowTimeStr <= '05:54:59') {
     var prev = new Date(now.getTime() - 24 * 3600000);
     nowDateStr = Utilities.formatDate(prev, timeZone, "yyyy-MM-dd");
  }
  affectedDates.add(nowDateStr); 

  // XỬ LÝ TOÀN BỘ TRẠNG THÁI (Ghi nhận trạng thái cuối cùng của máy)
  dataValues.forEach(row => {
    var status = row[CONFIG.COLUMNS.STATUS];
    var timestamp = getOperationalTimestamp(row[CONFIG.COLUMNS.CREATED_AT], row[CONFIG.COLUMNS.JOB_ID]);
    var machine = row[CONFIG.COLUMNS.MACHINE];

    // Chỉ bỏ qua nếu không có thời gian hoặc máy không có trong Master
    if (!timestamp || !masterSet.has(machine)) return;

    var shiftInfo = getShiftInfoFromTimestamp(timestamp, timeZone);
    if (!shiftInfo) return;
    var dateStr = getSafeDateString(row[CONFIG.COLUMNS.BUSINESS_DATE], timeZone) || shiftInfo.date;
    var shift = shiftInfo.shift;

    if (!completedRecords[dateStr]) completedRecords[dateStr] = {};
    if (!completedRecords[dateStr][shift]) completedRecords[dateStr][shift] = {};
    
    // Lưu lại thông tin và status gốc (có thể ghi đè nhiều lần, sẽ lấy giá trị cuối cùng/mới nhất)
    completedRecords[dateStr][shift][machine] = {
      email: row[CONFIG.COLUMNS.EMAIL] || "-",
      group: row[CONFIG.COLUMNS.GROUP] || "-",
      machineStatus: row[CONFIG.COLUMNS.MACHINE_STATUS] || "-",
      rawStatus: status 
    };
    affectedDates.add(dateStr);
  });

  // Lấy dữ liệu KPI_New hiện tại để cập nhật
  var kpiFullData = kpiSheet.getDataRange().getValues();
  var rawHeader = kpiFullData.shift(); 
  var kpiHeader = rawHeader ? rawHeader.slice(0, 8) : ["Ngày", "Ca", "Máy", "Nhóm", "Người thực hiện", "Tình trạng máy", "Trạng thái", "Phê duyệt"]; 
  while(kpiHeader.length < 8) kpiHeader.push(""); 
  if (!kpiHeader[7]) kpiHeader[7] = "Phê duyệt"; 

  var kpiMap = new Map();
  kpiFullData.forEach(row => {
    if (!row[0]) return; 
    
    var rowDate = getSafeDateString(row[0], timeZone);
    var key = rowDate + "_" + row[1] + "_" + row[2];
    
    var cleanRow = row.slice(0, 8); 
    while(cleanRow.length < 8) cleanRow.push("");
    kpiMap.set(key, cleanRow);
  });

  // ÁNH XẠ VÀO KPI_NEW VỚI LOGIC STATUS CHUẨN XÁC
  affectedDates.forEach(dStr => {
    ['Ca 1', 'Ca 2', 'Ca 3'].forEach(sStr => {
      masterData.forEach(mach => {
        var record = (completedRecords[dStr] && completedRecords[dStr][sStr]) ? completedRecords[dStr][sStr][mach] : null;
        var existingKey = dStr + "_" + sStr + "_" + mach;
        var existingRow = kpiMap.get(existingKey);
        
        var group = "-", email = "-", machineStatus = "-", finalStatus = "Chưa hoàn thành", approvalStatus = "-";

        if (record) {
          group = record.group;
          email = record.email;
          machineStatus = record.machineStatus;

          // QUY ĐỔI STATUS
          if (record.rawStatus === 'completed') {
            finalStatus = "Hoàn thành";
            approvalStatus = "Đã duyệt";
          } else if (record.rawStatus === 'paused') {
            finalStatus = "Hoàn thành";
            approvalStatus = "Chưa duyệt";
          } else if (record.rawStatus === 'in-progress' || record.rawStatus === 'created') {
            finalStatus = "Đang thực hiện";
            approvalStatus = "-";
          } else {
            finalStatus = record.rawStatus;
          }
        } 
        // Nếu không có dữ liệu mới, nhưng dòng cũ trong KPI_New đang ghi nhận đã làm thì giữ lại lịch sử
        else if (existingRow && existingRow[6] !== "Chưa hoàn thành") {
          group = existingRow[3];
          email = existingRow[4];
          machineStatus = existingRow[5];
          finalStatus = existingRow[6];
          approvalStatus = existingRow[7]; 
        }

        var rowData = [dStr, sStr, mach, group, email, machineStatus, finalStatus, approvalStatus];
        kpiMap.set(existingKey, rowData);
      });
    });
  });

  // Ghi toàn bộ dữ liệu mới xuống Sheet KPI_New
  var finalData = [kpiHeader, ...Array.from(kpiMap.values())];
  var requiredRows = finalData.length;
  var currentMaxRows = kpiSheet.getMaxRows();

  if (requiredRows > currentMaxRows) {
    kpiSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows + 100);
  }

  var lastRow = kpiSheet.getLastRow();
  if (lastRow > 0) {
    kpiSheet.getRange(1, 1, lastRow, 8).clearContent(); 
  }
  
  kpiSheet.getRange(1, 1, requiredRows, 8).setValues(finalData); 
  Logger.log("Đã đồng bộ xong dữ liệu (" + CONFIG.AREA_NAME + ")!");
}

// 5. Hàm xác định Ca
function getCurrentShift() {
  var timeZone = "Asia/Ho_Chi_Minh";
  var now = new Date();
  var timeStr = Utilities.formatDate(now, timeZone, "HH:mm:ss");
  var dateStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  var shift = "Ca 3";

  if (timeStr >= '06:00:00' && timeStr <= '13:59:59') shift = 'Ca 1';
  else if (timeStr >= '14:00:00' && timeStr <= '21:59:59') shift = 'Ca 2';
  else {
    shift = 'Ca 3';
    if (timeStr <= '05:59:59') {
      var prevDate = new Date(now.getTime() - 24 * 3600000);
      dateStr = Utilities.formatDate(prevDate, timeZone, "yyyy-MM-dd");
    }
  }
  return { date: dateStr, shift: shift };
}

function getOperationalTimestamp(value, jobId) {
  var parsed = parseFlexibleDateKpi(value);
  var objectIdDate = dateFromObjectId(jobId);
  if (!parsed) return objectIdDate;

  var isDateOnly = parsed.getHours() === 0 && parsed.getMinutes() === 0 && parsed.getSeconds() === 0;
  if (isDateOnly && objectIdDate) return objectIdDate;
  return parsed;
}

function dateFromObjectId(jobId) {
  if (!jobId) return null;
  var text = jobId.toString().trim();
  if (!/^[0-9a-fA-F]{8}/.test(text)) return null;
  return new Date(parseInt(text.slice(0, 8), 16) * 1000);
}

function parseFlexibleDateKpi(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  var text = value.toString().trim();
  if (!text) return null;

  var isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?(?:\s*([+-]\d{2}):?(\d{2})|Z)?$/);
  if (isoMatch) {
    var year = Number(isoMatch[1]);
    var month = Number(isoMatch[2]) - 1;
    var day = Number(isoMatch[3]);
    var hour = Number(isoMatch[4] || 0);
    var minute = Number(isoMatch[5] || 0);
    var second = Number(isoMatch[6] || 0);
    if (text.slice(-1) === "Z") return new Date(Date.UTC(year, month, day, hour, minute, second));
    return new Date(year, month, day, hour, minute, second);
  }

  var dmyMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (dmyMatch) {
    return new Date(
      Number(dmyMatch[3]), Number(dmyMatch[2]) - 1, Number(dmyMatch[1]),
      Number(dmyMatch[4] || 0), Number(dmyMatch[5] || 0), Number(dmyMatch[6] || 0)
    );
  }

  var nativeDate = new Date(text);
  return isNaN(nativeDate.getTime()) ? null : nativeDate;
}

function getShiftInfoFromTimestamp(timestamp, timeZone) {
  if (!timestamp || isNaN(timestamp.getTime())) return null;
  var timeStr = Utilities.formatDate(timestamp, timeZone, "HH:mm:ss");
  var dateStr = Utilities.formatDate(timestamp, timeZone, "yyyy-MM-dd");
  var shift = "Ca 3";

  if (timeStr >= '06:00:00' && timeStr <= '13:59:59') {
    shift = 'Ca 1';
  } else if (timeStr >= '14:00:00' && timeStr <= '21:59:59') {
    shift = 'Ca 2';
  } else if (timeStr <= '05:59:59') {
    var prevDate = new Date(timestamp.getTime() - 24 * 3600000);
    dateStr = Utilities.formatDate(prevDate, timeZone, "yyyy-MM-dd");
  }

  return { date: dateStr, shift: shift };
}
