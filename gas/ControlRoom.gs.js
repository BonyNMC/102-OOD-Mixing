// =========================================================================
// CONTROL ROOM DASHBOARD — Backend Engine
// =========================================================================
// Mục đích: Xử lý dữ liệu cho Dashboard Control Room, hiển thị trạng thái
// real-time của Ca hiện tại.
//
// Rule chuẩn hóa: Khi áp dụng cho khu vực khác, chỉ cần sửa CR_CONFIG.
// =========================================================================

// =========================================================================
// 1. CẤU HÌNH CONTROL ROOM (SỬA KHI ÁP DỤNG CHO KHU VỰC KHÁC)
// =========================================================================
const CR_CONFIG = {
  AREA_NAME: "MIXING",
  TIME_ZONE: "Asia/Ho_Chi_Minh",
  CACHE_TTL_SECONDS: 180, // Cache 3 phút

  // Quy tắc Ca — giống Update KPI (cHÍNH).js
  SHIFTS: {
    "Ca 1": { start: "05:55:00", end: "13:54:59", label: "Ca 1", number: "1" },
    "Ca 2": { start: "13:55:00", end: "21:54:59", label: "Ca 2", number: "2" },
    "Ca 3": { start: "21:55:00", end: "05:54:59", label: "Ca 3", number: "3" },
  },

  // Panel 1: Machine Checklist
  CHECKLIST: {
    enabled: true,
    dataSheet: "DATA",
    masterSheet: "Master_data",
  },

  // Panel 2: Changeover Compliance
  CHANGEOVER: {
    enabled: true,
    transactionSheet: "DATA_MixingTransaction",
    changeoverSheet: "DATA_Changeover-Mixing",
  },

  // Panel 3: Error Alerts
  ERRORS: {
    enabled: true,
    errorSheet: "History_Error_Mixing_none_Match_Data",
  },

  // Snapshot History
  SNAPSHOT: {
    enabled: true,
    sheetName: "CR_History_Snapshot",
  },
};

// =========================================================================
// 2. WEB APP ENTRY POINT
// =========================================================================

/**
 * Serve Dashboard HTML khi truy cập Web App URL.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Control Room - ' + CR_CONFIG.AREA_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API chính cho Dashboard. Trả về JSON chứa toàn bộ dữ liệu hiển thị.
 * Gọi từ client qua google.script.run.getDashboardData().
 * @param {Object} params - { date: "yyyy-MM-dd", shift: "Ca 1" } (optional)
 * @returns {string} JSON string
 */
function getDashboardData(params) {
  var shiftInfo = params && params.date && params.shift
    ? { date: params.date, shift: params.shift }
    : _getCurrentShiftCR();

  var cacheKey = "CR_" + CR_CONFIG.AREA_NAME + "_" + shiftInfo.date + "_" + shiftInfo.shift;
  var cache = CacheService.getScriptCache();
  var cached = (params && params.forceRefresh === true) ? null : cache.get(cacheKey);

  if (cached) {
    try {
      return cached;
    } catch (e) {
      // Cache corrupted, recompute
    }
  }

  var result = _buildDashboardPayload(shiftInfo);
  var jsonStr = JSON.stringify(result);

  // CacheService max value size = 100KB, split nếu cần
  try {
    cache.put(cacheKey, jsonStr, CR_CONFIG.CACHE_TTL_SECONDS);
  } catch (e) {
    Logger.log("Cache put failed (data too large?): " + e.toString());
  }

  return jsonStr;
}

// =========================================================================
// 3. PAYLOAD BUILDER
// =========================================================================

function _buildDashboardPayload(shiftInfo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var payload = {
    area: CR_CONFIG.AREA_NAME,
    date: shiftInfo.date,
    shift: shiftInfo.shift,
    generatedAt: Utilities.formatDate(new Date(), CR_CONFIG.TIME_ZONE, "yyyy-MM-dd HH:mm:ss"),
    checklist: null,
    changeover: null,
    errors: null,
    summary: { total: 0, done: 0, inProgress: 0, missing: 0, pendingApproval: 0 },
  };

  // Panel 1: Checklist
  if (CR_CONFIG.CHECKLIST.enabled) {
    payload.checklist = _getChecklistStatus(ss, shiftInfo.date, shiftInfo.shift);
    // Tính summary từ checklist
    if (payload.checklist) {
      payload.summary.total = payload.checklist.length;
      payload.checklist.forEach(function (item) {
        if (item.status === "Hoàn thành") payload.summary.done++;
        else if (item.status === "Đang thực hiện") payload.summary.inProgress++;
        else if (item.status === "Chưa hoàn thành") payload.summary.missing++;
      });
      payload.summary.pendingApproval = payload.checklist.filter(function (item) {
        return item.status === "Hoàn thành" && item.approval === "Chưa duyệt";
      }).length;
    }
  }

  // Panel 2: Changeover
  if (CR_CONFIG.CHANGEOVER.enabled) {
    payload.changeover = _getChangeoverCompliance(ss, shiftInfo.date, shiftInfo.shift);
  }

  // Panel 3: Errors
  if (CR_CONFIG.ERRORS.enabled) {
    payload.errors = _getErrorAlerts(ss, shiftInfo.date, shiftInfo.shift);
  }

  // Panel 4: Unapproved Transactions (Mixing & Supply)
  payload.unapprovedTransactions = _getUnapprovedTransactions(ss, shiftInfo.date, shiftInfo.shift);

  return payload;
}

// =========================================================================
// 4. PANEL 1: MACHINE CHECKLIST STATUS
// =========================================================================

function _getChecklistStatus(ss, date, shift) {
  var dataSheet = ss.getSheetByName(CR_CONFIG.CHECKLIST.dataSheet);
  var masterSheet = ss.getSheetByName(CR_CONFIG.CHECKLIST.masterSheet);
  
  if (!dataSheet || !masterSheet || dataSheet.getLastRow() < 2) return [];

  // 1. Get Master Data (list of machines)
  var masterData = masterSheet.getRange("A2:A" + masterSheet.getLastRow()).getValues().flat().filter(String);
  var masterSet = new Set(masterData);

  // 2. Load DATA
  var dataValues = dataSheet.getDataRange().getValues();
  var header = dataValues.shift();
  
  var colIdx = _buildColumnIndex(header, [
    "jobId", "status", "createdAt", "data.Máy", "data.Ngày thực hiện", "data.Ca", "data.Người thực hiện", "data.Nhóm thực hiện", 
    "data.Kiểm_tra_Thông_tin_chung_Cụm_máy_có_kế_hoạch_hay_không", "data.Phê_duyệt__Duyệt"
  ]);

  // --- DEBUGGING LOG TO SHEET ---
  var logSheet = ss.getSheetByName("DEBUG_LOG");
  if (!logSheet) {
    logSheet = ss.insertSheet("DEBUG_LOG");
  }
  logSheet.clear();
  var logRows = [
    ["Target Date", date],
    ["Target Shift", shift],
    ["colIdx", JSON.stringify(colIdx)],
    ["Master Set Size", masterSet.size],
    ["Master Data", JSON.stringify(masterData)],
    [],
    ["Row #", "Machine", "jobId", "createdAt", "data.Ngày thực hiện", "data.Ca", "Parsed DateStr", "Parsed Shift", "Match Date?", "Match Shift?", "Match Both?", "Status", "isCompleted"]
  ];

  // 3. Process records
  var machineStatusMap = {};
  
  // Khởi tạo toàn bộ máy theo list Master bằng "Chưa hoàn thành"
  masterData.forEach(function(mach) {
    machineStatusMap[mach] = {
      machine: mach,
      group: "-",
      operator: "-",
      machineStatus: "-",
      status: "Chưa hoàn thành",
      approval: "-"
    };
  });

  dataValues.forEach(function(row, idx) {
    var mach = (row[colIdx["data.Máy"]] || "").toString().trim();
    if (!mach || !masterSet.has(mach)) return;

    var jobId = (row[colIdx["jobId"]] || "").toString().trim();
    var timestamp = _getOperationalTimestamp(row[colIdx["createdAt"]], jobId);
    var timestampShift = _getShiftInfoFromTimestamp(timestamp);
    
    var rawDateVal = row[colIdx["data.Ngày thực hiện"]];
    // Ưu tiên sử dụng ngày và ca làm việc tính từ thời điểm tạo (createdAt) bằng giờ VN (GMT+7)
    // để tránh lệch múi giờ của file Spreadsheet (America/Los_Angeles) làm cột data.Ngày thực hiện bị lùi 1 ngày vào ca sáng.
    var dateStr = (timestampShift ? timestampShift.date : "") || _safeDateStr(rawDateVal);
    var rowShift = (timestampShift ? timestampShift.shift : "") || _normalizeShiftLabel(row[colIdx["data.Ca"]]);
    
    var matchDate = (dateStr === date);
    var matchShift = (rowShift === shift);
    var matchBoth = matchDate && matchShift;
    var rawStatus = (row[colIdx["status"]] || "").toString().trim().toLowerCase();
    
    logRows.push([
      idx + 2, 
      mach, 
      jobId, 
      (row[colIdx["createdAt"]] ? row[colIdx["createdAt"]].toString() : "empty"), 
      (rawDateVal ? rawDateVal.toString() : "empty"), 
      (row[colIdx["data.Ca"]] ? row[colIdx["data.Ca"]].toString() : "empty"), 
      dateStr, 
      rowShift, 
      matchDate, 
      matchShift, 
      matchBoth, 
      rawStatus,
      rawStatus === "completed"
    ]);

    if (!dateStr || !rowShift) return;

    // Chỉ lấy bản ghi khớp với date và shift của Dashboard
    if (matchBoth) {
      var approvalStatus = (row[colIdx["data.Phê_duyệt__Duyệt"]] || "").toString().trim();
      var isCompleted = (rawStatus === "completed");
      var isInProgress = (rawStatus === "created" || rawStatus === "in-progress" || rawStatus === "paused");
      
      var displayStatus = "Chưa hoàn thành";
      if (isCompleted) displayStatus = "Hoàn thành";
      else if (isInProgress) displayStatus = "Đang thực hiện";

      var displayApproval = "-";
      if (isCompleted) {
        if (approvalStatus.toLowerCase() === "approve") displayApproval = "Đã duyệt";
        else displayApproval = "Chưa duyệt";
      }

      // Lưu lại giá trị mới nhất (ghi đè nếu có nhiều giao dịch)
      machineStatusMap[mach] = {
        machine: mach,
        group: (row[colIdx["data.Nhóm thực hiện"]] || "-").toString().trim(),
        operator: (row[colIdx["data.Người thực hiện"]] || "-").toString().trim(),
        machineStatus: (row[colIdx["data.Kiểm_tra_Thông_tin_chung_Cụm_máy_có_kế_hoạch_hay_không"]] || "-").toString().trim(),
        status: displayStatus,
        approval: displayApproval
      };
    }
  });

  // Ghi toàn bộ logs ra sheet DEBUG_LOG
  if (logRows.length > 0) {
    var maxCols = 0;
    logRows.forEach(function(r) {
      if (r.length > maxCols) maxCols = r.length;
    });
    logRows = logRows.map(function(r) {
      var arr = r.slice();
      while (arr.length < maxCols) {
        arr.push("");
      }
      return arr;
    });
    logSheet.getRange(1, 1, logRows.length, maxCols).setValues(logRows);
  }

  // Trả kết quả theo đúng thứ tự Master Data
  var results = [];
  masterData.forEach(function(mach) {
    results.push(machineStatusMap[mach]);
  });

  return results;
}

// =========================================================================
// 5. PANEL 2: CHANGEOVER COMPLIANCE
// =========================================================================

function _getChangeoverCompliance(ss, date, shift) {
  // --- Bước 1: Đọc DATA_MixingTransaction ---
  var txSheet = ss.getSheetByName(CR_CONFIG.CHANGEOVER.transactionSheet);
  if (!txSheet || txSheet.getLastRow() < 2) return { required: [], summary: { total: 0, done: 0, missing: 0, inProgress: 0 } };

  var txData = txSheet.getDataRange().getValues();
  var txHeader = txData.shift();

  // Tìm index các cột cần thiết
  var txColIdx = _buildColumnIndex(txHeader, [
    "MixingMachine", "ProductName", "Color", "CreatedDate",
    "DayShift", "Shift", "State", "Importer",
    "MaterialCode", "MaterialCode2", "ColorCode"
  ]);

  var shiftNumber = _shiftLabelToNumber(shift);
  var validTx = [];

  // Lọc bỏ giao dịch CANCELED, giữ lại tất cả để so sánh liên ca
  for (var i = 0; i < txData.length; i++) {
    var row = txData[i];
    var state = (row[txColIdx["State"]] || "").toString().trim().toUpperCase();
    if (state === "CANCELED") continue;

    validTx.push({
      machine: (row[txColIdx["MixingMachine"]] || "").toString().trim(),
      product: (row[txColIdx["ProductName"]] || "").toString().trim(),
      color: (row[txColIdx["Color"]] || "").toString().trim(),
      createdDate: row[txColIdx["CreatedDate"]],
      dayShift: _safeDateStr(row[txColIdx["DayShift"]]),
      shiftNum: (row[txColIdx["Shift"]] || "").toString().trim(),
      importer: (row[txColIdx["Importer"]] || "").toString().trim(),
      materialCode: (row[txColIdx["MaterialCode"]] || "").toString().trim(),
      materialCode2: (row[txColIdx["MaterialCode2"]] || "").toString().trim(),
      colorCode: (row[txColIdx["ColorCode"]] || "").toString().trim(),
    });
  }

  // --- Bước 2: Nhóm theo máy, sắp xếp theo CreatedDate ---
  var byMachine = {};
  validTx.forEach(function (tx) {
    if (!byMachine[tx.machine]) byMachine[tx.machine] = [];
    byMachine[tx.machine].push(tx);
  });

  // Helper check xem code nguyên vật liệu có thay đổi thật sự không
  function isCodeChanged(cVal, nVal) {
    var cv = (cVal || "").toString().trim().toUpperCase();
    var nv = (nVal || "").toString().trim().toUpperCase();
    
    if (cv === nv) return false;
    if (cv === "N/A" || nv === "N/A") return false;
    if (cv === "" || nv === "") return false;
    
    return true;
  }

  var requiredChangeovers = [];

  Object.keys(byMachine).forEach(function (machine) {
    var txs = byMachine[machine];
    // Sắp xếp theo CreatedDate tăng dần
    txs.sort(function (a, b) {
      var da = _parseFlexDate(a.createdDate);
      var db = _parseFlexDate(b.createdDate);
      return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
    });

    for (var j = 0; j < txs.length - 1; j++) {
      var curr = txs[j];
      var next = txs[j + 1];

      // Chỉ bắt những sự chuyển đổi xảy ra TẠI ca hiện tại
      // Nghĩa là: giao dịch gây ra thay đổi (next) thuộc về ca hiện tại
      if (next.dayShift === date && next.shiftNum === shiftNumber) {
        var matChanged = isCodeChanged(curr.materialCode, next.materialCode);
        var mat2Changed = isCodeChanged(curr.materialCode2, next.materialCode2);
        var colorCodeChanged = isCodeChanged(curr.colorCode, next.colorCode);

        // Chỉ trigger changeover khi có ít nhất 1 trong 3 mã thay đổi thực sự (và không liên quan N/A)
        if (matChanged || mat2Changed || colorCodeChanged) {
          requiredChangeovers.push({
            machine: machine,
            oldProduct: curr.product,
            newProduct: next.product,
            oldColor: curr.color,
            newColor: next.color,
            responsiblePerson: next.importer || "-",
            status: "❌ Thiếu", // Default, sẽ cập nhật ở bước tiếp
            changeoverJobId: null,
          });
        }
      }
    }
  });

  // --- Bước 3: Đọc DATA_Changeover-Mixing ---
  var coSheet = ss.getSheetByName(CR_CONFIG.CHANGEOVER.changeoverSheet);
  var changeoverRecords = [];

  if (coSheet && coSheet.getLastRow() >= 2) {
    var coData = coSheet.getDataRange().getValues();
    var coHeader = coData.shift();

    var coColIdx = _buildColumnIndex(coHeader, [
      "data.Máy", "data.Ngày thực hiện", "data.Ca",
      "data.Old_Bundle", "data.New_Bundle", "data.Old_Color", "data.New_Color",
      "status", "jobId"
    ]);

    for (var k = 0; k < coData.length; k++) {
      var coRow = coData[k];
      var coStatus = (coRow[coColIdx["status"]] || "").toString().trim().toLowerCase();
      if (coStatus === "cancelled") continue;

      var coDate = _safeDateStr(coRow[coColIdx["data.Ngày thực hiện"]]);
      var coCa = (coRow[coColIdx["data.Ca"]] || "").toString().trim();

      if (coDate === date && coCa === shiftNumber) {
        changeoverRecords.push({
          machine: (coRow[coColIdx["data.Máy"]] || "").toString().trim(),
          oldBundle: (coRow[coColIdx["data.Old_Bundle"]] || "").toString().trim(),
          newBundle: (coRow[coColIdx["data.New_Bundle"]] || "").toString().trim(),
          oldColor: (coRow[coColIdx["data.Old_Color"]] || "").toString().trim(),
          newColor: (coRow[coColIdx["data.New_Color"]] || "").toString().trim(),
          status: coStatus,
          jobId: (coRow[coColIdx["jobId"]] || "").toString().trim(),
        });
      }
    }
  }

  // --- Bước 4: Match changeover required với changeover thực tế ---
  requiredChangeovers.forEach(function (req) {
    var matched = changeoverRecords.find(function (co) {
      return co.machine === req.machine
        && co.oldBundle === req.oldProduct
        && co.newBundle === req.newProduct;
    });

    // Nếu không match theo Bundle, thử match theo Bundle + Color
    if (!matched) {
      matched = changeoverRecords.find(function (co) {
        return co.machine === req.machine
          && co.oldColor === req.oldColor
          && co.newColor === req.newColor
          && co.oldBundle === req.oldProduct;
      });
    }

    if (matched) {
      if (matched.status === "completed") {
        req.status = "✅ Đã thực hiện";
      } else if (matched.status === "paused") {
        req.status = "⏳ Chờ duyệt";
      } else if (matched.status === "in-progress" || matched.status === "created") {
        req.status = "🔄 Đang thực hiện";
      } else {
        req.status = "✅ Đã thực hiện";
      }
      req.changeoverJobId = matched.jobId;
    }
  });

  var summary = { total: requiredChangeovers.length, done: 0, missing: 0, inProgress: 0, pendingApproval: 0 };
  requiredChangeovers.forEach(function (r) {
    if (r.status.indexOf("✅") >= 0) summary.done++;
    else if (r.status.indexOf("🔄") >= 0) summary.inProgress++;
    else if (r.status.indexOf("⏳") >= 0) summary.pendingApproval++;
    else summary.missing++;
  });

  return { required: requiredChangeovers, summary: summary };
}

// =========================================================================
// 6. PANEL 3: ERROR ALERTS
// =========================================================================

function _getErrorAlerts(ss, date, shift) {
  var errSheet = ss.getSheetByName(CR_CONFIG.ERRORS.errorSheet);
  if (!errSheet || errSheet.getLastRow() < 2) return [];

  var errData = errSheet.getDataRange().getValues();
  var errHeader = errData.shift();

  // Header: Ngày, Ca, Máy, Sản phẩm, Batch, Màu Y/C, Vị trí lỗi, Mã Thực Tế,
  //         TS Thực, TS Chuẩn (Master), Độ lệch, BOM Chuẩn (Gợi ý),
  //         Ghi chú Lỗi, Người nhập, Nguyên nhân, Trạng thái khắc phục
  var colIdx = _buildColumnIndex(errHeader, [
    "Ngày", "Ca", "Máy", "Sản phẩm", "Batch", "Màu Y/C",
    "Vị trí lỗi", "Ghi chú Lỗi", "Người nhập", "Nguyên nhân", "Trạng thái khắc phục"
  ]);

  // Ca trong History_Error dùng số "1", "2", "3"
  var shiftNumber = _shiftLabelToNumber(shift);
  var results = [];

  for (var i = 0; i < errData.length; i++) {
    var row = errData[i];
    var rowDate = _safeDateStr(row[colIdx["Ngày"]]);
    var rowCa = (row[colIdx["Ca"]] || "").toString().trim();

    if (rowDate === date && rowCa === shiftNumber) {
      var remedyRaw = (row[colIdx["Trạng thái khắc phục"]] || "").toString().trim();
      var isRemedied = (remedyRaw.toLowerCase() === "đã khắc phục");

      results.push({
        machine: (row[colIdx["Máy"]] || "").toString().trim(),
        product: (row[colIdx["Sản phẩm"]] || "").toString().trim(),
        batch: (row[colIdx["Batch"]] || "").toString().trim(),
        color: (row[colIdx["Màu Y/C"]] || "").toString().trim(),
        errorLocation: (row[colIdx["Vị trí lỗi"]] || "").toString().trim(),
        errorNote: (row[colIdx["Ghi chú Lỗi"]] || "").toString().trim(),
        operator: (row[colIdx["Người nhập"]] || "").toString().trim(),
        rootCause: (row[colIdx["Nguyên nhân"]] || "").toString().trim(),
        remedyStatus: isRemedied ? "Đã khắc phục" : "Chưa khắc phục",
      });
    }
  }

  return results;
}

// =========================================================================
// 6B. PANEL 4: UNAPPROVED TRANSACTIONS
// =========================================================================

function _getUnapprovedTransactions(ss, date, shift) {
  var unapproved = [];

  // 1. DATA_MixingTransaction
  var mixingSheet = ss.getSheetByName("DATA_MixingTransaction");
  if (mixingSheet && mixingSheet.getLastRow() >= 2) {
    var mixingData = mixingSheet.getDataRange().getValues();
    var mixingHeader = mixingData.shift();
    var mixingIdx = _buildColumnIndex(mixingHeader, [
      "MixingMachine", "ProductName", "MixingBatch", "State", "Importer", "DayShift", "Shift"
    ]);

    var shiftNum = _shiftLabelToNumber(shift);

    mixingData.forEach(function (row) {
      var state = (row[mixingIdx["State"]] || "").toString().trim();
      var stateUpper = state.toUpperCase();
      // Bỏ qua APPROVED, CANCELED và trống
      if (stateUpper === "APPROVED" || stateUpper === "CANCELED" || stateUpper === "") return;

      var dayShift = _safeDateStr(row[mixingIdx["DayShift"]]);
      var rowShift = (row[mixingIdx["Shift"]] || "").toString().trim();

      if (dayShift === date && rowShift === shiftNum) {
        unapproved.push({
          type: "Mixing",
          machine: (row[mixingIdx["MixingMachine"]] || "").toString().trim(),
          product: (row[mixingIdx["ProductName"]] || "").toString().trim(),
          batch: (row[mixingIdx["MixingBatch"]] || "").toString().trim(),
          state: state,
          operator: (row[mixingIdx["Importer"]] || "").toString().trim()
        });
      }
    });
  }

  return unapproved;
}

// =========================================================================
// 7. UTILITY FUNCTIONS
// =========================================================================

/**
 * Xác định Ca hiện tại. Logic giống getCurrentShift() trong Update KPI.
 */
function _getCurrentShiftCR() {
  var tz = CR_CONFIG.TIME_ZONE;
  var now = new Date();
  var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");
  var dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  var shift = "Ca 3";

  if (timeStr >= "05:55:00" && timeStr <= "13:54:59") {
    shift = "Ca 1";
  } else if (timeStr >= "13:55:00" && timeStr <= "21:54:59") {
    shift = "Ca 2";
  } else {
    shift = "Ca 3";
    if (timeStr <= "05:54:59") {
      var prevDate = new Date(now.getTime() - 24 * 3600000);
      dateStr = Utilities.formatDate(prevDate, tz, "yyyy-MM-dd");
    }
  }

  return { date: dateStr, shift: shift };
}

/**
 * Chuyển "Ca 1" → "1", "Ca 2" → "2", "Ca 3" → "3"
 */
function _shiftLabelToNumber(label) {
  if (label === "Ca 1" || label === "1") return "1";
  if (label === "Ca 2" || label === "2") return "2";
  if (label === "Ca 3" || label === "3") return "3";
  return label.replace(/\D/g, "") || label;
}

/**
 * Ép mọi định dạng ngày về "yyyy-MM-dd".
 */
function _safeDateStr(rawDate) {
  if (!rawDate) return "";
  var tz = CR_CONFIG.TIME_ZONE;

  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    try {
      tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch(e) {}
    return Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
  }

  var text = rawDate.toString().trim();
  if (!text) return "";

  // Đã là yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  // yyyy-MM-dd HH:mm:ss...
  var isoMatch = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoMatch) return isoMatch[1];

  // DD/MM/YYYY
  var dmyMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmyMatch) {
    var d = dmyMatch[1].padStart(2, "0");
    var m = dmyMatch[2].padStart(2, "0");
    return dmyMatch[3] + "-" + m + "-" + d;
  }

  // Thử parse bằng new Date()
  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    try {
      tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch(e) {}
    return Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
  }

  return text;
}

function _normalizeShiftLabel(value) {
  if (value === null || value === undefined || value === "") return "";
  var text = value.toString().trim();
  if (text === "1" || text === "Ca 1") return "Ca 1";
  if (text === "2" || text === "Ca 2") return "Ca 2";
  if (text === "3" || text === "Ca 3") return "Ca 3";
  return text;
}

function _getOperationalTimestamp(value, jobId) {
  var objectIdDate = _dateFromObjectId(jobId);
  if (objectIdDate) return objectIdDate;

  var parsed = _parseFlexDate(value);
  return parsed;
}

function _dateFromObjectId(jobId) {
  if (!jobId) return null;
  var text = jobId.toString().trim();
  if (!/^[0-9a-fA-F]{8}/.test(text)) return null;
  return new Date(parseInt(text.slice(0, 8), 16) * 1000);
}

function _getShiftInfoFromTimestamp(timestamp) {
  if (!timestamp || isNaN(timestamp.getTime())) return null;
  var tz = CR_CONFIG.TIME_ZONE;
  var timeStr = Utilities.formatDate(timestamp, tz, "HH:mm:ss");
  var dateStr = Utilities.formatDate(timestamp, tz, "yyyy-MM-dd");
  var shift = "Ca 3";

  if (timeStr >= "05:55:00" && timeStr <= "13:54:59") {
    shift = "Ca 1";
  } else if (timeStr >= "13:55:00" && timeStr <= "21:54:59") {
    shift = "Ca 2";
  } else if (timeStr <= "05:54:59") {
    var prevDate = new Date(timestamp.getTime() - 24 * 3600000);
    dateStr = Utilities.formatDate(prevDate, tz, "yyyy-MM-dd");
  }

  return { date: dateStr, shift: shift };
}

/**
 * Parse ngày linh hoạt (wrapper nhẹ hơn _parseFlexibleDate trong Code.gs).
 */
function _parseFlexDate(value) {
  if (!value) return null;
  
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  
  var text = value.toString().trim();

  if (!text) return null;

  // Hỗ trợ định dạng Domo: yyyy-MM-dd HH:mm:ss.nnnnnnn +07:00
  var domoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?(?:\s*([+-]\d{2}):?(\d{2}))?$/);
  if (domoMatch) {
    var year = Number(domoMatch[1]);
    var month = Number(domoMatch[2]) - 1;
    var day = Number(domoMatch[3]);
    var hour = Number(domoMatch[4] || 0);
    var minute = Number(domoMatch[5] || 0);
    var second = Number(domoMatch[6] || 0);
    var offsetHourText = domoMatch[7];
    var offsetMinute = Number(domoMatch[8] || 0);

    if (offsetHourText) {
      var offsetHour = Number(offsetHourText);
      var utcMillis = Date.UTC(year, month, day, hour, minute, second) - ((offsetHour * 60) + Math.sign(offsetHour) * offsetMinute) * 60000;
      return new Date(utcMillis);
    }

    return new Date(year, month, day, hour, minute, second);
  }

  // yyyy-MM-dd không có anchor chặt chẽ (để match phần đầu)
  var isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (isoMatch) {
    return new Date(
      Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]),
      Number(isoMatch[4] || 0), Number(isoMatch[5] || 0), Number(isoMatch[6] || 0)
    );
  }

  // DD/MM/YYYY
  var dmyMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmyMatch) {
    return new Date(Number(dmyMatch[3]), Number(dmyMatch[2]) - 1, Number(dmyMatch[1]));
  }

  var nativeDate = new Date(text);
  return isNaN(nativeDate.getTime()) ? null : nativeDate;
}

/**
 * Tạo map { columnName → columnIndex } từ header row.
 */
function _buildColumnIndex(headerRow, columns) {
  var idx = {};
  columns.forEach(function (col) {
    for (var i = 0; i < headerRow.length; i++) {
      if ((headerRow[i] || "").toString().trim() === col) {
        idx[col] = i;
        break;
      }
    }
    if (idx[col] === undefined) idx[col] = -1;
  });
  return idx;
}

// =========================================================================
// 8. TRIGGER — Sync data + ghi Snapshot
// =========================================================================

/**
 * Trigger function — cài chạy mỗi giờ hoặc mỗi 30 phút.
 * 1. Gọi update() để sync dữ liệu mới nhất từ Augmentir/Domo về Sheet.
 * 2. Xóa cache Dashboard để lần mở kế tiếp luôn có data mới.
 * 3. Ghi snapshot vào CR_History_Snapshot.
 */
function triggerControlRoom() {
  // Bước 1: Sync dữ liệu
  try {
    if (typeof update === "function") {
      update();
      SpreadsheetApp.flush();
    }
  } catch (e) {
    Logger.log("triggerControlRoom: update() failed — " + e.toString());
  }

  // Bước 2: Clear cache
  try {
    CacheService.getScriptCache().removeAll([
      "CR_" + CR_CONFIG.AREA_NAME + "_" + _getCurrentShiftCR().date + "_Ca 1",
      "CR_" + CR_CONFIG.AREA_NAME + "_" + _getCurrentShiftCR().date + "_Ca 2",
      "CR_" + CR_CONFIG.AREA_NAME + "_" + _getCurrentShiftCR().date + "_Ca 3",
    ]);
  } catch (e) {
    Logger.log("triggerControlRoom: cache clear failed — " + e.toString());
  }

  // Bước 3: Ghi snapshot
  if (CR_CONFIG.SNAPSHOT.enabled) {
    try {
      _writeSnapshot();
    } catch (e) {
      Logger.log("triggerControlRoom: snapshot failed — " + e.toString());
    }
  }
}

// =========================================================================
// 9. SNAPSHOT — Ghi lịch sử ra sheet cho Domo
// =========================================================================

var _SNAPSHOT_HEADERS = [
  "timestamp", "shift_date", "shift", "panel", "machine", "status",
  "details", "responsible_person"
];

/**
 * Ghi 1 bản snapshot của ca hiện tại vào CR_History_Snapshot.
 * Mỗi dòng = 1 record (checklist / changeover / error).
 * Domo chỉ cần trỏ connector vào sheet này.
 */
function _writeSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = CR_CONFIG.SNAPSHOT.sheetName;
  var sheet = ss.getSheetByName(sheetName);

  // Tạo sheet nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(_SNAPSHOT_HEADERS);
    sheet.getRange(1, 1, 1, _SNAPSHOT_HEADERS.length)
      .setBackground("#1a237e").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var shiftInfo = _getCurrentShiftCR();
  var tz = CR_CONFIG.TIME_ZONE;
  var now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  var targetShiftDWH = _translateShift(shiftInfo.shift);

  // Đọc dữ liệu lịch sử hiện có
  var lastRow = sheet.getLastRow();
  var existingData = [];
  var headers = [];
  if (lastRow > 0) {
    var allValues = sheet.getRange(1, 1, lastRow, _SNAPSHOT_HEADERS.length).getValues();
    headers = allValues[0];
    existingData = allValues.slice(1);
  } else {
    headers = _SNAPSHOT_HEADERS;
  }

  // Lọc bỏ các dòng cũ của Ca hiện tại (trùng Ngày và Ca) để ghi đè phiên bản mới nhất
  var cleanRows = existingData.filter(function (row) {
    if (row.length < 3) return true;
    
    var rowDateStr = "";
    if (row[1] instanceof Date) {
      rowDateStr = Utilities.formatDate(row[1], tz, "yyyy-MM-dd");
    } else {
      rowDateStr = row[1] ? row[1].toString().trim() : "";
    }
    
    var rowShift = row[2] ? row[2].toString().trim() : "";
    var isCurrentShift = (rowDateStr === shiftInfo.date && rowShift === targetShiftDWH);
    return !isCurrentShift; // Giữ lại các dòng KHÔNG phải ca hiện tại
  });

  // Tạo dữ liệu snapshot mới cho Ca hiện tại
  var payload = _buildDashboardPayload(shiftInfo);
  var newRows = [];

  // --- Checklist records ---
  if (payload.checklist) {
    payload.checklist.forEach(function (item) {
      newRows.push([
        now,
        shiftInfo.date,
        _translateShift(shiftInfo.shift),
        "Checklist",
        item.machine,
        _translateChecklistStatus(item.status),
        _translateChecklistDetails(item.status === "Hoàn thành" ? ("Duyệt: " + item.approval) : "-"),
        item.operator || "-"
      ]);
    });
  }

  // --- Changeover records ---
  if (payload.changeover && payload.changeover.required) {
    payload.changeover.required.forEach(function (r) {
      newRows.push([
        now,
        shiftInfo.date,
        _translateShift(shiftInfo.shift),
        "Changeover",
        r.machine,
        _translateChangeoverStatus(r.status),
        _translateChangeoverDetails(r.oldProduct + " -> " + r.newProduct + " | " + r.oldColor + " -> " + r.newColor),
        r.responsiblePerson || "-"
      ]);
    });
  }

  // --- Error records ---
  if (payload.errors) {
    payload.errors.forEach(function (e) {
      newRows.push([
        now,
        shiftInfo.date,
        _translateShift(shiftInfo.shift),
        "Error",
        e.machine,
        _translateErrorStatus(e.remedyStatus),
        e.errorLocation + " - " + e.errorNote + " | " + e.product + " / " + e.batch,
        e.operator || "-"
      ]);
    });
  }

  // Ghép dòng cũ đã lọc sạch + dòng mới của ca hiện hành
  var finalValues = [headers].concat(cleanRows).concat(newRows);

  // Clear nội dung cũ và ghi toàn bộ dữ liệu mới xuống sheet
  sheet.clearContents();
  sheet.getRange(1, 1, finalValues.length, _SNAPSHOT_HEADERS.length).setValues(finalValues);

  // Áp dụng lại định dạng tiêu đề
  sheet.getRange(1, 1, 1, _SNAPSHOT_HEADERS.length)
    .setBackground("#1a237e").setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}

// =========================================================================
// 9B. SNAPSHOT TRANSLATION UTILITIES (DWH STANDARD)
// =========================================================================

function _translateShift(shiftVal) {
  var s = (shiftVal || "").toString().trim();
  if (s === "Ca 1" || s === "1" || s === "Shift 1") return "Shift 1";
  if (s === "Ca 2" || s === "2" || s === "Shift 2") return "Shift 2";
  if (s === "Ca 3" || s === "3" || s === "Shift 3") return "Shift 3";
  return s;
}

function _translateChecklistStatus(statusVal) {
  var s = (statusVal || "").toString().trim();
  if (s === "Hoàn thành" || s === "Completed") return "Completed";
  if (s === "Đang thực hiện" || s === "In Progress") return "In Progress";
  if (s === "Chưa hoàn thành" || s === "Not Started") return "Not Started";
  return s;
}

function _translateChecklistDetails(detailVal) {
  var d = (detailVal || "").toString().trim();
  if (d === "Duyệt: Đã duyệt" || d === "Duyệt: Approved" || d === "Approval: Approved") return "Approval: Approved";
  if (d === "Duyệt: Chưa duyệt" || d === "Duyệt: Pending" || d === "Approval: Pending") return "Approval: Pending";
  return d;
}

function _translateChangeoverStatus(statusVal) {
  var s = (statusVal || "").toString().trim().replace(/[\u2700-\u27BF]|[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD00-\uDFFF]/g, '').trim();
  if (s === "Thiếu" || s === "Missing") return "Missing";
  if (s === "Đã thực hiện" || s === "Completed") return "Completed";
  if (s === "Chờ duyệt" || s === "Pending Approval") return "Pending Approval";
  if (s === "Đang thực hiện" || s === "In Progress") return "In Progress";
  return s;
}

function _translateChangeoverDetails(detailVal) {
  return (detailVal || "").toString().replace(/→/g, "->").trim();
}

function _translateErrorStatus(statusVal) {
  var s = (statusVal || "").toString().trim();
  if (s === "Đã khắc phục" || s === "Resolved") return "Resolved";
  if (s === "Chưa khắc phục" || s === "Unresolved") return "Unresolved";
  return s;
}

/**
 * One-time migration function to translate all historical snapshot records to English
 * and apply Data Warehouse naming conventions. Run this from the Apps Script Editor once.
 */
function migrateSnapshotSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = CR_CONFIG.SNAPSHOT.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No historical data to migrate.");
    return;
  }

  var range = sheet.getRange(1, 1, lastRow, _SNAPSHOT_HEADERS.length);
  var values = range.getValues();

  // 1. Update headers
  values[0] = _SNAPSHOT_HEADERS;

  // 2. Translate all rows
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    
    // Format timestamp
    if (row[0] instanceof Date) {
      row[0] = Utilities.formatDate(row[0], CR_CONFIG.TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
    } else if (row[0]) {
      row[0] = row[0].toString().trim();
    }
    
    // Format date
    if (row[1] instanceof Date) {
      row[1] = Utilities.formatDate(row[1], CR_CONFIG.TIME_ZONE, "yyyy-MM-dd");
    } else if (row[1]) {
      row[1] = row[1].toString().trim();
    }

    // Translate shift
    row[2] = _translateShift(row[2]);

    // Translate panel
    var panel = (row[3] || "").toString().trim();
    if (panel === "Checklist") {
      row[5] = _translateChecklistStatus(row[5]);
      row[6] = _translateChecklistDetails(row[6]);
    } else if (panel === "Changeover") {
      row[5] = _translateChangeoverStatus(row[5]);
      row[6] = _translateChangeoverDetails(row[6]);
    } else if (panel === "Error") {
      row[5] = _translateErrorStatus(row[5]);
      row[6] = (row[6] || "").toString().replace(/—/g, "-").replace(/→/g, "->").trim();
    }
  }

  // 3. Clear and write back
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, _SNAPSHOT_HEADERS.length).setValues(values);
  
  // Set formatting
  sheet.getRange(1, 1, 1, _SNAPSHOT_HEADERS.length)
    .setBackground("#1a237e").setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);

  Logger.log("Migration to English DWH standard completed successfully for " + (values.length - 1) + " rows.");
}

// =========================================================================
// 10. TEST FUNCTIONS
// =========================================================================

/**
 * Test getDashboardData() cho ca hiện tại. Chạy từ Editor để debug.
 */
function testDashboard() {
  var result = getDashboardData();
  Logger.log(result);
}

/**
 * Test với ngày + ca cụ thể.
 */
function testDashboardSpecific() {
  var result = getDashboardData({ date: "2026-05-14", shift: "Ca 3" });
  Logger.log(result);
}

/**
 * Xóa cache và chạy chẩn đoán để ghi logs vào sheet DEBUG_LOG.
 */
function runDiagnostic() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftInfo = _getCurrentShiftCR();
  
  // Xóa cache của ca hiện tại để đảm bảo nạp dữ liệu mới
  var cacheKey = "CR_" + CR_CONFIG.AREA_NAME + "_" + shiftInfo.date + "_" + shiftInfo.shift;
  CacheService.getScriptCache().remove(cacheKey);
  
  // Gọi payload builder để tự động kích hoạt tạo DEBUG_LOG
  _buildDashboardPayload(shiftInfo);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Chẩn đoán hoàn tất! Hãy kiểm tra sheet DEBUG_LOG.", "Diagnostic", 5);
}
