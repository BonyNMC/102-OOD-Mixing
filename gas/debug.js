function saveDebugFile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("DATA");
  if (!dataSheet) {
    saveTextFile("Error: Sheet 'DATA' not found");
    return;
  }
  
  var dataValues = dataSheet.getDataRange().getValues();
  var header = dataValues.shift();
  
  var colIdx = _buildColumnIndex(header, [
    "jobId", "status", "createdAt", "data.Máy", "data.Ngày thực hiện", "data.Ca"
  ]);
  
  var output = [];
  output.push("Total rows in DATA: " + dataValues.length);
  output.push("Spreadsheet Timezone: " + ss.getSpreadsheetTimeZone());
  output.push("Script Timezone: " + Session.getScriptTimeZone());
  output.push("CR_CONFIG.TIME_ZONE: " + CR_CONFIG.TIME_ZONE);
  
  var foundCount = 0;
  dataValues.forEach(function(row, idx) {
    var mach = (row[colIdx["data.Máy"]] || "").toString().trim();
    var createdAtRaw = row[colIdx["createdAt"]];
    var rawDateVal = row[colIdx["data.Ngày thực hiện"]];
    var jobId = (row[colIdx["jobId"]] || "").toString().trim();
    
    var dateStr = _safeDateStr(rawDateVal);
    var timestamp = _getOperationalTimestamp(createdAtRaw, jobId);
    var timestampShift = _getShiftInfoFromTimestamp(timestamp);
    var calcDate = timestampShift ? timestampShift.date : "";
    var calcShift = timestampShift ? timestampShift.shift : "";
    
    var isToday = (dateStr === "2026-05-22" || calcDate === "2026-05-22");
    if (isToday) {
      foundCount++;
      output.push("Row " + (idx + 2) + " -> Machine: " + mach + 
                  " | rawDateVal: " + (rawDateVal ? rawDateVal.toString() : "empty") + " (parsed: " + dateStr + ")" +
                  " | createdAt: " + (createdAtRaw ? createdAtRaw.toString() : "empty") + " (calcDate: " + calcDate + ", calcShift: " + calcShift + ")" +
                  " | status: " + row[colIdx["status"]]);
    }
  });
  
  output.push("Found today's records count: " + foundCount);
  saveTextFile(output.join("\n"));
}

function saveTextFile(text) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Write to Sheet "DEBUG_FILE_OUTPUT"
  try {
    var sheetName = "DEBUG_FILE_OUTPUT";
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    sheet.clear();
    
    var lines = text.split("\n").map(function(line) {
      return [line];
    });
    
    sheet.getRange(1, 1, lines.length, 1).setValues(lines);
    Logger.log("Successfully wrote debug log to sheet: " + sheetName);
  } catch(sheetErr) {
    Logger.log("Failed to write to sheet: " + sheetErr.message);
  }

  // 2. Also try writing to Google Drive folder "AI Sheet Snapshots" or parent folder
  try {
    var folders = DriveApp.getFoldersByName("AI Sheet Snapshots");
    var folder = null;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      // Fallback: get the parent folder of the spreadsheet itself
      var parentFolders = DriveApp.getFileById(ss.getId()).getParents();
      if (parentFolders.hasNext()) {
        folder = parentFolders.next();
      }
    }
    
    if (folder) {
      var files = folder.getFilesByName("debug_output.txt");
      while (files.hasNext()) {
        files.next().setTrashed(true);
      }
      folder.createFile("debug_output.txt", text);
      Logger.log("Successfully created debug_output.txt in " + folder.getName());
    } else {
      Logger.log("Could not resolve destination folder on Google Drive.");
    }
  } catch(driveErr) {
    Logger.log("Could not write file to Drive: " + driveErr.message);
  }
}

function runDiagnosticErrors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errSheet = ss.getSheetByName("History_Error_Mixing_none_Match_Data");
  if (!errSheet) {
    saveTextFile("Error: Sheet 'History_Error_Mixing_none_Match_Data' not found");
    return;
  }

  var data = errSheet.getDataRange().getValues();
  var header = data.shift();
  var output = [];
  output.push("Total rows in History_Error_Mixing_none_Match_Data: " + (data.length + 1));
  output.push("Spreadsheet Timezone: " + ss.getSpreadsheetTimeZone());
  output.push("Current Time: " + Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss"));
  
  // Find column indices
  var colIdx = _buildColumnIndex(header, [
    "Ngày", "Ca", "Máy", "Sản phẩm", "Batch", "Màu Y/C",
    "Vị trí lỗi", "Ghi chú Lỗi", "Người nhập", "Nguyên nhân", "Trạng thái khắc phục"
  ]);

  if (colIdx["Ngày"] === -1) {
    colIdx["Ngày"] = 0;
  }

  var lastRows = data.slice(-10); // get last 10 rows
  output.push("\n--- LAST 10 ROWS IN SHEET ---");
  
  var targetDate = "2026-06-02";
  var targetShiftNum = "1"; // Ca 1

  lastRows.forEach(function(row, idx) {
    var rawDate = row[colIdx["Ngày"]];
    var rawCa = row[colIdx["Ca"]];
    
    var parsedDate = _safeDateStr(rawDate);
    var parsedShiftNum = _shiftLabelToNumber(rawCa ? rawCa.toString().trim() : "");
    
    var matchDate = (parsedDate === targetDate);
    var matchShift = (parsedShiftNum === targetShiftNum);
    var status = (row[colIdx["Trạng thái khắc phục"]] || "").toString().trim();
    
    output.push("Row " + (data.length - 10 + idx + 2) + ":");
    output.push("  Raw: Ngày=" + (rawDate ? rawDate.toString() : "null") + " (Type: " + (typeof rawDate) + "), Ca=" + rawCa);
    output.push("  Parsed: Ngày=" + parsedDate + " (Match target? " + matchDate + "), Ca=" + parsedShiftNum + " (Match target? " + matchShift + ")");
    output.push("  Machine=" + row[colIdx["Máy"]] + ", Batch=" + row[colIdx["Batch"]] + ", Status=" + status);
  });

  saveTextFile(output.join("\n"));
}

function runSyncDiagnostics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var output = [];
  output.push("=== START SYNC DIAGNOSTICS ===");
  output.push("Current Time: " + Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss"));
  
  // Log SETUP_NEW state
  try {
    var setupNew = ss.getSheetByName("SETUP_NEW");
    if (setupNew) {
      var data = setupNew.getDataRange().getValues();
      output.push("SETUP_NEW row count: " + data.length);
      data.forEach(function(row, idx) {
        if (idx === 0 || idx === 1) return; // skip headers
        if (row[0] === true || row[0] === "TRUE") {
          output.push("Procedure Enabled: Platform=" + row[1] + " | Sheet=" + row[2] + " | Source=" + row[3] + " | ID=" + row[4] + " | LastUpdated=" + row[20]);
        }
      });
    } else {
      output.push("SETUP_NEW sheet not found!");
    }
  } catch(e) {
    output.push("Failed to read SETUP_NEW: " + e.toString());
  }

  // Attempt to run update()
  try {
    output.push("\nRunning update()...");
    if (typeof update === "function") {
      update();
      output.push("update() completed successfully!");
    } else {
      output.push("Error: update() function is not defined!");
    }
  } catch (e) {
    output.push("Error during update(): " + e.toString());
    if (e.stack) {
      output.push("Stack trace:\n" + e.stack);
    }
  }

  output.push("=== END SYNC DIAGNOSTICS ===");
  saveTextFile(output.join("\n"));
}
