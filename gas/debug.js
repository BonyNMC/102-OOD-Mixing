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
