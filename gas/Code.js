/** LodashGS library: 1SQ0PlSMwndIuOAgtVJdjxsuXueECtY9OGejVDS37ckSVbMll73EXf2PW */

const APP_KEY = '69091b790253df048876a1e8';
const SERVER_ENDPOINT = 'https://app.augmentir.com';
const SETUP_SHEET_NAME = 'SETUP';
const SETUP_NEW_SHEET_NAME = 'SETUP_NEW';
const DATA_SHEET_NAME = 'DATA';
const DOMO_CREDENTIAL_SPREADSHEET_ID = '1LwX3FWLJOlssznZKAH2Xu6C3j3WAq1yhMU-xOb9k_Rs';
const DOMO_CREDENTIAL_SHEET_NAME = 'Domo';
const LOCAL_TIME_ZONE = 'Asia/Ho_Chi_Minh';
const LIMIT = 1000;
const DEFAULT_HEADERS = [
  "jobId",
  "jobEventUrl",
  "procedureId",
  "title",
  "publishState",
  "procedureVersion",
  "status",

  "createdAt",
  "startedAt",
  "updatedAt",
  "completedAt",

  "stationId",
  "isExpired",
  "minutesRemaining",
  "approxMinutesElapsed",

  "cardInTask",
  "cardsInTask",

  "unitNumber",
  "unitCount",
];
const _ = LodashGS.load();
const CONFIG_RANGE = {
  PROCEDURE: {
    NAME: 'C2',
    ID: 'C3',
  },
  PROCEDURE_TABLE: {
    START_ROW: 3,
    START_COLUMN: 1,
    COLUMN_COUNT: 21,
  },
  STATUS: {
    CREATED: 'C4',
    INPROGRESS: 'C5',
    PAUSED: 'C6',
    CANCELLED: 'C7',
    COMPLETED: 'C8',
  },
  PUBLISH_STATE: {
    PRODUCTION: 'C9',
    PREPRODUCTION: 'C10',
    DRAFT: 'C11',
  },
  START_DATE: 'C12',
  END_DATE: 'C13',
  LOCK_HEADERS: 'C14',
  INCLUDE_EXPIRED_JOBS: 'C15',
};

const ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Augmentir')
    .addItem('Update', 'update')
    .addItem('Create SETUP_NEW', 'setupMultiProcedureTable')
    .addItem('Create README', 'createReadmeSheet')
    .addItem('Run Diagnostics', 'runDiagnostic')
    .addItem('Archive Old Data (30 Days)', 'runManualArchive')
    .addSeparator()
    .addSubMenu(ui.createMenu('Clear')
      .addItem('Clear Updated From', 'clearUpdatedFrom')
      .addItem('Clear Data', 'clearData')
      .addItem('Clear Headers', 'clearHeaders'))
    .addToUi();
}

function clearUpdatedFrom() {
  const newConfigSheet = ss.getSheetByName(SETUP_NEW_SHEET_NAME);
  if (newConfigSheet) {
    const table = CONFIG_RANGE.PROCEDURE_TABLE;
    const lastRow = newConfigSheet.getLastRow();
    if (lastRow >= table.START_ROW) {
      const fromColumn = newConfigSheet.getRange(2, 2).getValue() === 'Platform' ? 17 : 13;
      newConfigSheet.getRange(table.START_ROW, fromColumn, lastRow - table.START_ROW + 1, 1).clearContent();
      return;
    }
  }

  const configSheet = ss.getSheetByName(SETUP_SHEET_NAME);
  configSheet.getRange(CONFIG_RANGE.START_DATE).activate();
  configSheet.getRange(CONFIG_RANGE.START_DATE).clearContent();
}

function clearData() {
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  dataSheet.activate();
  if (dataSheet.getLastRow() < 2) return;
  dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).clear();
}

function clearHeaders() {
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  dataSheet.activate();
  dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).clear();
}

function update() {
  const configSheet = ss.getSheetByName(SETUP_NEW_SHEET_NAME) || ss.getSheetByName(SETUP_SHEET_NAME);
  if (!configSheet) throw new Error(`Không tìm thấy Sheet '${SETUP_NEW_SHEET_NAME}' hoặc '${SETUP_SHEET_NAME}'`);

  const procedureConfigs = _getProcedureConfigs(configSheet);
  if (procedureConfigs.length === 0) {
    throw new Error(`Chưa cấu hình Procedure nào trong Sheet '${configSheet.getName()}'`);
  }

  for (const procedureConfig of procedureConfigs) {
    try {
      if (procedureConfig.platform === 'DOMO') {
        _updateDomoDataset(procedureConfig);
      } else if (procedureConfig.platform === 'GOOGLE SHEET') {
        _copyGoogleSheetData(procedureConfig);
      } else {
        _updateProcedure(procedureConfig);
      }
    } catch (err) {
      const sourceName = procedureConfig.dataSheetName || procedureConfig.procedureId || procedureConfig.procedureName || procedureConfig.platform;
      throw new Error(`Update source '${sourceName}' failed: ${err.message || err}`);
    }
  }
}

function _getProcedureConfigs(configSheet) {
  const tableConfigs = configSheet.getName() === SETUP_NEW_SHEET_NAME ? _getProcedureConfigsFromTable(configSheet) : [];
  if (tableConfigs.length > 0) return tableConfigs;

  const procedureName = _cleanConfigText(configSheet.getRange(CONFIG_RANGE.PROCEDURE.NAME).getValue());
  const procedureId = _cleanConfigText(configSheet.getRange(CONFIG_RANGE.PROCEDURE.ID).getValue());

  if (!procedureName && !procedureId) return [];
  if (procedureName && procedureId) {
    throw new Error(`Không thể chỉ định đồng thời 'Procedure Name' và 'Procedure ID'`);
  }

  return [{
    procedureName,
    procedureId,
    platform: 'AUGMENTIR',
    dataSheetName: DATA_SHEET_NAME,
    status: {
      created: configSheet.getRange(CONFIG_RANGE.STATUS.CREATED).getValue(),
      inProgress: configSheet.getRange(CONFIG_RANGE.STATUS.INPROGRESS).getValue(),
      paused: configSheet.getRange(CONFIG_RANGE.STATUS.PAUSED).getValue(),
      cancelled: configSheet.getRange(CONFIG_RANGE.STATUS.CANCELLED).getValue(),
      completed: configSheet.getRange(CONFIG_RANGE.STATUS.COMPLETED).getValue(),
    },
    publishState: {
      production: configSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PRODUCTION).getValue(),
      preProduction: configSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PREPRODUCTION).getValue(),
      draft: configSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.DRAFT).getValue(),
    },
    startDate: configSheet.getRange(CONFIG_RANGE.START_DATE).getValue(),
    endDate: configSheet.getRange(CONFIG_RANGE.END_DATE).getValue(),
    lockHeaders: configSheet.getRange(CONFIG_RANGE.LOCK_HEADERS).getValue(),
    includeExpiredJobs: configSheet.getRange(CONFIG_RANGE.INCLUDE_EXPIRED_JOBS).getValue(),
    configSheet,
    updatedFromCell: CONFIG_RANGE.START_DATE,
    lastUpdatedCell: CONFIG_RANGE.START_DATE,
  }];
}

function _cleanConfigText(value) {
  if (value === null || value === undefined) return '';
  return value.toString().trim();
}

function _getProcedureConfigsFromTable(configSheet) {
  const table = CONFIG_RANGE.PROCEDURE_TABLE;
  const lastRow = configSheet.getLastRow();
  if (lastRow < table.START_ROW) return [];

  const isNewLayout = configSheet.getRange(2, 2).getValue() === 'Platform';
  const values = configSheet
    .getRange(table.START_ROW, table.START_COLUMN, lastRow - table.START_ROW + 1, table.COLUMN_COUNT)
    .getValues();

  return values
    .map((row, index) => {
      const rowNumber = table.START_ROW + index;
      return isNewLayout
        ? _parseSetupNewRow(configSheet, row, rowNumber)
        : _parseLegacySetupNewRow(configSheet, row, rowNumber);
    })
    .filter(Boolean);
}

function _parseLegacySetupNewRow(configSheet, row, rowNumber) {
  const enabled = row[0];
  const dataSheetName = _cleanConfigText(row[1]) || DATA_SHEET_NAME;
  const procedureName = _cleanConfigText(row[2]);
  const procedureId = _cleanConfigText(row[3]);
  const updatedFrom = row[12];
  const updatedTo = row[13];

  if (!_isEnabled(enabled)) return null;
  if (!procedureName && !procedureId) return null;
  if (procedureName && procedureId) {
    throw new Error(`Dòng ${rowNumber}: chỉ nhập một trong hai cột Procedure Name hoặc Procedure ID`);
  }

  return {
    platform: 'AUGMENTIR',
    procedureName,
    procedureId,
    dataSheetName,
    status: {
      created: row[4],
      inProgress: row[5],
      paused: row[6],
      cancelled: row[7],
      completed: row[8],
    },
    publishState: {
      production: row[9],
      preProduction: row[10],
      draft: row[11],
    },
    startDate: updatedFrom,
    endDate: updatedTo,
    lockHeaders: row[14],
    includeExpiredJobs: row[15],
    configSheet,
    updatedFromCell: configSheet.getRange(rowNumber, 13).getA1Notation(),
    lastUpdatedCell: configSheet.getRange(rowNumber, 17).getA1Notation(),
  };
}

function _parseSetupNewRow(configSheet, row, rowNumber) {
  const enabled = row[0];
  const platform = (_cleanConfigText(row[1]) || 'AUGMENTIR').toUpperCase();
  const dataSheetName = _cleanConfigText(row[2]) || DATA_SHEET_NAME;
  const procedureName = _cleanConfigText(row[3]);
  const procedureId = _cleanConfigText(row[4]);
  const dateColumn = _cleanConfigText(row[5]);
  const dateType = (row[6] || 'AUTO').toString().trim().toUpperCase();
  const query = row[7];
  const updatedFrom = row[16];
  const updatedTo = row[17];
  const updatedFromDisplay = configSheet.getRange(rowNumber, 17).getDisplayValue();
  const updatedToDisplay = configSheet.getRange(rowNumber, 18).getDisplayValue();

  if (!_isEnabled(enabled)) return null;
  if (platform !== 'AUGMENTIR' && platform !== 'DOMO' && platform !== 'GOOGLE SHEET') {
    throw new Error(`Dòng ${rowNumber}: Platform chỉ được nhập Augmentir, Domo hoặc Google sheet`);
  }
  if (!procedureId && !procedureName) return null;
  if (platform === 'AUGMENTIR' && procedureName && procedureId) {
    throw new Error(`Dòng ${rowNumber}: Augmentir chỉ nhập một trong hai cột Procedure Name hoặc Procedure ID`);
  }
  if (platform === 'DOMO' && !procedureId) {
    throw new Error(`Dòng ${rowNumber}: Domo cần Dataset ID trong cột Procedure ID / Dataset ID`);
  }
  if (platform === 'GOOGLE SHEET' && (!procedureName || !procedureId)) {
    throw new Error(`Dòng ${rowNumber}: Google sheet cần Procedure Name là tên sheet nguồn và Procedure ID / Dataset ID là Spreadsheet ID nguồn`);
  }

  return {
    platform,
    procedureName,
    procedureId,
    datasetId: procedureId,
    dataSheetName,
    dateColumn,
    dateType,
    query,
    status: {
      created: row[8],
      inProgress: row[9],
      paused: row[10],
      cancelled: row[11],
      completed: row[12],
    },
    publishState: {
      production: row[13],
      preProduction: row[14],
      draft: row[15],
    },
    startDate: platform === 'DOMO' ? (updatedFromDisplay || updatedFrom) : updatedFrom,
    endDate: platform === 'DOMO' ? (updatedToDisplay || updatedTo) : updatedTo,
    lockHeaders: row[18],
    includeExpiredJobs: row[19],
    configSheet,
    updatedFromCell: configSheet.getRange(rowNumber, 17).getA1Notation(),
    lastUpdatedCell: configSheet.getRange(rowNumber, 21).getA1Notation(),
  };
}

function _isEnabled(value) {
  if (value === false) return false;
  if (typeof value === 'string' && ['false', 'no', 'n', '0'].includes(value.toLowerCase())) return false;
  return true;
}

function _updateProcedure(procedureConfig) {
  let posibleMore = true;
  let lastupdated = _parseFlexibleDate(procedureConfig.startDate) || new Date(1999, 1, 1);
  let latestAcceptedUpdatedAt = null;
  _ensureDataSheet(procedureConfig.dataSheetName);

  while (posibleMore) {
    const response = _getJobsFromProcedure({
      procedureId: procedureConfig.procedureId,
      procedureName: procedureConfig.procedureName,
      startDate: new Date(lastupdated.getTime() + 1).toISOString(),
      endDate: procedureConfig.endDate ? (_parseFlexibleDate(procedureConfig.endDate) || new Date(procedureConfig.endDate)).toISOString() : undefined,
    });

    const actualResults = response.actualResults || [];
    posibleMore = response.possiblyMore || false;
    Logger.log(`Augmentir rows returned for ${procedureConfig.dataSheetName}: ${actualResults.length}, possiblyMore=${posibleMore}`);

    for (const job of actualResults) {
      const _updatedAt = new Date(job.updatedAt);
      if (_updatedAt > lastupdated) lastupdated = new Date(_updatedAt);
    }

    if (actualResults.length === 0) break;
    const writeResult = writeData(actualResults, procedureConfig);
    Logger.log(`Augmentir write result for ${procedureConfig.dataSheetName}: filtered=${writeResult.filteredRows}, generated=${writeResult.generatedRows}, updated=${writeResult.updatedRows}, added=${writeResult.addedRows}`);
    if (writeResult.latestAcceptedUpdatedAt && (!latestAcceptedUpdatedAt || writeResult.latestAcceptedUpdatedAt > latestAcceptedUpdatedAt)) {
      latestAcceptedUpdatedAt = writeResult.latestAcceptedUpdatedAt;
    }
  }

  if (latestAcceptedUpdatedAt && procedureConfig.lastUpdatedCell) {
    if (procedureConfig.updatedFromCell) {
      _setConfigDateValue(procedureConfig.configSheet, procedureConfig.updatedFromCell, latestAcceptedUpdatedAt);
    }
    _setConfigDateValue(procedureConfig.configSheet, procedureConfig.lastUpdatedCell, latestAcceptedUpdatedAt);
  }
}

function _setConfigDateValue(configSheet, cellA1, value) {
  const date = _parseFlexibleDate(value);
  if (!date) {
    configSheet.getRange(cellA1).setValue(value);
    return;
  }
  configSheet.getRange(cellA1).setValue(Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'));
}

function _ensureDataSheet(dataSheetName) {
  let sheet = ss.getSheetByName(dataSheetName);
  if (!sheet) sheet = ss.insertSheet(dataSheetName);
  return sheet;
}

function _updateDomoDataset(config) {
  if (!config.dateColumn) {
    throw new Error(`Domo dataset '${config.datasetId}' chưa có Date column`);
  }

  const sql = _buildDomoQuery(config.query, config.dateColumn, config.startDate, config.endDate, config.dateType);
  Logger.log('Domo SQL: ' + sql);
  const data = _queryDomoData(config.datasetId, sql);
  Logger.log('Domo rows returned: ' + ((data.rows || []).length));
  _writeDomoDataToSheet(data, config);

  const latestDate = _getMaxDateFromDomoRows(data, config.dateColumn);
  if (latestDate && config.updatedFromCell) {
    _setConfigDateValue(config.configSheet, config.updatedFromCell, latestDate);
    _setConfigDateValue(config.configSheet, config.lastUpdatedCell, latestDate);
  }
}

function _buildDomoQuery(baseQuery, dateColumn, startDate, endDate, dateType) {
  let sql = (baseQuery || 'SELECT * FROM table').toString().trim();
  sql = sql.replace(/;+\s*$/g, '');

  const filters = [];
  const dateExpression = _formatDomoColumn(dateColumn);
  if (startDate) filters.push(`${dateExpression} > '${_formatDomoDateLiteral(startDate, dateType)}'`);
  if (endDate) filters.push(`${dateExpression} <= '${_formatDomoDateLiteral(endDate, dateType)}'`);
  if (filters.length === 0) return sql;

  const orderByMatch = sql.match(/\s+order\s+by\s+/i);
  const orderByClause = orderByMatch ? sql.slice(orderByMatch.index) : '';
  const queryWithoutOrderBy = orderByMatch ? sql.slice(0, orderByMatch.index) : sql;
  const joiner = /\swhere\s/i.test(queryWithoutOrderBy) ? ' AND ' : ' WHERE ';
  return queryWithoutOrderBy + joiner + filters.join(' AND ') + orderByClause;
}

function _formatDomoColumn(columnName) {
  const name = columnName.toString().trim();
  if (/^`.*`$/.test(name) || /^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return '`' + name.replace(/`/g, '``') + '`';
}

function _formatDomoDateLiteral(value, dateType) {
  const date = _parseFlexibleDate(value);
  if (date && dateType === 'DATE') return Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd');
  if (date && dateType === 'DATETIME') return Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
  if (date) return Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss') + '.0000000 +07:00';
  return value.toString().trim();
}

function _getDomoAccessToken() {
  const credentialSpreadsheet = SpreadsheetApp.openById(DOMO_CREDENTIAL_SPREADSHEET_ID);
  const credentialSheet = credentialSpreadsheet.getSheetByName(DOMO_CREDENTIAL_SHEET_NAME);
  if (!credentialSheet) throw new Error(`Không tìm thấy Sheet '${DOMO_CREDENTIAL_SHEET_NAME}' chứa Domo credential`);

  const clientId = credentialSheet.getRange('B1').getValue();
  const clientSecret = credentialSheet.getRange('B2').getValue();
  const options = {
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
    },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch('https://api.domo.com/oauth/token?grant_type=client_credentials&scope=data', options);
  if (response.getResponseCode() >= 300) {
    throw new Error('Không lấy được Domo access token: ' + response.getContentText());
  }
  return JSON.parse(response.getContentText()).access_token;
}

function _queryDomoData(datasetId, query) {
  const accessToken = _getDomoAccessToken();
  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      Authorization: 'bearer ' + accessToken,
    },
    payload: JSON.stringify({ sql: query }),
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch('https://api.domo.com/v1/datasets/query/execute/' + datasetId, options);
  if (response.getResponseCode() >= 300) {
    throw new Error('Không query được Domo dataset ' + datasetId + ': ' + response.getContentText());
  }
  return JSON.parse(response.getContentText());
}

function _writeDomoDataToSheet(data, config) {
  const sheet = _ensureDataSheet(config.dataSheetName);
  const columns = data.columns || [];
  const rows = data.rows || [];

  if (columns.length === 0) return;

  const keyIndex = columns.indexOf('Id');
  if (keyIndex < 0) throw new Error(`Domo dataset '${config.datasetId}' không có cột Id để upsert dữ liệu`);

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (JSON.stringify(currentHeaders.slice(0, columns.length)) !== JSON.stringify(columns)) {
      throw new Error(`Header sheet '${config.dataSheetName}' không khớp Domo dataset, dừng để tránh ghi sai dữ liệu lịch sử`);
    }
  }

  const currentRowCount = Math.max(sheet.getLastRow() - 1, 0);
  const currentRows = currentRowCount > 0
    ? sheet.getRange(2, 1, currentRowCount, columns.length).getValues()
    : [];
  const rowIndexById = {};
  currentRows.forEach((row, index) => {
    const id = row[keyIndex];
    if (id !== '' && id != null) rowIndexById[id.toString()] = index + 2;
  });

  const newRows = [];
  rows.forEach(row => {
    row.length = columns.length;
    const id = row[keyIndex];
    const targetRow = id !== '' && id != null ? rowIndexById[id.toString()] : null;
    if (targetRow) {
      sheet.getRange(targetRow, 1, 1, columns.length).setValues([row]);
    } else {
      newRows.push(row);
    }
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, columns.length).setValues(newRows);
  }
  sheet.getRange(1, 1, 1, columns.length)
    .setBackground('#b9f79c')
    .setFontColor('blue')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function _copyGoogleSheetData(config) {
  const sourceSpreadsheet = SpreadsheetApp.openById(config.procedureId);
  const sourceSheet = sourceSpreadsheet.getSheetByName(config.procedureName);
  if (!sourceSheet) {
    throw new Error(`Không tìm thấy sheet nguồn '${config.procedureName}' trong Spreadsheet '${config.procedureId}'`);
  }

  const targetSheet = _ensureDataSheet(config.dataSheetName);
  const sourceRange = sourceSheet.getDataRange();
  const values = sourceRange.getValues();

  targetSheet.clearContents();
  if (values.length === 0 || values[0].length === 0) return;

  const sourceTz = sourceSpreadsheet.getSpreadsheetTimeZone();
  const filteredValues = _filterGoogleSheetValuesByDate(values, config, sourceTz);
  if (filteredValues.length === 0 || filteredValues[0].length === 0) return;

  const sanitizedValues = _sanitizeValuesForWrite(filteredValues, sourceTz);
  targetSheet.getRange(1, 1, sanitizedValues.length, sanitizedValues[0].length).setValues(sanitizedValues);
  targetSheet.getRange(1, 1, 1, sanitizedValues[0].length)
    .setBackground('#b9f79c')
    .setFontColor('blue')
    .setFontWeight('bold');
  targetSheet.setFrozenRows(1);

  const latestDate = config.dateColumn
    ? _getMaxDateFromTableRows(filteredValues, config.dateColumn, sourceTz)
    : new Date();
  if (latestDate && config.updatedFromCell) _setConfigDateValue(config.configSheet, config.updatedFromCell, latestDate);
  if (latestDate && config.lastUpdatedCell) _setConfigDateValue(config.configSheet, config.lastUpdatedCell, latestDate);
}

function _sanitizeValuesForWrite(values, sourceTz) {
  const ssTz = sourceTz || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return values.map(row => row.map(cell => {
    if (cell instanceof Date && !isNaN(cell.getTime())) {
      const timeStr = Utilities.formatDate(cell, ssTz, "HH:mm:ss");
      if (timeStr === "00:00:00") {
        return Utilities.formatDate(cell, ssTz, "yyyy-MM-dd");
      }
      return Utilities.formatDate(cell, ssTz, "yyyy-MM-dd HH:mm:ss");
    }
    return cell;
  }));
}

function _filterGoogleSheetValuesByDate(values, config, sourceTz) {
  if (!config.dateColumn) return values;

  const headers = values[0];
  const dateIndex = headers.indexOf(config.dateColumn);
  if (dateIndex < 0) {
    throw new Error(`Không tìm thấy cột ngày '${config.dateColumn}' trong sheet nguồn '${config.procedureName}'`);
  }

  const startDate = _parseFlexibleDate(config.startDate);
  const endDate = _parseFlexibleDate(config.endDate);
  const filteredRows = values.slice(1).filter(row => {
    const rowDate = _parseFlexibleDate(row[dateIndex], sourceTz);
    if (!rowDate) return false;
    if (startDate && rowDate <= startDate) return false;
    if (endDate && rowDate > endDate) return false;
    return true;
  });

  return [headers].concat(filteredRows);
}

function _getMaxDateFromTableRows(values, dateColumn, tz) {
  if (values.length < 2) return null;
  const headers = values[0];
  const dateIndex = headers.indexOf(dateColumn);
  if (dateIndex < 0) return null;

  let maxDate = null;
  for (const row of values.slice(1)) {
    const parsedDate = _parseFlexibleDate(row[dateIndex], tz);
    if (parsedDate && (!maxDate || parsedDate > maxDate)) maxDate = parsedDate;
  }
  return maxDate;
}

function _getMaxDateFromDomoRows(data, dateColumn) {
  const columns = data.columns || [];
  const rows = data.rows || [];
  const dateIndex = columns.indexOf(dateColumn);
  if (dateIndex < 0) {
    throw new Error(`Không tìm thấy cột ngày '${dateColumn}' trong dữ liệu Domo`);
  }

  let maxDate = null;
  for (const row of rows) {
    const parsedDate = _parseFlexibleDate(row[dateIndex]);
    if (parsedDate && (!maxDate || parsedDate > maxDate)) maxDate = parsedDate;
  }
  return maxDate;
}

function _parseFlexibleDate(value, tz) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    const ssTz = tz || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const str = Utilities.formatDate(value, ssTz, "yyyy-MM-dd HH:mm:ss");
    return new Date(str.replace(/-/g, "/"));
  }

  const text = value.toString().trim();
  if (!text) return null;

  const domoTextDateTime = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?(?:\s*([+-]\d{2}):?(\d{2}))?$/);
  if (domoTextDateTime) {
    const year = Number(domoTextDateTime[1]);
    const month = Number(domoTextDateTime[2]) - 1;
    const day = Number(domoTextDateTime[3]);
    const hour = Number(domoTextDateTime[4] || 0);
    const minute = Number(domoTextDateTime[5] || 0);
    const second = Number(domoTextDateTime[6] || 0);
    const offsetHourText = domoTextDateTime[7];
    const offsetMinute = Number(domoTextDateTime[8] || 0);

    if (offsetHourText) {
      const offsetHour = Number(offsetHourText);
      const utcMillis = Date.UTC(year, month, day, hour, minute, second) - ((offsetHour * 60) + Math.sign(offsetHour) * offsetMinute) * 60000;
      return new Date(utcMillis);
    }

    return new Date(year, month, day, hour, minute, second);
  }

  const localDateTime = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$/);
  if (localDateTime) {
    const year = Number(localDateTime[1]);
    const month = Number(localDateTime[2]) - 1;
    const day = Number(localDateTime[3]);
    const hour = Number(localDateTime[4] || 0);
    const minute = Number(localDateTime[5] || 0);
    const second = Number(localDateTime[6] || 0);
    return new Date(year, month, day, hour, minute, second);
  }

  const nativeDate = new Date(text);
  if (!isNaN(nativeDate.getTime())) return nativeDate;

  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);
    const hour = Number(match[4] || 0);
    const minute = Number(match[5] || 0);
    const second = Number(match[6] || 0);
    return new Date(year, month, day, hour, minute, second);
  }

  return null;
}

function setupMultiProcedureTable() {
  const legacyConfigSheet = ss.getSheetByName(SETUP_SHEET_NAME);
  let configSheet = ss.getSheetByName(SETUP_NEW_SHEET_NAME);
  const isNewSheet = !configSheet;
  if (!configSheet) configSheet = ss.insertSheet(SETUP_NEW_SHEET_NAME);
  const wasNewLayout = configSheet.getRange(2, 2).getValue() === 'Platform';
  const migratedRows = wasNewLayout ? [] : _getMigratedSetupRows(configSheet);

  configSheet.getRange('A1:U1')
    .setValues([['MULTI PLATFORM JOB REPORT', '', '', '', '', 'Domo / Google sheet', '', '', 'Status', '', '', '', '', 'Publish state', '', '', 'Updated', '', 'Options', '', 'Sync']])
    .setBackground('#008b18')
    .setFontColor('white')
    .setFontWeight('bold');

  const headers = [
    'Enabled',
    'Platform',
    'Data sheet',
    'Procedure Name / Source sheet',
    'Procedure ID / Dataset ID / Spreadsheet ID',
    'Date column',
    'Date type',
    'Query',
    'created',
    'in-progress',
    'paused',
    'cancelled',
    'completed',
    'production',
    'pre-production',
    'draft',
    'from',
    'to',
    'Lock headers',
    'Include expired jobs',
    'Last Updated',
  ];

  configSheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#4caf50')
    .setFontColor('white')
    .setFontWeight('bold');

  const defaultRow = [
    true,
    'Augmentir',
    DATA_SHEET_NAME,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PROCEDURE.NAME).getValue() : '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PROCEDURE.ID).getValue() : '',
    '',
    'AUTO',
    '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.CREATED).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.INPROGRESS).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.PAUSED).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.CANCELLED).getValue() : false,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.COMPLETED).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PRODUCTION).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PREPRODUCTION).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.DRAFT).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.START_DATE).getValue() : '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.END_DATE).getValue() : '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.LOCK_HEADERS).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.INCLUDE_EXPIRED_JOBS).getValue() : false,
    '',
  ];

  if (migratedRows.length > 0) {
    configSheet.getRange(3, 1, Math.max(configSheet.getLastRow() - 2, 1), headers.length).clearContent();
    configSheet.getRange(3, 1, migratedRows.length, headers.length).setValues(migratedRows);
  }

  const hasProcedure = configSheet.getLastRow() >= 3
    && configSheet.getRange(3, 4, Math.max(configSheet.getLastRow() - 2, 1), 2).getValues().some(row => row[0] || row[1]);
  if ((isNewSheet || !hasProcedure) && migratedRows.length === 0) {
    configSheet.getRange(3, 1, 1, headers.length).setValues([defaultRow]);
  }
  configSheet.getRange(3, 1, 50, 1).insertCheckboxes();
  configSheet.getRange(3, 9, 50, 8).insertCheckboxes();
  configSheet.getRange(3, 19, 50, 2).insertCheckboxes();
  configSheet.setFrozenRows(2);
  configSheet.autoResizeColumns(1, headers.length);
}

function _getMigratedSetupRows(configSheet) {
  if (configSheet.getLastRow() < 3) return [];
  const oldRows = configSheet.getRange(3, 1, configSheet.getLastRow() - 2, 17).getValues();
  return oldRows
    .filter(row => row[2] || row[3])
    .map(row => [
      row[0],
      'Augmentir',
      row[1] || DATA_SHEET_NAME,
      row[2],
      row[3],
      '',
      'AUTO',
      '',
      row[4],
      row[5],
      row[6],
      row[7],
      row[8],
      row[9],
      row[10],
      row[11],
      row[12],
      row[13],
      row[14],
      row[15],
      row[16],
    ]);
}

function createReadmeSheet() {
  const sheetName = 'Readme';
  let readmeSheet = ss.getSheetByName(sheetName);
  if (!readmeSheet) readmeSheet = ss.insertSheet(sheetName);
  readmeSheet.clear();

  const rows = [
    ['CONTROL ROOM & MIXING AUTOMATION README', '', ''],
    ['Last generated', Utilities.formatDate(new Date(), LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'), 'Timezone: ' + LOCAL_TIME_ZONE],
    ['', '', ''],
    ['1. Mục đích tổng thể', 'Hệ thống tự động hóa khu vực Mixing: Đồng bộ dữ liệu từ Augmentir/Domo, tính toán KPI, gửi cảnh báo Google Chat, cung cấp API thời gian thực cho Web App Dashboard và ghi lịch sử Snapshot cho Domo.', 'Nguồn hỗ trợ: Augmentir, Domo, Google Sheets.'],
    ['2. Các sheet cốt lõi', 'SETUP_NEW (Cấu hình nguồn dữ liệu), DATA (Lưu trữ checklist), CR_History_Snapshot (Lịch sử DWH tiếng Anh), Master_data (Danh sách máy).', 'Không thay đổi cấu trúc của các sheet này.'],
    ['3. Luồng chạy chính', 'Menu Augmentir > Update hoặc trigger tự động gọi hàm update() để sync; trigger gọi triggerControlRoom() để ghi snapshot ca.', 'update() đọc cấu hình SETUP_NEW và đồng bộ dữ liệu.'],
    ['', '', ''],
    ['SETUP_NEW - Cột', 'Ý nghĩa', 'Ghi chú'],
    ['Enabled', 'TRUE/FALSE để bật hoặc bỏ qua dòng cấu hình.', 'FALSE thì dòng đó không chạy.'],
    ['Platform', 'Loại nguồn dữ liệu.', 'Augmentir, DOMO, hoặc GOOGLE SHEET.'],
    ['Data sheet', 'Sheet đích lưu trữ dữ liệu trong file này.', 'Ví dụ: DATA, DATA_MixingTransaction, DATA_Changeover-Mixing.'],
    ['Procedure Name / Source sheet', 'Augmentir: Tên quy trình. Google Sheets: Tên sheet nguồn.', 'DOMO để trống.'],
    ['Procedure ID / Dataset ID / Spreadsheet ID', 'Augmentir: Procedure ID. DOMO: Dataset ID. Google Sheets: Spreadsheet ID nguồn.', 'Dùng để kết nối API.'],
    ['Date column', 'Tên cột ngày dùng để lọc dữ liệu.', 'Domo và Google Sheets dùng được. Augmentir không cần.'],
    ['Date type', 'Kiểu dữ liệu ngày.', 'AUTO, TEXT, DATE, DATETIME.'],
    ['Query', 'SQL Query dùng cho Domo.', 'Để trống sẽ tự động lấy toàn bộ bảng.'],
    ['created/in-progress/paused/cancelled/completed', 'Bộ lọc trạng thái cho Augmentir.', 'Platform khác bỏ qua.'],
    ['production/pre-production/draft', 'Bộ lọc trạng thái phát hành cho Augmentir.', 'Platform khác bỏ qua.'],
    ['from / to', 'Mốc thời gian lấy dữ liệu (Incremental).', 'from sẽ tự động cập nhật mốc mới nhất sau mỗi lần sync thành công.'],
    ['Lock headers', 'Khóa tiêu đề cột cho Augmentir.', 'TRUE sẽ không cho phép tự động thêm cột mới.'],
    ['Include expired jobs', 'Có lấy job hết hạn của Augmentir không.', 'TRUE thì lấy cả expired jobs.'],
    ['Last Updated', 'Thời điểm cập nhật gần nhất.', 'Ghi nhận theo giờ Việt Nam.'],
    ['', '', ''],
    ['Control Room & Snapshot Engine', 'Tính năng giám sát và lưu trữ lịch sử DWH', 'Tệp tin logic chính: ControlRoom.gs.js'],
    ['Dashboard API', 'Cung cấp dữ liệu real-time cho Web App Dashboard qua getDashboardData().', 'Hỗ trợ tham số forceRefresh để bypass bộ nhớ đệm (cache 3 phút).'],
    ['Timezone Safety (An toàn múi giờ)', 'Sử dụng Unix timestamp trích xuất trực tiếp từ 8 ký tự đầu của jobId (MongoDB ObjectId).', 'Khắc phục hoàn toàn lỗi lệch 14 tiếng (America/Los_Angeles) khi vận hành ca sáng (Ca 1).'],
    ['DWH Snapshot (Lịch sử Domo)', 'Ghi lại trạng thái Checklist, Changeover và Errors của mỗi ca trực vào sheet CR_History_Snapshot.', 'Tự động dịch sang tiếng Anh, xóa Emojis, dùng chuẩn snake_case chuẩn hóa dữ liệu cho Domo.'],
    ['Công cụ di dân (Migration Tool)', 'Hàm migrateSnapshotSheet() dùng để chuẩn hóa và dịch toàn bộ dữ liệu lịch sử cũ sang tiếng Anh.', 'Chạy thủ công một lần từ Apps Script Editor.'],
    ['', '', ''],
    ['NGUYÊN TẮC HOẠT ĐỘNG & LOGIC CHÍNH', 'Chi tiết cách tính toán và tần suất', 'Thông tin kỹ thuật chuyên sâu'],
    ['1. Tần suất cập nhật (Frequency)', '1. Đồng bộ dữ liệu thô: Hàm update() chạy qua Triggers tự động của Google Apps Script (thiết lập chạy mỗi 30 phút hoặc 1 tiếng) để kéo dữ liệu mới từ Domo/Augmentir.\n2. Cập nhật Dashboard: Giao diện Web App có bộ đếm tự động làm mới mỗi 5 phút. Dữ liệu phía backend được lưu đệm (Cache) trong 180 giây (3 phút) để tối ưu hiệu năng. Bấm "⟳ Ca hiện tại" để ép buộc tải lại dữ liệu mới nhất tức thì.', 'Có thể chạy thủ công qua Menu Augmentir > Update.'],
    ['2. Logic Machine Checklist', '1. Danh sách máy mục tiêu: Lấy từ cột A của sheet "Master_data" (18 máy).\n2. Xác định ca/ngày làm việc: Đọc cột "jobId" và trích xuất timestamp gốc của MongoDB, chuyển đổi sang giờ Việt Nam (GMT+7) để xác định chính xác ngày và ca của dòng dữ liệu (bỏ qua lệch múi giờ America/Los_Angeles).\n3. Trạng thái hiển thị: Khớp mẻ mới nhất của từng máy trong ca mục tiêu:\n- Trạng thái hoàn thành (completed) -> "Completed".\n- Trạng thái đang làm (created/in-progress/paused) -> "In Progress".\n- Không có dữ liệu -> "Not Started".', 'Đảm bảo không bỏ sót mẻ trộn nào của ca sáng.'],
    ['3. Logic Changeover Compliance', '1. Phát hiện sự đổi mẻ: Đọc toàn bộ lịch sử trộn của từng máy trong "DATA_MixingTransaction". Sắp xếp theo thời gian tăng dần và đối chiếu cặp giao dịch liên ca.\n2. Cảnh báo Changeover đầu ca/trong ca: Nếu giao dịch sau nằm trong ca mục tiêu và có sự thay đổi ít nhất 1 trong 2 thông số: ProductName hoặc Color (khác rỗng, N/A được xem là 1 màu hợp lệ) -> Kích hoạt yêu cầu làm Changeover.\n3. Khớp nối: Đối chiếu với các phiếu đã làm trong "DATA_Changeover-Mixing" theo 3 cấp độ (Khớp mẻ chi tiết -> Khớp theo sản phẩm -> Khớp màu của cùng sản phẩm cũ) để hiển thị trạng thái tương ứng (Đã thực hiện, Chờ duyệt, Đang thực hiện, hoặc Thiếu).', 'Hỗ trợ so sánh liên ca (Cross-shift) hoàn hảo.'],
    ['4. Logic Error Alerts', '1. Nguồn dữ liệu: Quét sheet "History_Error_Mixing_none_Match_Data".\n2. Bộ lọc: Lấy tất cả các lỗi xảy ra trùng Ngày và Ca mục tiêu.\n3. Phân loại hiển thị: Hiển thị danh sách các lỗi có Trạng thái khắc phục là "Chưa khắc phục" hoặc trống để cảnh báo đỏ lên Dashboard và hệ thống Google Chat.', 'Lỗi đã khắc phục sẽ tự động ẩn khỏi danh sách cảnh báo chủ đạo.'],
    ['5. Cơ chế Lưu trữ Tự động (Archiving)', 'Hệ thống tự động di chuyển các bản ghi cũ hơn 30 ngày từ 5 sheet chính sang tệp lưu trữ "Mixing_Data_Archive_[Năm]" đặt tại Google Drive Folder ID "1i_FA0pNDasWPt0l9TK3uA7fZBHLPz1DS". Chỉ giữ lại 30 ngày dữ liệu hoạt động trực tuyến để tối ưu hiệu năng.', 'Có thể kích hoạt thủ công qua Menu Augmentir > Archive Old Data (30 Days) hoặc thiết lập Trigger tự động hàng tuần.'],
    ['', '', ''],
    ['Platform: Augmentir', 'Lấy danh sách jobs từ API Augmentir.', 'Ghi vào sheet DATA theo mốc incremental dựa trên updatedAt.'],
    ['Platform: DOMO', 'Đồng bộ giao dịch sản xuất từ DOMO Dataset API.', 'Ghi theo upsert mode bằng cột Id để giữ nguyên lịch sử giao dịch.'],
    ['Platform: GOOGLE SHEET', 'Copy dữ liệu trực tiếp từ các file Spreadsheet dùng chung khác.', 'Clear sheet đích và chép đè dữ liệu đã lọc theo ngày.'],
    ['', '', ''],
    ['Quy tắc Ca trực (Shifts)', 'Ca 1: 06:00 - 13:59:59 | Ca 2: 14:00 - 21:59:59 | Ca 3: 22:00 - 05:59:59 (ngày hôm sau).', 'Giờ từ 00:00 đến 05:59:59 sáng vẫn được quy về ngày làm việc hôm trước.'],
    ['Google Chat Alerts', 'Gửi báo cáo KPI tự động qua Webhook.', 'Phân loại chi tiết các máy chưa kiểm tra, đang thực hiện, chờ duyệt, hoàn thành.'],
    ['Lưu ý vận hành', '1. Không thay đổi cấu trúc cột của sheet DATA, Master_data, CR_History_Snapshot.\n2. Không đổi múi giờ của script (luôn để Asia/Bangkok).\n3. Không chia sẻ file chứa API Key / Webhook URL ra ngoài.', 'Vi phạm có thể làm hỏng luồng báo cáo hoặc rò rỉ dữ liệu bảo mật.'],
  ];

  readmeSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  readmeSheet.getRange(1, 1, 1, 3)
    .setBackground('#008b18')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(8, 1, 1, 3)
    .setBackground('#4caf50')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(24, 1, 1, 3)
    .setBackground('#1a237e')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(30, 1, 1, 3)
    .setBackground('#2e7d32')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(37, 1, 1, 3)
    .setBackground('#0288d1')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(1, 1, rows.length, 3).setWrap(true).setVerticalAlignment('top');
  readmeSheet.setFrozenRows(2);
  readmeSheet.setColumnWidths(1, 1, 220);
  readmeSheet.setColumnWidths(2, 1, 420);
  readmeSheet.setColumnWidths(3, 1, 520);
}

function readSheetList() {
  return ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn(),
  }));
}

function readSheetChunk(sheetName, startRow, rowCount) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Không tìm thấy sheet '${sheetName}'`);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0 || startRow > lastRow) return [];

  const numRows = Math.min(rowCount, lastRow - startRow + 1);
  return sheet.getRange(startRow, 1, numRows, lastColumn).getDisplayValues();
}

function exportAiSheetSnapshots() {
  const folderName = 'AI Sheet Snapshots';
  const folder = _getOrCreateSiblingFolder(folderName);
  const maxRows = 200;

  for (const sheet of ss.getSheets()) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow === 0 || lastColumn === 0) continue;

    const rowCount = Math.min(lastRow, maxRows);
    const values = sheet.getRange(1, 1, rowCount, lastColumn).getDisplayValues();
    const csv = _valuesToCsv(values);
    const fileName = _safeFileName(sheet.getName()) + '.csv';
    _replaceFileInFolder(folder, fileName, csv, MimeType.CSV);
  }

  const manifest = {
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    generatedAt: Utilities.formatDate(new Date(), LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'),
    maxRowsPerSheet: maxRows,
    sheets: readSheetList(),
  };
  _replaceFileInFolder(folder, 'manifest.json', JSON.stringify(manifest, null, 2), MimeType.PLAIN_TEXT);
}

function _getOrCreateSiblingFolder(folderName) {
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folders = parent.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
}

function _replaceFileInFolder(folder, fileName, content, mimeType) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) files.next().setTrashed(true);
  folder.createFile(fileName, content, mimeType);
}

function _valuesToCsv(values) {
  return values.map(row => row.map(value => {
    const text = value == null ? '' : value.toString();
    return /[",\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  }).join(',')).join('\r\n');
}

function _safeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_');
}

function writeData(actualResults, procedureConfig) {
  const dataSet = actualResults.filter(x => [
    procedureConfig.status.created ? 'created' : undefined,
    procedureConfig.status.inProgress ? 'in-progress' : undefined,
    procedureConfig.status.paused ? 'paused' : undefined,
    procedureConfig.status.cancelled ? 'cancelled' : undefined,
    procedureConfig.status.completed ? 'completed' : undefined,
  ].includes(x.status))
    .filter(x => [
      procedureConfig.publishState.production ? 'production' : undefined,
      procedureConfig.publishState.preProduction ? 'pre-production' : undefined,
      procedureConfig.publishState.draft ? 'draft' : undefined,
      procedureConfig.publishState.draft ? null : undefined,
    ].includes(x.publishState))
    .filter(x => {
      if (procedureConfig.includeExpiredJobs) return true;
      return x.isExpired === false;
    });
  let latestAcceptedUpdatedAt = null;
  dataSet.forEach(job => {
    const updatedAt = new Date(job.updatedAt);
    if (!isNaN(updatedAt.getTime()) && (!latestAcceptedUpdatedAt || updatedAt > latestAcceptedUpdatedAt)) {
      latestAcceptedUpdatedAt = updatedAt;
    }
  });

  const sheet = _ensureDataSheet(procedureConfig.dataSheetName);
  const lockHeaders = procedureConfig.lockHeaders;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  let headers = [...DEFAULT_HEADERS];

  if (lastRow > 0 && lastColumn > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  if (JSON.stringify(headers.slice(0, DEFAULT_HEADERS.length))
    !== JSON.stringify(DEFAULT_HEADERS)) {
    throw new Error('Headers mặc định đã bị thay đổi, không thể tiếp tục');
  }

  const newData = [];

  const formatValue = (v, key) => {
    if (v instanceof Array) return v.join('\n');
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) {
      const d = new Date(v);
      if (!isNaN(d.getTime())) {
        // Cố định múi giờ hiển thị cho các cột hệ thống của Augmentir 
        // để tránh lỗi Google Sheet tự đổi theo timezone của file
        if (['createdAt', 'startedAt', 'updatedAt', 'completedAt'].includes(key)) {
          return Utilities.formatDate(d, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
        }
        return d;
      }
    }
    return v;
  }

  for (const job of dataSet) {

    const unitDataList = job.data.unitData?.length ? job.data.unitData : [{}];
    for (let i = 0; i < unitDataList.length; i++) {
      const row = [];
      const unitData = unitDataList[i];
      const obj = {
        ..._.mapValues(_.pick(job, DEFAULT_HEADERS), formatValue),
        unitCount: i + 1,
        ..._.mapValues(_.mapKeys(unitData, (v, k) => `unitData.${k}`), formatValue),
        ..._.mapValues(_.mapKeys(job.data, (v, k) => `data.${k}`), formatValue),
      };
      delete obj.unitData;

      const pickByDefaultHeaders = _.pick(obj, headers);

      for (const h of Object.keys(pickByDefaultHeaders)) {
        const index = headers.indexOf(h);
        if (index >= 0) row[index] = pickByDefaultHeaders[h];
      }

      if (!lockHeaders) {
        const pickByNewHeaders = _.omit(obj, [...headers, 'data.unitData']);
        headers.push(...Object.keys(pickByNewHeaders));
        row.push(...Object.values(pickByNewHeaders));
      }

      newData.push(row);

    }

  }

  const headersRange = sheet.getRange(1, 1, 1, headers.length);
  const defaultHeaderRange = sheet.getRange(1, 1, 1, DEFAULT_HEADERS.length);

  headersRange.setValues([headers]).setBackground('#b9f79c').setFontColor('blue').setFontWeight('bold');
  defaultHeaderRange.setBackground('black').setFontColor('white').setFontWeight('bold');
  _formatAugmentirDateColumns(sheet, headers);
  sheet.setFrozenRows(1);

  let currentData = [];
  if (sheet.getLastRow() >= 2) currentData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  const needToUpdatedList = [];
  const newRowList = [];

  const jobIdIndex = headers.indexOf('jobId');
  const updatedAtIndex = headers.indexOf('updatedAt');
  const unitCountIndex = headers.indexOf('unitCount');

  for (const _newRow of newData) {
    const jobId = _newRow[jobIdIndex];
    const updatedAt = _newRow[updatedAtIndex];
    const unitCount = _newRow[unitCountIndex];
    const currentRowIndex = currentData.findIndex(x => x[jobIdIndex] == jobId && x[unitCountIndex] == unitCount);
    if (currentRowIndex >= 0) {
      const currentRow = currentData[currentRowIndex];
      const currentRowUpdatedAt = currentRow[updatedAtIndex];

      if (new Date(currentRowUpdatedAt).getTime() !== new Date(updatedAt).getTime()) {
        needToUpdatedList.push({
          rowIndex: currentRowIndex + 2,
          values: _newRow,
        });
      }
      continue;
    }

    newRowList.push(_newRow);
  }

  for (const item of needToUpdatedList) {
    const values = [...item.values];
    values.length = headers.length;
    sheet.getRange(item.rowIndex, 1, 1, headers.length).setValues([values]);
    console.log('-> Row #' + item.rowIndex + ' has been updated!');
  }

  newRowList.forEach(x => x.length = headers.length);

  if (newRowList.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRowList.length, headers.length).setValues(newRowList);
    console.log('=> ' + newRowList.length + ' new row(s) added!')
  }

  return {
    receivedRows: actualResults.length,
    filteredRows: dataSet.length,
    generatedRows: newData.length,
    updatedRows: needToUpdatedList.length,
    addedRows: newRowList.length,
    latestAcceptedUpdatedAt,
  };
}

function _formatAugmentirDateColumns(sheet, headers) {
  ['createdAt', 'startedAt', 'updatedAt', 'completedAt'].forEach(header => {
    const index = headers.indexOf(header);
    if (index < 0 || sheet.getMaxRows() < 2) return;
    sheet.getRange(2, index + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  });
}


function _getJobsFromProcedure({
  procedureId,
  procedureName,
  startDate,
  endDate,
}) {
  const options =
  {
    method: "POST",
    excludeArchived: true,
    payload: JSON.stringify({
      procedureId: procedureId || '',
      procedureName: procedureName || '',
      startDate: startDate || '',
      endDate: endDate || '',
      status: '',
      excludeArchived: true,
      limit: LIMIT,
      pagingupdate: true,
    }),
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { "X-aug-api-key": APP_KEY, "Content-Type": "application/json" }
  };
  try {
    const target = procedureId ? `procedureId=${procedureId}` : `procedureName=${procedureName || '(blank)'}`;
    Logger.log(`Augmentir request: ${target}, startDate=${startDate || '(blank)'}, endDate=${endDate || '(blank)'}`);
    const response = UrlFetchApp.fetch(SERVER_ENDPOINT + '/rest/v1/GetJobsFromProcedure', options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    if (statusCode >= 300) {
      throw new Error(`Augmentir API HTTP ${statusCode} for ${target}. Body: ${responseText || '(empty)'}`);
    }
    return JSON.parse(responseText);
  } catch (err) {
    if (err && err.message) throw new Error(err.message);
    throw new Error('Không thể lấy dữ liệu, vui lòng kiểm tra lại thông tin procedure (tên hoặc id)')
  }
}
