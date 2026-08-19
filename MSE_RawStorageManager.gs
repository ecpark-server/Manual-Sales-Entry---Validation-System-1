/**
 * ============================================================
 * MSE VALIDATOR
 * SEPARATE RAW WORKBOOK STORAGE MANAGER
 * ============================================================
 *
 * PURPOSE
 * ------------------------------------------------------------
 * Moves the RAW transaction index out of the main MSE database.
 *
 * STORAGE MODEL
 *
 * MSE DATABASE
 *   └── RAW_STORAGE_CATALOG
 *
 * MSE RAW DATABASES folder
 *   └── YYYYMM
 *       ├── YYYYMM_PART-01 (Google Spreadsheet)
 *       ├── YYYYMM_PART-02 (Google Spreadsheet)
 *       └── ...
 *
 * Each PART workbook contains one sheet named RAW_INDEX.
 *
 * IMPORTANT
 * ------------------------------------------------------------
 * This manager is intentionally independent from the existing
 * RAW_FILE_CATALOG and does not change the source Excel workflow.
 *
 * It is designed to store the existing 12-column RAW index rows
 * produced by MSE_RawIngestion.gs.
 *
 * ============================================================
 */


const MSE_RAW_STORAGE_CONFIG = {

  ROOT_FOLDER_PROPERTY:
    "MSE_RAW_DATABASE_FOLDER_ID",

  ROOT_FOLDER_NAME:
    "MSE RAW DATABASE",

  PERIOD_FOLDER_PREFIX:
    "",

  PART_PREFIX:
    "_PART-",

  SHEET_NAME:
    "RAW_INDEX",

  /**
   * 12 columns × 500,000 rows = 6,000,000 cells.
   * This leaves substantial room below the 10M-cell limit.
   */
  MAX_DATA_ROWS_PER_PART:
    500000,

  CATALOG_SHEET:
    "RAW_STORAGE_CATALOG"

};


const MSE_RAW_STORAGE_HEADERS = [

  "RAW_FILE_ID",
  "RAW_ROW_ID",
  "CARD_NUMBER",
  "TRANSACTION_DATE",
  "DATE_KEY",
  "REGISTER_NO",
  "TRANSACTION_NO",
  "STORE_CODE",
  "TENDER_TYPE",
  "TENDER_AMOUNT",
  "LOOKUP_KEY",
  "ROW_STATUS"

];


const MSE_RAW_STORAGE_CATALOG_HEADERS = [

  "STORAGE_ID",
  "PERIOD",
  "PART",
  "SPREADSHEET_ID",
  "SPREADSHEET_NAME",
  "SHEET_NAME",
  "ROW_COUNT",
  "MAX_ROWS",
  "STATUS",
  "ACTIVE",
  "CREATED_TIMESTAMP",
  "UPDATED_TIMESTAMP"

];


/**
 * ============================================================
 * INITIALIZE RAW STORAGE
 * ============================================================
 *
 * Creates the parent RAW DATABASE folder and catalog if needed.
 *
 * Run once from Apps Script:
 *
 *   MSE_RAW_STORAGE_Initialize()
 *
 * ============================================================
 */
function MSE_RAW_STORAGE_Initialize() {

  const rootFolder =
    MSE_RAW_STORAGE_GetRootFolder_();

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  return {

    success: true,

    rootFolderId:
      rootFolder.getId(),

    rootFolderName:
      rootFolder.getName(),

    catalogSpreadsheetId:
      getMSEDatabase_().getId(),

    catalogSheet:
      catalog.getName(),

    message:
      "Separate RAW workbook storage is initialized."

  };

}


/**
 * ============================================================
 * GET / CREATE ROOT FOLDER
 * ============================================================
 */
function MSE_RAW_STORAGE_GetRootFolder_() {

  const props =
    PropertiesService
      .getScriptProperties();

  const configuredId =
    String(
      props.getProperty(
        MSE_RAW_STORAGE_CONFIG.ROOT_FOLDER_PROPERTY
      ) || ""
    ).trim();

  if (configuredId) {

    try {

      return DriveApp.getFolderById(
        configuredId
      );

    } catch (error) {

      Logger.log(
        "Configured RAW database folder is no longer accessible. A new folder will be created."
      );

    }

  }

  const parent =
    DriveApp.getRootFolder();

  const folders =
    parent.getFoldersByName(
      MSE_RAW_STORAGE_CONFIG.ROOT_FOLDER_NAME
    );

  let folder =
    folders.hasNext()
      ? folders.next()
      : parent.createFolder(
          MSE_RAW_STORAGE_CONFIG.ROOT_FOLDER_NAME
        );

  props.setProperty(
    MSE_RAW_STORAGE_CONFIG.ROOT_FOLDER_PROPERTY,
    folder.getId()
  );

  return folder;

}


/**
 * ============================================================
 * GET / CREATE PERIOD FOLDER
 * ============================================================
 */
function MSE_RAW_STORAGE_GetPeriodFolder_(
  period
) {

  const normalized =
    MSE_RAW_STORAGE_NormalizePeriod_(
      period
    );

  if (!normalized) {

    throw new Error(
      "RAW storage period must be YYYYMM. Received: " +
      period
    );

  }

  const root =
    MSE_RAW_STORAGE_GetRootFolder_();

  const folders =
    root.getFoldersByName(
      normalized
    );

  return folders.hasNext()
    ? folders.next()
    : root.createFolder(
        normalized
      );

}


/**
 * ============================================================
 * NORMALIZE PERIOD
 * ============================================================
 */
function MSE_RAW_STORAGE_NormalizePeriod_(
  period
) {

  const value =
    String(
      period === null ||
      period === undefined
        ? ""
        : period
    ).trim();

  if (/^\d{6}$/.test(value)) {

    return value;

  }

  if (/^\d{8}$/.test(value)) {

    return value.substring(0, 6);

  }

  return "";

}


/**
 * ============================================================
 * GET / CREATE STORAGE CATALOG
 * ============================================================
 */
function MSE_RAW_STORAGE_GetCatalogSheet_() {

  const ss =
    getMSEDatabase_();

  let sheet =
    ss.getSheetByName(
      MSE_RAW_STORAGE_CONFIG.CATALOG_SHEET
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        MSE_RAW_STORAGE_CONFIG.CATALOG_SHEET
      );

    sheet
      .getRange(
        1,
        1,
        1,
        MSE_RAW_STORAGE_CATALOG_HEADERS.length
      )
      .setValues([
        MSE_RAW_STORAGE_CATALOG_HEADERS
      ])
      .setFontWeight("bold");

    sheet.setFrozenRows(1);

  } else {

    const existing =
      sheet
        .getRange(
          1,
          1,
          1,
          MSE_RAW_STORAGE_CATALOG_HEADERS.length
        )
        .getValues()[0];

    const repaired =
      MSE_RAW_STORAGE_CATALOG_HEADERS.map(
        function(header, index) {
          return existing[index] === header
            ? existing[index]
            : header;
        }
      );

    sheet
      .getRange(
        1,
        1,
        1,
        MSE_RAW_STORAGE_CATALOG_HEADERS.length
      )
      .setValues([repaired]);

  }

  return sheet;

}


/**
 * ============================================================
 * GET ACTIVE PART
 * ============================================================
 */
function MSE_RAW_STORAGE_GetActivePart_(
  period,
  rowsNeeded
) {

  const normalizedPeriod =
    MSE_RAW_STORAGE_NormalizePeriod_(
      period
    );

  if (!normalizedPeriod) {
    throw new Error("Invalid RAW storage period: " + period);
  }

  const needed =
    Math.max(1, Number(rowsNeeded) || 1);

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  const data =
    catalog.getDataRange().getValues();

  const periodRows = [];

  for (let r = 1; r < data.length; r++) {

    const rowPeriod =
      String(data[r][1] || "").trim();

    const active =
      data[r][9] === true ||
      String(data[r][9] || "").toUpperCase() === "TRUE";

    const status =
      String(data[r][8] || "").toUpperCase();

    if (
      rowPeriod === normalizedPeriod &&
      active &&
      status !== "FULL"
    ) {

      periodRows.push({
        row: r + 1,
        part: Number(data[r][2]) || 0,
        spreadsheetId: String(data[r][3] || "").trim(),
        spreadsheetName: String(data[r][4] || "").trim(),
        sheetName: String(data[r][5] || "").trim(),
        rowCount: Number(data[r][6]) || 0,
        maxRows: Number(data[r][7]) || MSE_RAW_STORAGE_CONFIG.MAX_DATA_ROWS_PER_PART
      });

    }

  }

  periodRows.sort(function(a, b) {
    return a.part - b.part;
  });

  for (let i = 0; i < periodRows.length; i++) {

    const item = periodRows[i];

    if (
      item.rowCount + needed <=
      item.maxRows
    ) {

      return item;

    }

    catalog
      .getRange(item.row, 9)
      .setValue("FULL");

    catalog
      .getRange(item.row, 10)
      .setValue(false);

  }

  return MSE_RAW_STORAGE_CreatePart_(
    normalizedPeriod,
    periodRows.length
      ? periodRows[periodRows.length - 1].part + 1
      : 1
  );

}


/**
 * ============================================================
 * CREATE PART WORKBOOK
 * ============================================================
 */
function MSE_RAW_STORAGE_CreatePart_(
  period,
  partNumber
) {

  const periodFolder =
    MSE_RAW_STORAGE_GetPeriodFolder_(
      period
    );

  const part =
    String(partNumber).padStart(2, "0");

  const name =
    period +
    MSE_RAW_STORAGE_CONFIG.PART_PREFIX +
    part;

  const ss =
    SpreadsheetApp.create(name);

  const file =
    DriveApp.getFileById(
      ss.getId()
    );

  periodFolder.addFile(file);

  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (error) {
    // The file may already have no root-folder membership.
  }

  let sheet =
    ss.getSheetByName("Sheet1");

  if (!sheet) {
    sheet = ss.insertSheet("RAW_INDEX");
  }

  sheet.setName(
    MSE_RAW_STORAGE_CONFIG.SHEET_NAME
  );

  sheet
    .getRange(
      1,
      1,
      1,
      MSE_RAW_STORAGE_HEADERS.length
    )
    .setValues([
      MSE_RAW_STORAGE_HEADERS
    ])
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  const storageId =
    period + "|PART-" + part;

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  const now =
    new Date();

  catalog.appendRow([

    storageId,
    period,
    Number(part),
    ss.getId(),
    name,
    MSE_RAW_STORAGE_CONFIG.SHEET_NAME,
    0,
    MSE_RAW_STORAGE_CONFIG.MAX_DATA_ROWS_PER_PART,
    "ACTIVE",
    true,
    now,
    now

  ]);

  return {

    row:
      catalog.getLastRow(),

    part:
      Number(part),

    spreadsheetId:
      ss.getId(),

    spreadsheetName:
      name,

    sheetName:
      MSE_RAW_STORAGE_CONFIG.SHEET_NAME,

    rowCount:
      0,

    maxRows:
      MSE_RAW_STORAGE_CONFIG.MAX_DATA_ROWS_PER_PART

  };

}


/**
 * ============================================================
 * APPEND RAW INDEX ROWS
 * ============================================================
 *
 * Rows are automatically routed by TRANSACTION_DATE.
 *
 * This function accepts the same 12-column rows currently
 * generated by MSE_RawIngestion.gs.
 * ============================================================
 */
function MSE_RAW_STORAGE_AppendRows(
  rows
) {

  if (!rows || !rows.length) {
    return { rowsWritten: 0, periods: [] };
  }

  const grouped = {};

  rows.forEach(function(row) {

    if (!row || row.length < MSE_RAW_STORAGE_HEADERS.length) {
      throw new Error(
        "RAW storage row must contain exactly 12 columns."
      );
    }

    const period =
      MSE_RAW_STORAGE_NormalizePeriod_(
        row[3]
      );

    if (!period) {
      throw new Error(
        "Cannot route RAW index row because TRANSACTION_DATE is invalid: " +
        row[3]
      );
    }

    if (!grouped[period]) {
      grouped[period] = [];
    }

    grouped[period].push(row);

  });

  const periods =
    Object.keys(grouped).sort();

  let totalWritten = 0;

  periods.forEach(function(period) {

    const periodRows =
      grouped[period];

    let offset = 0;

    while (offset < periodRows.length) {

      const active =
        MSE_RAW_STORAGE_GetActivePart_(
          period,
          1
        );

      const capacity =
        active.maxRows -
        active.rowCount;

      if (capacity <= 0) {
        throw new Error(
          "RAW storage part has no capacity: " +
          active.spreadsheetName
        );
      }

      const count =
        Math.min(
          capacity,
          periodRows.length - offset
        );

      const chunk =
        periodRows.slice(
          offset,
          offset + count
        );

      const ss =
        SpreadsheetApp.openById(
          active.spreadsheetId
        );

      const sheet =
        ss.getSheetByName(
          active.sheetName
        );

      if (!sheet) {
        throw new Error(
          "RAW storage sheet not found: " +
          active.sheetName +
          " in " +
          active.spreadsheetName
        );
      }

      const startRow =
        Math.max(2, sheet.getLastRow() + 1);

      const requiredRows =
        startRow + chunk.length - 1;

      if (requiredRows > sheet.getMaxRows()) {
        sheet.insertRowsAfter(
          sheet.getMaxRows(),
          requiredRows - sheet.getMaxRows()
        );
      }

      sheet
        .getRange(
          startRow,
          1,
          chunk.length,
          MSE_RAW_STORAGE_HEADERS.length
        )
        .setValues(chunk);

      const catalog =
        MSE_RAW_STORAGE_GetCatalogSheet_();

      catalog
        .getRange(active.row, 7)
        .setValue(
          active.rowCount + chunk.length
        );

      catalog
        .getRange(active.row, 8)
        .setValue(
          active.maxRows
        );

      catalog
        .getRange(active.row, 9)
        .setValue(
          active.rowCount + chunk.length >= active.maxRows
            ? "FULL"
            : "ACTIVE"
        );

      catalog
        .getRange(active.row, 10)
        .setValue(
          active.rowCount + chunk.length < active.maxRows
        );

      catalog
        .getRange(active.row, 12)
        .setValue(new Date());

      offset += chunk.length;
      totalWritten += chunk.length;

    }

  });

  return {
    rowsWritten: totalWritten,
    periods: periods
  };

}


/**
 * ============================================================
 * GET PROCESSED ROW COUNT FOR ONE RAW FILE
 * ============================================================
 */
function MSE_RAW_STORAGE_GetProcessedRows(
  rawFileId
) {

  const target =
    String(rawFileId || "").trim();

  if (!target) {
    return 0;
  }

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  const data =
    catalog.getDataRange().getValues();

  let total = 0;

  for (let r = 1; r < data.length; r++) {

    const spreadsheetId =
      String(data[r][3] || "").trim();

    const sheetName =
      String(data[r][5] || "").trim();

    if (!spreadsheetId || !sheetName) {
      continue;
    }

    const ss =
      SpreadsheetApp.openById(
        spreadsheetId
      );

    const sheet =
      ss.getSheetByName(
        sheetName
      );

    if (!sheet || sheet.getLastRow() < 2) {
      continue;
    }

    const lastRow =
      sheet.getLastRow();

    const rawIds =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getDisplayValues();

    for (let i = 0; i < rawIds.length; i++) {
      if (String(rawIds[i][0] || "").trim() === target) {
        total++;
      }
    }

  }

  return total;

}


/**
 * ============================================================
 * CLEAR INDEX ROWS FOR RAW FILE
 * ============================================================
 *
 * Used before a complete re-processing of a RAW file.
 * ============================================================
 */
function MSE_RAW_STORAGE_ClearFile(
  rawFileId
) {

  const target =
    String(rawFileId || "").trim();

  if (!target) {
    return 0;
  }

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  const data =
    catalog.getDataRange().getValues();

  let deleted = 0;

  for (let r = 1; r < data.length; r++) {

    const spreadsheetId =
      String(data[r][3] || "").trim();

    const sheetName =
      String(data[r][5] || "").trim();

    if (!spreadsheetId || !sheetName) {
      continue;
    }

    const ss =
      SpreadsheetApp.openById(
        spreadsheetId
      );

    const sheet =
      ss.getSheetByName(
        sheetName
      );

    if (!sheet || sheet.getLastRow() < 2) {
      continue;
    }

    const lastRow =
      sheet.getLastRow();

    const ids =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getDisplayValues();

    const rows = [];

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === target) {
        rows.push(i + 2);
      }
    }

    for (let i = rows.length - 1; i >= 0; i--) {
      sheet.deleteRow(rows[i]);
      deleted++;
    }

    if (rows.length) {

      const catalogRow = r + 1;
      const newCount =
        Math.max(
          0,
          Number(data[r][6] || 0) - rows.length
        );

      catalog
        .getRange(catalogRow, 7)
        .setValue(newCount);

      catalog
        .getRange(catalogRow, 9)
        .setValue("ACTIVE");

      catalog
        .getRange(catalogRow, 10)
        .setValue(true);

      catalog
        .getRange(catalogRow, 12)
        .setValue(new Date());

    }

  }

  return deleted;

}


/**
 * ============================================================
 * LOOKUP TRANSACTION IN SEPARATE RAW WORKBOOKS
 * ============================================================
 *
 * Uses DATE to route directly to the YYYYMM partition.
 * It does not scan the MSE database RAW index.
 * ============================================================
 */
function MSE_RAW_STORAGE_LookupTransaction(
  transactionDate,
  storeCode,
  registerNo,
  transactionNo
) {

  const date =
    typeof normalizeMSERawDate_ === "function"
      ? normalizeMSERawDate_(transactionDate)
      : String(transactionDate || "").trim();

  const store =
    typeof normalizeMSERawCode_ === "function"
      ? normalizeMSERawCode_(storeCode)
      : String(storeCode || "").trim();

  const register =
    typeof normalizeMSERawCode_ === "function"
      ? normalizeMSERawCode_(registerNo)
      : String(registerNo || "").trim();

  const trans =
    typeof normalizeMSERawCode_ === "function"
      ? normalizeMSERawCode_(transactionNo)
      : String(transactionNo || "").trim();

  const lookupKey =
    [date, store, register, trans].join("|");

  const period =
    MSE_RAW_STORAGE_NormalizePeriod_(
      date
    );

  if (!period) {
    return {
      found: false,
      lookupKey: lookupKey,
      reason: "INVALID_DATE"
    };
  }

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  const data =
    catalog.getDataRange().getValues();

  for (let r = 1; r < data.length; r++) {

    if (String(data[r][1] || "").trim() !== period) {
      continue;
    }

    const spreadsheetId =
      String(data[r][3] || "").trim();

    const sheetName =
      String(data[r][5] || "").trim();

    if (!spreadsheetId || !sheetName) {
      continue;
    }

    const ss =
      SpreadsheetApp.openById(
        spreadsheetId
      );

    const sheet =
      ss.getSheetByName(
        sheetName
      );

    if (!sheet || sheet.getLastRow() < 2) {
      continue;
    }

    const values =
      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          MSE_RAW_STORAGE_HEADERS.length
        )
        .getValues();

    for (let i = 0; i < values.length; i++) {

      const row = values[i];

      const rowKey =
        [
          String(row[3] || "").trim(),
          String(row[7] || "").trim(),
          String(row[5] || "").trim(),
          String(row[6] || "").trim()
        ].join("|");

      if (rowKey !== lookupKey) {
        continue;
      }

      return {

        found: true,

        lookupKey: lookupKey,

        rawFileId: row[0],
        rawRowId: row[1],

        cardNumbers:
          row[2]
            ? String(row[2]).split(" | ")
            : [],

        transactionDate: row[3],
        registerNo: row[5],
        transactionNo: row[6],
        storeCode: row[7],

        tenderBreakdown: row[8],
        gross: Number(row[8 + 1]) || 0,
        w8: Number(row[9]) || 0,
        net: Number(row[10]) || 0,

        storagePeriod: period,
        storagePart:
          data[r][2],
        storageSpreadsheetId:
          spreadsheetId

      };

    }

  }

  return {
    found: false,
    lookupKey: lookupKey,
    storagePeriod: period
  };

}


/**
 * ============================================================
 * STORAGE HEALTH CHECK
 * ============================================================
 */
function MSE_RAW_STORAGE_HealthCheck() {

  const root =
    MSE_RAW_STORAGE_GetRootFolder_();

  const catalog =
    MSE_RAW_STORAGE_GetCatalogSheet_();

  const data =
    catalog.getDataRange().getValues();

  const parts = [];

  for (let r = 1; r < data.length; r++) {

    parts.push({

      storageId: data[r][0],
      period: data[r][1],
      part: data[r][2],
      spreadsheetId: data[r][3],
      spreadsheetName: data[r][4],
      sheetName: data[r][5],
      rowCount: data[r][6],
      maxRows: data[r][7],
      status: data[r][8],
      active: data[r][9]

    });

  }

  return {

    success: true,

    rootFolderId:
      root.getId(),

    rootFolderName:
      root.getName(),

    partCount:
      parts.length,

    parts: parts

  };

}


/**
 * ============================================================
 * TEST CREATE / WRITE STORAGE
 * ============================================================
 *
 * This writes ONLY a clearly marked sample row.
 * It is intentionally disabled by default.
 * ============================================================
 */
function TEST_MSE_RAW_STORAGE_CreateSample() {

  throw new Error(
    "Safety stop: this test intentionally does not write sample RAW data. Use MSE_RAW_STORAGE_Initialize() first, then test against an actual ingestion batch."
  );

}
