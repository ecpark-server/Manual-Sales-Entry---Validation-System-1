/**
 * ============================================================
 * MSE VALIDATOR
 * RAW EXCEL INGESTION / TRANSACTION INDEXER
 * ============================================================
 *
 * PURPOSE
 * ------------------------------------------------------------
 * Reads registered RAW Excel files and builds the
 * RAW_TRANSACTION_INDEX used by the MSE Validator.
 *
 * SOURCE COLUMNS EXPECTED
 *
 * Card Number
 * Date
 * Register Num
 * Trans Num
 * Store Code
 * Tender Type
 * Tender Amount
 *
 * IMPORTANT
 * ------------------------------------------------------------
 * The source contains TENDER-LEVEL rows.
 *
 * Multiple rows may represent ONE transaction.
 *
 * Example:
 *
 * 20260430 | 343 | 14 | 5942 | CA  | 3000
 * 20260430 | 343 | 14 | 5942 | W8  | 500
 *
 * Becomes:
 *
 * GROSS = 3500
 * W8    = 500
 * NET   = 3000
 *
 * The index therefore stores ONE aggregated record
 * per transaction.
 *
 * ============================================================
 */


/**
 * ============================================================
 * INGESTION SETTINGS
 * ============================================================
 */

const MSE_RAW_INGESTION_CONFIG = {

  /**
   * Number of source rows processed per batch.
   *
   * This controls memory usage.
   */
  CHUNK_SIZE: 5000,


  /**
   * Number of indexed transactions written per batch.
   */
  WRITE_CHUNK_SIZE: 2000,


  /**
   * Maximum number of transactions kept in memory
   * during one aggregation pass.
   *
   * The source file is processed in chunks.
   */
  MAX_MEMORY_TRANSACTIONS: 100000,


  /**
   * Temporary converted Google Sheet name prefix.
   */
  TEMP_SHEET_PREFIX:
    "_MSE_RAW_TEMP_"

};


/**
 * ============================================================
 * REQUIRED SOURCE HEADERS
 * ============================================================
 */

const MSE_RAW_REQUIRED_HEADERS = {

  CARD:
    [
      "CARD NUMBER",
      "CARDNUMBER",
      "CARD"
    ],

  DATE:
    [
      "DATE",
      "TRANSACTION DATE",
      "TRANSACTION_DATE"
    ],

  REGISTER:
    [
      "REGISTER NUM",
      "REGISTER NUMBER",
      "REGISTER NO",
      "REGISTER_NO"
    ],

  TRANS:
    [
      "TRANS NUM",
      "TRANSACTION NUM",
      "TRANSACTION NUMBER",
      "TRANS NO",
      "TRANS_NO"
    ],

  STORE:
    [
      "STORE CODE",
      "STORE_CODE",
      "STORE"
    ],

  TENDER_TYPE:
    [
      "TENDER TYPE",
      "TENDER_TYPE",
      "TENDER"
    ],

  TENDER_AMOUNT:
    [
      "TENDER AMOUNT",
      "TENDER_AMOUNT",
      "AMOUNT"
    ]

};


/**
 * ============================================================
 * NORMALIZE HEADER
 * ============================================================
 */

function normalizeMSERawHeader_(
  value
) {

  return String(
    value === null ||
    value === undefined
      ? ""
      : value
  )
    .trim()
    .toUpperCase()
    .replace(
      /\s+/g,
      " "
    );

}


/**
 * ============================================================
 * FIND HEADER COLUMN
 * ============================================================
 */

function findMSERawHeaderColumn_(
  headers,
  aliases
) {

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    const normalized =
      normalizeMSERawHeader_(
        headers[i]
      );


    if (
      aliases.indexOf(
        normalized
      ) !== -1
    ) {

      return i;

    }

  }


  return -1;

}


/**
 * ============================================================
 * DETECT SOURCE COLUMNS
 * ============================================================
 */

function detectMSERawColumns_(
  headers
) {

  const columns = {

    card:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.CARD
      ),

    date:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.DATE
      ),

    register:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.REGISTER
      ),

    trans:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.TRANS
      ),

    store:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.STORE
      ),

    tenderType:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.TENDER_TYPE
      ),

    tenderAmount:
      findMSERawHeaderColumn_(
        headers,
        MSE_RAW_REQUIRED_HEADERS.TENDER_AMOUNT
      )

  };


  const errors = [];


  Object.keys(
    columns
  ).forEach(
    key => {

      if (
        columns[key] === -1
      ) {

        errors.push(
          "Missing required column: " +
          key
        );

      }

    }
  );


  if (
    errors.length
  ) {

    throw new Error(
      "RAW Excel structure is invalid.\n\n" +
      errors.join("\n") +
      "\n\nDetected headers:\n" +
      headers.join(" | ")
    );

  }


  return columns;

}


/**
 * ============================================================
 * NORMALIZE CARD
 * ============================================================
 *
 * We preserve the source card as digits where possible.
 *
 * Card normalization rules for validation will be handled
 * separately from the RAW ingestion layer.
 * ============================================================
 */

function normalizeMSERawCard_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  const text =
    String(
      value
    ).trim();


  if (
    !text
  ) {

    return "";

  }


  if (
    text === "-"
  ) {

    return "-";

  }


  return text
    .replace(
      /\D/g,
      ""
    );

}


/**
 * ============================================================
 * NORMALIZE DATE
 *
 * Supported:
 *
 * 20260430
 * 2026-04-30
 * 04/30/2026
 * Date object
 * ============================================================
 */

function normalizeMSERawDate_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return "";

  }


  /**
   * Date object
   */
  if (
    Object.prototype.toString
      .call(value) ===
      "[object Date]"
  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return "";

    }


    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );

  }


  const text =
    String(
      value
    ).trim();


  /**
   * YYYYMMDD
   */
  if (
    /^\d{8}$/.test(
      text
    )
  ) {

    return text;

  }


  /**
   * YYYY-MM-DD
   */
  let match =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (
    match
  ) {

    return (
      match[1] +
      match[2] +
      match[3]
    );

  }


  /**
   * MM/DD/YYYY
   */
  match =
    text.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );


  if (
    match
  ) {

    return (
      match[3] +
      String(
        match[1]
      ).padStart(
        2,
        "0"
      ) +
      String(
        match[2]
      ).padStart(
        2,
        "0"
      )
    );

  }


  /**
   * Try JavaScript Date.
   */
  const parsed =
    new Date(
      text
    );


  if (
    !isNaN(
      parsed.getTime()
    )
  ) {

    return Utilities.formatDate(
      parsed,
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );

  }


  return text;

}


/**
 * ============================================================
 * NORMALIZE CODE
 * ============================================================
 */

function normalizeMSERawCode_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  const text =
    String(
      value
    ).trim();


  /**
   * Remove Excel numeric .0
   *
   * 343.0 → 343
   */
  if (
    /^\d+\.0$/.test(
      text
    )
  ) {

    return text.substring(
      0,
      text.length - 2
    );

  }


  return text;

}


/**
 * ============================================================
 * NORMALIZE AMOUNT
 * ============================================================
 */

function normalizeMSERawAmount_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return 0;

  }


  if (
    typeof value ===
    "number"
  ) {

    return isFinite(
      value
    )
      ? value
      : 0;

  }


  let text =
    String(
      value
    )
      .trim();


  if (
    !text
  ) {

    return 0;

  }


  /**
   * Remove currency symbols,
   * commas and spaces.
   */
  text =
    text
      .replace(
        /,/g,
        ""
      )
      .replace(
        /₱/g,
        ""
      )
      .replace(
        /\$/g,
        ""
      )
      .replace(
        /\s/g,
        ""
      );


  const number =
    Number(
      text
    );


  if (
    !isFinite(
      number
    )
  ) {

    return 0;

  }


  return number;

}


/**
 * ============================================================
 * ROUND MONEY
 * ============================================================
 */

function roundMSEMoney_(
  value
) {

  return Math.round(
    (
      Number(value) || 0
    ) *
    100
  ) / 100;

}


/**
 * ============================================================
 * TRANSACTION KEY
 * ============================================================
 */

function buildMSERawTransactionKey_(
  date,
  store,
  register,
  trans
) {

  return [

    date,

    store,

    register,

    trans

  ].join("|");

}


/**
 * ============================================================
 * CREATE TRANSACTION OBJECT
 * ============================================================
 */

function createMSERawTransaction_(
  date,
  store,
  register,
  trans
) {

  return {

    transactionDate:
      date,

    storeCode:
      store,

    registerNo:
      register,

    transactionNo:
      trans,

    cardNumbers:
      {},

    gross:
      0,

    w8:
      0,

    net:
      0,

    tenderBreakdown:
      {},

    tenderCount:
      0

  };

}


/**
 * ============================================================
 * ADD TENDER ROW TO TRANSACTION
 * ============================================================
 */

function addMSERawTenderRow_(
  transaction,
  card,
  tenderType,
  amount
) {

  /**
   * Card
   */
  if (
    card
  ) {

    transaction
      .cardNumbers[
        card
      ] = true;

  }


  /**
   * Tender type
   */
  const tender =
    String(
      tenderType || ""
    )
      .trim()
      .toUpperCase();


  /**
   * Amount
   */
  const money =
    roundMSEMoney_(
      amount
    );


  /**
   * Gross
   */
  transaction.gross =
    roundMSEMoney_(
      transaction.gross +
      money
    );


  /**
   * W8
   */
  if (
    tender === "W8"
  ) {

    transaction.w8 =
      roundMSEMoney_(
        transaction.w8 +
        money
      );

  }


  /**
   * Tender breakdown
   */
  if (
    !transaction
      .tenderBreakdown[
        tender
      ]
  ) {

    transaction
      .tenderBreakdown[
        tender
      ] = 0;

  }


  transaction
    .tenderBreakdown[
      tender
    ] =
    roundMSEMoney_(
      transaction
        .tenderBreakdown[
          tender
        ] +
      money
    );


  transaction.tenderCount++;

}


/**
 * ============================================================
 * FINALIZE TRANSACTION
 * ============================================================
 */

function finalizeMSERawTransaction_(
  transaction
) {

  transaction.gross =
    roundMSEMoney_(
      transaction.gross
    );


  transaction.w8 =
    roundMSEMoney_(
      transaction.w8
    );


  transaction.net =
    roundMSEMoney_(
      transaction.gross -
      transaction.w8
    );


  /**
   * W8 ONLY
   *
   * A transaction is W8-only when all gross
   * amount is W8.
   */
  transaction.w8Only =
    transaction.gross > 0 &&
    transaction.w8 ===
      transaction.gross;


  return transaction;

}


/**
 * ============================================================
 * CARD ARRAY
 * ============================================================
 */

function getMSERawCardArray_(
  cardMap
) {

  return Object.keys(
    cardMap || {}
  );

}


/**
 * ============================================================
 * TENDER BREAKDOWN STRING
 * ============================================================
 *
 * Example:
 *
 * CA:3000 | W8:500
 *
 * ============================================================
 */

function serializeMSERawTenderBreakdown_(
  breakdown
) {

  return Object.keys(
    breakdown || {}
  )
    .sort()
    .map(
      tender => {

        return (
          tender +
          ":" +
          roundMSEMoney_(
            breakdown[
              tender
            ]
          )
        );

      }
    )
    .join(
      " | "
    );

}


/**
 * ============================================================
 * LOOKUP KEY
 * ============================================================
 *
 * This is the primary transaction lookup key.
 * ============================================================
 */

function buildMSERawLookupKey_(
  transaction
) {

  return buildMSERawTransactionKey_(
    transaction.transactionDate,
    transaction.storeCode,
    transaction.registerNo,
    transaction.transactionNo
  );

}


/**
 * ============================================================
 * OPEN RAW EXCEL AS TEMP GOOGLE SHEET
 *
 * Requires Advanced Drive Service.
 *
 * IMPORTANT:
 * In Apps Script:
 *
 * Services → + → Drive API
 *
 * must be enabled.
 *
 * ============================================================
 */

function convertMSERawExcelToGoogleSheet_(
  fileId
) {

  if (
    typeof Drive ===
    "undefined"
  ) {

    throw new Error(
      "Advanced Drive Service is not enabled.\n\n" +
      "In Apps Script open:\n" +
      "Services → + → Drive API\n\n" +
      "Then enable Drive API and run again."
    );

  }


  const sourceFile =
    DriveApp.getFileById(
      fileId
    );


  const originalName =
    sourceFile.getName();


  const temporaryName =
    MSE_RAW_INGESTION_CONFIG
      .TEMP_SHEET_PREFIX +
    originalName +
    "_" +
    Date.now();


  const blob =
    sourceFile.getBlob();


  const resource = {

    name:
      temporaryName,

    mimeType:
      MimeType.GOOGLE_SHEETS

  };


  const converted =
    Drive.Files.create(
      resource,
      blob,
      {
        fields:
          "id,name,mimeType"
      }
    );


  if (
    !converted ||
    !converted.id
  ) {

    throw new Error(
      "Failed to convert RAW Excel file to temporary Google Sheet."
    );

  }


  return {

    id:
      converted.id,

    name:
      converted.name,

    mimeType:
      converted.mimeType

  };

}


/**
 * ============================================================
 * DELETE TEMPORARY GOOGLE SHEET
 * ============================================================
 */

function deleteMSERawTemporarySheet_(
  fileId
) {

  if (
    !fileId
  ) {

    return;

  }


  try {

    DriveApp
      .getFileById(
        fileId
      )
      .setTrashed(
        true
      );

  } catch (error) {

    Logger.log(
      "WARNING: Unable to delete temporary sheet: " +
      error.message
    );

  }

}


/**
 * ============================================================
 * GET RAW SHEET
 *
 * We use the first non-empty sheet.
 * ============================================================
 */

function getMSERawSourceSheet_(
  spreadsheet
) {

  const sheets =
    spreadsheet
      .getSheets();


  for (
    let i = 0;
    i < sheets.length;
    i++
  ) {

    const sheet =
      sheets[i];


    if (
      sheet.getLastRow() > 0 &&
      sheet.getLastColumn() > 0
    ) {

      return sheet;

    }

  }


  throw new Error(
    "The RAW Excel workbook contains no usable data."
  );

}


/**
 * ============================================================
 * GET SOURCE HEADERS
 * ============================================================
 */

function getMSERawSourceHeaders_(
  sheet
) {

  const lastColumn =
    sheet.getLastColumn();


  if (
    lastColumn < 1
  ) {

    throw new Error(
      "RAW source sheet contains no columns."
    );

  }


  return sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

}


/**
 * ============================================================
 * CLEAR PREVIOUS INDEX FOR RAW FILE
 *
 * This allows safe re-processing of a file.
 *
 * ============================================================
 */

function clearMSERawIndexForFile_(
  rawFileId
) {

  const sheet =
    getMSERawIndexSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return 0;

  }


  const lastColumn =
    Math.max(
      11,
      sheet.getLastColumn()
    );


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();


  const target =
    String(
      rawFileId
    ).trim();


  const rowsToDelete = [];


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    if (
      String(
        data[i][0] || ""
      ).trim() ===
      target
    ) {

      rowsToDelete.push(
        i + 2
      );

    }

  }


  for (
    let i =
      rowsToDelete.length - 1;
    i >= 0;
    i--
  ) {

    sheet.deleteRow(
      rowsToDelete[i]
    );

  }


  return rowsToDelete.length;

}


/**
 * ============================================================
 * WRITE INDEX CHUNK
 * ============================================================
 */

function writeMSERawIndexChunk_(
  rows
) {

  if (
    !rows ||
    !rows.length
  ) {

    return;

  }


  const sheet =
    getMSERawIndexSheet_();


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      11
    )
    .setValues(
      rows
    );

}


/**
 * ============================================================
 * PROCESS ONE RAW FILE
 *
 * MAIN INGESTION FUNCTION
 * ============================================================
 */

function processMSERawFile(
  rawFileId
) {

  const startTime =
    new Date();


  const catalog =
    findMSERawCatalogByRawFileId_(
      rawFileId
    );


  if (
    !catalog
  ) {

    throw new Error(
      "RAW_FILE_ID not found:\n" +
      rawFileId
    );

  }


  const catalogData =
    catalog.data;


  const sourceFileId =
    String(
      catalogData[2] || ""
    ).trim();


  if (
    !sourceFileId
  ) {

    throw new Error(
      "RAW catalog has no FILE_ID for:\n" +
      rawFileId
    );

  }


  /**
   * ----------------------------------------------------------
   * Mark processing
   * ----------------------------------------------------------
   */

  updateMSERawCatalogStatus_(
    rawFileId,
    "PROCESSING",
    {
      processedTimestamp:
        "",
      errorMessage:
        ""
    }
  );


  let temporarySheetId =
    null;


  try {

    /**
     * --------------------------------------------------------
     * Convert XLSX → Google Sheet
     * --------------------------------------------------------
     */

    const converted =
      convertMSERawExcelToGoogleSheet_(
        sourceFileId
      );


    temporarySheetId =
      converted.id;


    /**
     * --------------------------------------------------------
     * Open temporary spreadsheet
     * --------------------------------------------------------
     */

    const temporarySS =
      SpreadsheetApp.openById(
        temporarySheetId
      );


    const sourceSheet =
      getMSERawSourceSheet_(
        temporarySS
      );


    const headers =
      getMSERawSourceHeaders_(
        sourceSheet
      );


    const columns =
      detectMSERawColumns_(
        headers
      );


    Logger.log(
      "Detected RAW columns:\n" +
      JSON.stringify(
        columns,
        null,
        2
      )
    );


    /**
     * --------------------------------------------------------
     * Source dimensions
     * --------------------------------------------------------
     */

    const lastRow =
      sourceSheet.getLastRow();


    const lastColumn =
      sourceSheet.getLastColumn();


    if (
      lastRow < 2
    ) {

      throw new Error(
        "RAW source contains no transaction rows."
      );

    }


    const sourceRowCount =
      lastRow - 1;


    /**
     * --------------------------------------------------------
     * Clear previous index
     *
     * Safe re-processing.
     * --------------------------------------------------------
     */

    const deletedRows =
      clearMSERawIndexForFile_(
        rawFileId
      );


    Logger.log(
      "Previous index rows removed: " +
      deletedRows
    );


    /**
     * --------------------------------------------------------
     * Aggregation
     *
     * We process the source in chunks.
     * --------------------------------------------------------
     */

    let transactionMap =
      new Map();


    let totalSourceRows =
      0;


    let invalidRows =
      0;


    let validRows =
      0;


    let earliestDate =
      "";


    let latestDate =
      "";


    let totalTransactionsWritten =
      0;


    let outputBuffer =
      [];


    /**
     * --------------------------------------------------------
     * Flush function
     * --------------------------------------------------------
     */

    const flushTransactions =
      () => {

        if (
          transactionMap.size ===
          0
        ) {

          return;

        }


        transactionMap
          .forEach(
            transaction => {

              finalizeMSERawTransaction_(
                transaction
              );


              const cardNumbers =
                getMSERawCardArray_(
                  transaction.cardNumbers
                );


              outputBuffer.push([

                rawFileId,

                totalTransactionsWritten +
                outputBuffer.length +
                1,

                cardNumbers.join(
                  " | "
                ),

                transaction.transactionDate,

                transaction.registerNo,

                transaction.transactionNo,

                transaction.storeCode,

                serializeMSERawTenderBreakdown_(
                  transaction.tenderBreakdown
                ),

                transaction.gross,

                transaction.w8,

                transaction.net

              ]);

            }
          );


        /**
         * Write buffer in manageable chunks.
         */
        while (
          outputBuffer.length >=
          MSE_RAW_INGESTION_CONFIG
            .WRITE_CHUNK_SIZE
        ) {

          const writeRows =
            outputBuffer.splice(
              0,
              MSE_RAW_INGESTION_CONFIG
                .WRITE_CHUNK_SIZE
            );


          writeMSERawIndexChunk_(
            writeRows
          );


          totalTransactionsWritten +=
            writeRows.length;

        }


        /**
         * Keep only unwritten rows
         * in the buffer.
         */
        transactionMap =
          new Map();

      };


    /**
     * --------------------------------------------------------
     * Read source in chunks
     * --------------------------------------------------------
     */

    for (
      let startRow = 2;
      startRow <= lastRow;
      startRow +=
        MSE_RAW_INGESTION_CONFIG
          .CHUNK_SIZE
    ) {

      const rowsToRead =
        Math.min(
          MSE_RAW_INGESTION_CONFIG
            .CHUNK_SIZE,

          lastRow -
          startRow +
          1
        );


      const values =
        sourceSheet
          .getRange(
            startRow,
            1,
            rowsToRead,
            lastColumn
          )
          .getValues();


      for (
        let i = 0;
        i < values.length;
        i++
      ) {

        const row =
          values[i];


        totalSourceRows++;


        /**
         * ----------------------------------------------------
         * Extract values
         * ----------------------------------------------------
         */

        const card =
          normalizeMSERawCard_(
            row[
              columns.card
            ]
          );


        const date =
          normalizeMSERawDate_(
            row[
              columns.date
            ]
          );


        const register =
          normalizeMSERawCode_(
            row[
              columns.register
            ]
          );


        const trans =
          normalizeMSERawCode_(
            row[
              columns.trans
            ]
          );


        const store =
          normalizeMSERawCode_(
            row[
              columns.store
            ]
          );


        const tenderType =
          String(
            row[
              columns.tenderType
            ] || ""
          )
            .trim()
            .toUpperCase();


        const amount =
          normalizeMSERawAmount_(
            row[
              columns.tenderAmount
            ]
          );


        /**
         * ----------------------------------------------------
         * Validate minimum transaction key
         * ----------------------------------------------------
         */

        if (
          !date ||
          !store ||
          !register ||
          !trans
        ) {

          invalidRows++;

          continue;

        }


        validRows++;


        /**
         * ----------------------------------------------------
         * Period tracking
         * ----------------------------------------------------
         */

        if (
          !earliestDate ||
          date <
            earliestDate
        ) {

          earliestDate =
            date;

        }


        if (
          !latestDate ||
          date >
            latestDate
        ) {

          latestDate =
            date;

        }


        /**
         * ----------------------------------------------------
         * Build transaction key
         * ----------------------------------------------------
         */

        const key =
          buildMSERawTransactionKey_(
            date,
            store,
            register,
            trans
          );


        let transaction =
          transactionMap.get(
            key
          );


        if (
          !transaction
        ) {

          transaction =
            createMSERawTransaction_(
              date,
              store,
              register,
              trans
            );


          transactionMap.set(
            key,
            transaction
          );

        }


        /**
         * ----------------------------------------------------
         * Add tender row
         * ----------------------------------------------------
         */

        addMSERawTenderRow_(
          transaction,
          card,
          tenderType,
          amount
        );


        /**
         * ----------------------------------------------------
         * Prevent unlimited memory growth.
         *
         * When transaction count becomes large,
         * flush the current aggregation.
         *
         * NOTE:
         * This assumes source rows belonging to the same
         * transaction are normally adjacent, as they are
         * in the system extraction.
         * ----------------------------------------------------
         */

        if (
          transactionMap.size >=
          MSE_RAW_INGESTION_CONFIG
            .MAX_MEMORY_TRANSACTIONS
        ) {

          flushTransactions();

        }

      }


      /**
       * ------------------------------------------------------
       * Progress logging
       * ------------------------------------------------------
       */

      if (
        totalSourceRows %
          50000 ===
        0
      ) {

        Logger.log(
          "RAW INGESTION PROGRESS: " +
          totalSourceRows +
          " / " +
          sourceRowCount
        );

      }

    }


    /**
     * --------------------------------------------------------
     * Flush remaining transactions
     * --------------------------------------------------------
     */

    flushTransactions();


    /**
     * --------------------------------------------------------
     * Write remaining output buffer
     * --------------------------------------------------------
     */

    if (
      outputBuffer.length
    ) {

      writeMSERawIndexChunk_(
        outputBuffer
      );


      totalTransactionsWritten +=
        outputBuffer.length;


      outputBuffer = [];

    }


    /**
     * --------------------------------------------------------
     * Determine status
     * --------------------------------------------------------
     */

    const processedTimestamp =
      new Date();


    updateMSERawCatalogStatus_(
      rawFileId,
      "ACTIVE",
      {

        totalRows:
          totalSourceRows,

        periodFrom:
          earliestDate,

        periodTo:
          latestDate,

        processedTimestamp:
          processedTimestamp,

        errorMessage:
          "",

        active:
          true,

        notes:
          "RAW Excel successfully indexed. " +
          "Valid source rows: " +
          validRows +
          ". Invalid source rows: " +
          invalidRows +
          "."

      }
    );


    /**
     * --------------------------------------------------------
     * Delete temporary conversion
     * --------------------------------------------------------
     */

    deleteMSERawTemporarySheet_(
      temporarySheetId
    );


    const duration =
      (
        (new Date() - startTime) /
        1000
      ).toFixed(2);


    const result = {

      success:
        true,

      rawFileId:
        rawFileId,

      sourceFileId:
        sourceFileId,

      sourceRows:
        totalSourceRows,

      validRows:
        validRows,

      invalidRows:
        invalidRows,

      indexedTransactions:
        totalTransactionsWritten,

      periodFrom:
        earliestDate,

      periodTo:
        latestDate,

      status:
        "ACTIVE",

      durationSeconds:
        Number(duration)

    };


    Logger.log(
      JSON.stringify(
        result,
        null,
        2
      )
    );


    return result;

  } catch (error) {

    /**
     * --------------------------------------------------------
     * Processing failure
     * --------------------------------------------------------
     */

    const message =
      error &&
      error.message
        ? error.message
        : String(error);


    try {

      updateMSERawCatalogStatus_(
        rawFileId,
        "ERROR",
        {

          processedTimestamp:
            new Date(),

          errorMessage:
            message,

          active:
            false

        }
      );

    } catch (
      catalogError
    ) {

      Logger.log(
        "Unable to update RAW catalog after failure: " +
        catalogError.message
      );

    }


    /**
     * Delete temporary conversion
     */
    deleteMSERawTemporarySheet_(
      temporarySheetId
    );


    throw new Error(
      "RAW ingestion failed.\n\n" +
      "RAW FILE ID:\n" +
      rawFileId +
      "\n\n" +
      message
    );

  }

}


/**
 * ============================================================
 * PROCESS REGISTERED RAW FILE
 * ============================================================
 *
 * This is the function you will use for the current sample.
 *
 * ============================================================
 */
function processRegisteredMSERawFile() {

  const rawFileId =
    "RAW-20260819080335-3092F15B";


  const result =
    processMSERawFile(
      rawFileId
    );


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/**
 * ============================================================
 * GET RAW INDEX TRANSACTION
 * ============================================================
 *
 * Lookup by:
 *
 * DATE | STORE | REGISTER | TRANS
 *
 * ============================================================
 */
function lookupMSERawTransaction(
  transactionDate,
  storeCode,
  registerNo,
  transactionNo
) {

  const date =
    normalizeMSERawDate_(
      transactionDate
    );


  const store =
    normalizeMSERawCode_(
      storeCode
    );


  const register =
    normalizeMSERawCode_(
      registerNo
    );


  const trans =
    normalizeMSERawCode_(
      transactionNo
    );


  const lookupKey =
    buildMSERawTransactionKey_(
      date,
      store,
      register,
      trans
    );


  const sheet =
    getMSERawIndexSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return {

      found:
        false,

      lookupKey:
        lookupKey

    };

  }


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        11
      )
      .getValues();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const rowKey =
      buildMSERawTransactionKey_(
        normalizeMSERawDate_(
          row[3]
        ),
        normalizeMSERawCode_(
          row[6]
        ),
        normalizeMSERawCode_(
          row[4]
        ),
        normalizeMSERawCode_(
          row[5]
        )
      );


    if (
      rowKey ===
      lookupKey
    ) {

      return {

        found:
          true,

        lookupKey:
          lookupKey,

        rawFileId:
          row[0],

        rawRowId:
          row[1],

        cardNumbers:
          row[2]
            ? String(
                row[2]
              )
                .split(
                  " | "
                )
            : [],

        transactionDate:
          normalizeMSERawDate_(
            row[3]
          ),

        registerNo:
          normalizeMSERawCode_(
            row[4]
          ),

        transactionNo:
          normalizeMSERawCode_(
            row[5]
          ),

        storeCode:
          normalizeMSERawCode_(
            row[6]
          ),

        tenderBreakdown:
          row[7],

        gross:
          Number(
            row[8]
          ) || 0,

        w8:
          Number(
            row[9]
          ) || 0,

        net:
          Number(
            row[10]
          ) || 0

      };

    }

  }


  return {

    found:
      false,

    lookupKey:
      lookupKey

  };

}


/**
 * ============================================================
 * TEST RAW TRANSACTION LOOKUP
 * ============================================================
 *
 * Uses your previously supplied example:
 *
 * 20260430 | 343 | 14 | 5942
 *
 * ============================================================
 */
function testMSERawTransactionLookup() {

  const result =
    lookupMSERawTransaction(
      "20260430",
      "343",
      "14",
      "5942"
    );


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}



function testMSEExcelConversion() {

  const excelFileId =
    "1baAWMos5Vw7gSKyGTOxL_mRojvbYFtoB";

  const sourceFile =
    DriveApp.getFileById(excelFileId);

  Logger.log(
    "SOURCE FILE: " +
    sourceFile.getName()
  );

  Logger.log(
    "SOURCE MIME TYPE: " +
    sourceFile.getMimeType()
  );

  const blob =
    sourceFile.getBlob();

  const metadata = {

    name:
      "[TEMP] " +
      sourceFile.getName(),

    mimeType:
      MimeType.GOOGLE_SHEETS

  };

  const convertedFile =
    Drive.Files.create(
      metadata,
      blob,
      {
        fields:
          "id,name,mimeType,size"
      }
    );

  Logger.log(
    "CONVERTED FILE:"
  );

  Logger.log(
    JSON.stringify(
      convertedFile,
      null,
      2
    )
  );

  return convertedFile;
}



function inspectConvertedMSEFile() {

  const spreadsheetId =
    "1UJUlI6r3dK-1vcIaMEiSPES3iBCqFhnf19SDUbN0UEw";

  const ss =
    SpreadsheetApp.openById(
      spreadsheetId
    );

  Logger.log(
    "TEMP SPREADSHEET: " +
    ss.getName()
  );

  const sheets =
    ss.getSheets();

  Logger.log(
    "SHEETS FOUND: " +
    sheets.length
  );

  sheets.forEach(
    function(sheet, index) {

      Logger.log(
        "SHEET " +
        (index + 1) +
        ": " +
        sheet.getName()
      );

      Logger.log(
        "ROWS: " +
        sheet.getLastRow()
      );

      Logger.log(
        "COLUMNS: " +
        sheet.getLastColumn()
      );

    }
  );


  const firstSheet =
    sheets[0];

  const rows =
    Math.min(
      firstSheet.getLastRow(),
      10
    );

  const columns =
    firstSheet.getLastColumn();


  if (
    rows === 0 ||
    columns === 0
  ) {

    throw new Error(
      "Converted RAW spreadsheet contains no data."
    );

  }


  const sample =
    firstSheet
      .getRange(
        1,
        1,
        rows,
        columns
      )
      .getDisplayValues();


  Logger.log(
    "RAW SAMPLE:"
  );

  Logger.log(
    JSON.stringify(
      sample,
      null,
      2
    )
  );


  return {

    success: true,

    spreadsheetId:
      spreadsheetId,

    spreadsheetName:
      ss.getName(),

    sheetName:
      firstSheet.getName(),

    rows:
      firstSheet.getLastRow(),

    columns:
      firstSheet.getLastColumn(),

    sample:
      sample

  };

}


/**
 * ============================================================
 * MSE RAW FILE INGESTION ENGINE
 * ============================================================
 *
 * PURPOSE:
 *   Convert registered Excel RAW files into indexed data.
 *
 * DESIGN:
 *
 *   ORIGINAL EXCEL
 *        ↓
 *   TEMP GOOGLE SHEET
 *        ↓
 *   CHUNKED READ
 *        ↓
 *   RAW INDEX SHEET FOR THAT FILE
 *        ↓
 *   RAW_FILE_CATALOG = PROCESSED
 *
 * IMPORTANT:
 *   Each RAW file gets its own index sheet.
 *
 *   This prevents millions of rows from different RAW files
 *   being forced into one Google Sheet.
 *
 * ============================================================
 */


/**
 * ============================================================
 * CONFIGURATION
 * ============================================================
 */

const MSE_INGEST_CONFIG = {

  // Number of RAW rows processed per execution chunk.
  //
  // 20,000 is intentionally conservative because each row
  // contains only 7 source columns but we also generate
  // normalized values.
  CHUNK_SIZE: 20000,

  // Maximum execution time we allow ourselves before stopping.
  //
  // Apps Script execution limits vary by account/runtime.
  // Leaving a safety margin allows us to save the progress
  // before the execution is forcibly terminated.
  MAX_RUNTIME_MS: 5 * 60 * 1000,

  // Temporary converted spreadsheets are created here.
  TEMP_NAME_PREFIX: "[TEMP] ",

  // Prefix used for individual RAW index sheets.
  INDEX_SHEET_PREFIX: "RAW_IDX_"

};


/**
 * ============================================================
 * MAIN FUNCTION
 * ============================================================
 *
 * Run this function after a RAW file has been REGISTERED.
 *
 * It automatically finds the next RAW file that needs
 * ingestion and processes it.
 *
 * ============================================================
 */

function MSE_INGEST_ProcessNextRawFile() {

  const startTime = Date.now();

  const ss =
    getMSEDatabase_();

  const catalog =
    ss.getSheetByName(
      "RAW_FILE_CATALOG"
    );

  if (!catalog) {

    throw new Error(
      "RAW_FILE_CATALOG sheet does not exist."
    );

  }


  const catalogData =
    catalog
      .getDataRange()
      .getValues();


  if (
    catalogData.length <= 1
  ) {

    Logger.log(
      "No RAW files found."
    );

    return {
      success: true,
      message: "No RAW files found."
    };

  }


  const headers =
    catalogData[0];


  const col =
    MSE_INGEST_GetColumnMap_(headers);


  /**
   * ----------------------------------------------------------
   * Find next REGISTERED / PROCESSING file
   * ----------------------------------------------------------
   */

  let targetRow = null;

  for (
    let r = 1;
    r < catalogData.length;
    r++
  ) {

    const status =
      String(
        catalogData[r][col.STATUS] || ""
      )
      .trim()
      .toUpperCase();


    if (
      status === "REGISTERED" ||
      status === "PROCESSING"
    ) {

      targetRow =
        r + 1;

      break;

    }

  }


  if (!targetRow) {

    Logger.log(
      "No RAW files require ingestion."
    );

    return {
      success: true,
      message:
        "No RAW files require ingestion."
    };

  }


  const result =
    MSE_INGEST_ProcessCatalogRow_(
      ss,
      catalog,
      targetRow,
      col,
      startTime
    );


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/**
 * ============================================================
 * PROCESS ONE CATALOG FILE
 * ============================================================
 */

function MSE_INGEST_ProcessCatalogRow_(
  ss,
  catalog,
  catalogRow,
  col,
  startTime
) {

  const rowValues =
    catalog
      .getRange(
        catalogRow,
        1,
        1,
        catalog.getLastColumn()
      )
      .getValues()[0];


  const rawFileId =
    String(
      rowValues[col.RAW_FILE_ID] || ""
    ).trim();


  const fileId =
    String(
      rowValues[col.FILE_ID] || ""
    ).trim();


  const fileName =
    String(
      rowValues[col.FILE_NAME] || ""
    ).trim();


  if (!rawFileId) {

    throw new Error(
      "RAW_FILE_ID is missing in catalog row " +
      catalogRow
    );

  }


  if (!fileId) {

    throw new Error(
      "FILE_ID is missing for RAW file " +
      rawFileId
    );

  }


  Logger.log(
    "=================================================="
  );

  Logger.log(
    "PROCESSING RAW FILE"
  );

  Logger.log(
    "RAW FILE ID: " +
    rawFileId
  );

  Logger.log(
    "FILE: " +
    fileName
  );

  Logger.log(
    "DRIVE FILE ID: " +
    fileId
  );

  Logger.log(
    "=================================================="
  );


  /**
   * ----------------------------------------------------------
   * Mark catalog as PROCESSING
   * ----------------------------------------------------------
   */

  catalog
    .getRange(
      catalogRow,
      col.STATUS + 1
    )
    .setValue(
      "PROCESSING"
    );


  /**
   * ----------------------------------------------------------
   * Create / locate target index sheet
   * ----------------------------------------------------------
   */

  const indexSheet =
    MSE_INGEST_GetOrCreateIndexSheet_(
      ss,
      rawFileId
    );


  /**
   * ----------------------------------------------------------
   * Check existing progress
   * ----------------------------------------------------------
   */

  let processedRows =
    MSE_INGEST_GetProcessedRows_(
      indexSheet
    );


  Logger.log(
    "Existing processed rows: " +
    processedRows
  );


  /**
   * ----------------------------------------------------------
   * Convert Excel to temporary Google Sheet
   * ----------------------------------------------------------
   */

  const tempInfo =
    MSE_INGEST_GetOrCreateTemporarySheet_(
      fileId,
      rawFileId
    );


  Logger.log(
    "TEMP FILE ID: " +
    tempInfo.id
  );


  const tempSS =
    SpreadsheetApp.openById(
      tempInfo.id
    );


  const sourceSheet =
    tempSS.getSheets()[0];


  const totalRows =
    sourceSheet.getLastRow();


  const totalColumns =
    sourceSheet.getLastColumn();


  Logger.log(
    "SOURCE ROWS: " +
    totalRows
  );

  Logger.log(
    "SOURCE COLUMNS: " +
    totalColumns
  );


  if (totalColumns < 7) {

    throw new Error(
      "RAW file has only " +
      totalColumns +
      " columns. Expected at least 7."
    );

  }


  /**
   * ----------------------------------------------------------
   * Validate header
   * ----------------------------------------------------------
   */

  MSE_INGEST_ValidateHeader_(
    sourceSheet
  );


  /**
   * ----------------------------------------------------------
   * Store total row count
   * ----------------------------------------------------------
   */

  if (
    col.TOTAL_ROWS !== undefined
  ) {

    catalog
      .getRange(
        catalogRow,
        col.TOTAL_ROWS + 1
      )
      .setValue(
        totalRows - 1
      );

  }


  /**
   * ----------------------------------------------------------
   * Calculate data rows
   * ----------------------------------------------------------
   */

  const totalDataRows =
    Math.max(
      0,
      totalRows - 1
    );


  /**
   * ----------------------------------------------------------
   * Nothing to process
   * ----------------------------------------------------------
   */

  if (
    totalDataRows === 0
  ) {

    MSE_INGEST_FinalizeRawFile_(
      catalog,
      catalogRow,
      col,
      tempInfo.id
    );

    return {
      success: true,
      rawFileId: rawFileId,
      status: "PROCESSED",
      rows: 0
    };

  }


  /**
   * ----------------------------------------------------------
   * Determine chunk
   * ----------------------------------------------------------
   */

  const remainingRows =
    totalDataRows -
    processedRows;


  if (
    remainingRows <= 0
  ) {

    MSE_INGEST_FinalizeRawFile_(
      catalog,
      catalogRow,
      col,
      tempInfo.id
    );

    return {
      success: true,
      rawFileId: rawFileId,
      status: "PROCESSED",
      rows: processedRows
    };

  }


  const chunkSize =
    Math.min(
      MSE_INGEST_CONFIG.CHUNK_SIZE,
      remainingRows
    );


  const startSourceRow =
    2 +
    processedRows;


  Logger.log(
    "READING SOURCE ROW: " +
    startSourceRow
  );

  Logger.log(
    "CHUNK SIZE: " +
    chunkSize
  );


  /**
   * ----------------------------------------------------------
   * Read chunk
   * ----------------------------------------------------------
   */

  const values =
    sourceSheet
      .getRange(
        startSourceRow,
        1,
        chunkSize,
        7
      )
      .getDisplayValues();


  /**
   * ----------------------------------------------------------
   * Convert RAW rows
   * ----------------------------------------------------------
   */

  const output =
    [];


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const sourceRow =
      startSourceRow +
      i;


    const raw =
      values[i];


    const parsed =
      MSE_INGEST_ParseRawRow_(
        raw
      );


    if (!parsed.valid) {

      /**
       * Keep invalid rows in the index.
       *
       * This is important for audit purposes.
       */

      output.push([

        rawFileId,

        sourceRow,

        parsed.cardNumber,

        parsed.transactionDate,

        parsed.dateKey,

        parsed.registerNo,

        parsed.transactionNo,

        parsed.storeCode,

        parsed.tenderType,

        parsed.tenderAmount,

        parsed.lookupKey,

        parsed.status

      ]);

    } else {

      output.push([

        rawFileId,

        sourceRow,

        parsed.cardNumber,

        parsed.transactionDate,

        parsed.dateKey,

        parsed.registerNo,

        parsed.transactionNo,

        parsed.storeCode,

        parsed.tenderType,

        parsed.tenderAmount,

        parsed.lookupKey,

        "VALID"

      ]);

    }

  }


  /**
   * ----------------------------------------------------------
   * Write chunk
   * ----------------------------------------------------------
   */

  MSE_INGEST_AppendRows_(
    indexSheet,
    output
  );


  processedRows +=
    values.length;


  /**
   * ----------------------------------------------------------
   * Save progress
   * ----------------------------------------------------------
   */

  catalog
    .getRange(
      catalogRow,
      col.STATUS + 1
    )
    .setValue(
      processedRows >= totalDataRows
        ? "PROCESSED"
        : "PROCESSING"
    );


  /**
   * ----------------------------------------------------------
   * Determine whether finished
   * ----------------------------------------------------------
   */

  const finished =
    processedRows >= totalDataRows;


  if (finished) {

    MSE_INGEST_FinalizeRawFile_(
      catalog,
      catalogRow,
      col,
      tempInfo.id
    );

  }


  const duration =
    (
      Date.now() -
      startTime
    ) / 1000;


  const result = {

    success: true,

    rawFileId:
      rawFileId,

    fileName:
      fileName,

    processedThisRun:
      values.length,

    processedTotal:
      processedRows,

    totalRows:
      totalDataRows,

    remainingRows:
      Math.max(
        0,
        totalDataRows -
        processedRows
      ),

    status:
      finished
        ? "PROCESSED"
        : "PROCESSING",

    durationSeconds:
      Number(
        duration.toFixed(2)
      )

  };


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/**
 * ============================================================
 * CREATE / GET RAW INDEX SHEET
 * ============================================================
 */

function MSE_INGEST_GetOrCreateIndexSheet_(
  ss,
  rawFileId
) {

  /**
   * Keep the sheet name below Google's 100-character limit.
   */

  const safeId =
    rawFileId
      .replace(
        /[^A-Za-z0-9_-]/g,
        ""
      )
      .slice(
        -30
      );


  const sheetName =
    MSE_INGEST_CONFIG.INDEX_SHEET_PREFIX +
    safeId;


  let sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        sheetName
      );


    const headers = [

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


    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);


    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setFontWeight(
        "bold"
      );


    sheet.setFrozenRows(
      1
    );


    Logger.log(
      "Created index sheet: " +
      sheetName
    );

  }


  return sheet;

}


/**
 * ============================================================
 * GET PROCESSED ROW COUNT
 * ============================================================
 */

function MSE_INGEST_GetProcessedRows_(
  sheet
) {

  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return 0;

  }


  /**
   * Since the index contains exactly one output row
   * for every RAW input row, this is our checkpoint.
   */

  return lastRow - 1;

}


/**
 * ============================================================
 * CREATE TEMPORARY GOOGLE SHEET
 * ============================================================
 */

function MSE_INGEST_GetOrCreateTemporarySheet_(
  fileId,
  rawFileId
) {

  /**
   * Search Drive for an existing temporary file first.
   *
   * This allows a processing execution to resume without
   * converting the Excel file again.
   */

  const expectedName =
    MSE_INGEST_CONFIG.TEMP_NAME_PREFIX +
    rawFileId;


  const existing =
    DriveApp
      .searchFiles(
        "title = '" +
        expectedName.replace(
          /'/g,
          "\\'"
        ) +
        "' and trashed = false"
      );


  if (
    existing.hasNext()
  ) {

    const file =
      existing.next();


    Logger.log(
      "Existing temporary spreadsheet found."
    );


    return {

      id:
        file.getId(),

      name:
        file.getName()

    };

  }


  /**
   * ----------------------------------------------------------
   * Get original Excel
   * ----------------------------------------------------------
   */

  const sourceFile =
    DriveApp.getFileById(
      fileId
    );


  const blob =
    sourceFile.getBlob();


  /**
   * ----------------------------------------------------------
   * Convert Excel → Google Sheets
   * ----------------------------------------------------------
   */

  const metadata = {

    name:
      expectedName,

    mimeType:
      MimeType.GOOGLE_SHEETS

  };


  const converted =
    Drive.Files.create(
      metadata,
      blob,
      {
        fields:
          "id,name,mimeType"
      }
    );


  Logger.log(
    "Created temporary spreadsheet: " +
    converted.id
  );


  return {

    id:
      converted.id,

    name:
      converted.name

  };

}


/**
 * ============================================================
 * VALIDATE RAW HEADER
 * ============================================================
 */

function MSE_INGEST_ValidateHeader_(
  sheet
) {

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        7
      )
      .getDisplayValues()[0]
      .map(
        function(value) {

          return String(
            value
          )
          .trim()
          .toUpperCase();

        }
      );


  const expected = [

    "CARD NUMBER",
    "DATE",
    "REGISTER NUM",
    "TRANS NUM",
    "STORE CODE",
    "TENDER TYPE",
    "TENDER AMOUNT"

  ];


  for (
    let i = 0;
    i < expected.length;
    i++
  ) {

    if (
      headers[i] !== expected[i]
    ) {

      throw new Error(

        "RAW header mismatch at column " +
        (i + 1) +

        ". Expected '" +
        expected[i] +

        "' but found '" +
        headers[i] +
        "'."

      );

    }

  }


  Logger.log(
    "RAW HEADER VALIDATED."
  );

}


/**
 * ============================================================
 * PARSE RAW ROW
 * ============================================================
 */

function MSE_INGEST_ParseRawRow_(
  row
) {

  const cardNumber =
    MSE_INGEST_NormalizeCard_(
      row[0]
    );


  const transactionDate =
    MSE_INGEST_NormalizeDate_(
      row[1]
    );


  const registerNo =
    MSE_INGEST_NormalizeInteger_(
      row[2]
    );


  const transactionNo =
    MSE_INGEST_NormalizeInteger_(
      row[3]
    );


  const storeCode =
    MSE_INGEST_NormalizeStoreCode_(
      row[4]
    );


  const tenderType =
    String(
      row[5] || ""
    )
    .trim()
    .toUpperCase();


  const tenderAmount =
    MSE_INGEST_NormalizeAmount_(
      row[6]
    );


  const dateKey =
    transactionDate;


  const lookupKey =

    dateKey +
    "|" +
    storeCode +
    "|" +
    registerNo +
    "|" +
    transactionNo;


  let valid = true;

  let status = "VALID";


  if (!dateKey) {

    valid = false;
    status = "INVALID_DATE";

  }


  if (!storeCode) {

    valid = false;
    status = "INVALID_STORE_CODE";

  }


  if (!registerNo) {

    valid = false;
    status = "INVALID_REGISTER";

  }


  if (!transactionNo) {

    valid = false;
    status = "INVALID_TRANSACTION";

  }


  if (!tenderType) {

    valid = false;
    status = "INVALID_TENDER_TYPE";

  }


  if (
    tenderAmount === null ||
    isNaN(tenderAmount)
  ) {

    valid = false;
    status = "INVALID_AMOUNT";

  }


  return {

    valid:
      valid,

    status:
      status,

    cardNumber:
      cardNumber,

    transactionDate:
      transactionDate,

    dateKey:
      dateKey,

    registerNo:
      registerNo,

    transactionNo:
      transactionNo,

    storeCode:
      storeCode,

    tenderType:
      tenderType,

    tenderAmount:
      tenderAmount,

    lookupKey:
      lookupKey

  };

}


/**
 * ============================================================
 * CARD NORMALIZATION
 * ============================================================
 */

function MSE_INGEST_NormalizeCard_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return String(
    value
  )
  .replace(
    /\D/g,
    ""
  );

}


/**
 * ============================================================
 * DATE NORMALIZATION
 * ============================================================
 *
 * RAW examples:
 *
 * 20251031
 * 20260430
 *
 * Result:
 *
 * 20251031
 * 20260430
 *
 * We deliberately keep DATE_KEY as YYYYMMDD because it is
 * ideal for transaction lookup.
 *
 * ============================================================
 */

function MSE_INGEST_NormalizeDate_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  let text =
    String(
      value
    )
    .trim();


  /**
   * Remove separators.
   */

  text =
    text.replace(
      /[^0-9]/g,
      ""
    );


  /**
   * YYYYMMDD
   */

  if (
    /^\d{8}$/.test(
      text
    )
  ) {

    const year =
      Number(
        text.substring(
          0,
          4
        )
      );


    const month =
      Number(
        text.substring(
          4,
          6
        )
      );


    const day =
      Number(
        text.substring(
          6,
          8
        )
      );


    if (
      year >= 1900 &&
      month >= 1 &&
      month <= 12 &&
      day >= 1 &&
      day <= 31
    ) {

      return text;

    }

  }


  /**
   * If a real Date somehow reaches this function.
   */

  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]"
  ) {

    if (
      !isNaN(
        value.getTime()
      )
    ) {

      return Utilities.formatDate(
        value,
        Session.getScriptTimeZone(),
        "yyyyMMdd"
      );

    }

  }


  return "";

}


/**
 * ============================================================
 * INTEGER NORMALIZATION
 * ============================================================
 */

function MSE_INGEST_NormalizeInteger_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  const text =
    String(
      value
    )
    .trim()
    .replace(
      /,/g,
      ""
    );


  if (
    text === ""
  ) {

    return "";

  }


  const number =
    Number(
      text
    );


  if (
    !isFinite(
      number
    )
  ) {

    return "";

  }


  return String(
    Math.trunc(
      number
    )
  );

}


/**
 * ============================================================
 * STORE CODE NORMALIZATION
 * ============================================================
 */

function MSE_INGEST_NormalizeStoreCode_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  const text =
    String(
      value
    )
    .trim()
    .replace(
      /,/g,
      ""
    );


  if (
    text === ""
  ) {

    return "";

  }


  const number =
    Number(
      text
    );


  if (
    isFinite(
      number
    )
  ) {

    return String(
      Math.trunc(
        number
      )
    );

  }


  return text;

}


/**
 * ============================================================
 * AMOUNT NORMALIZATION
 * ============================================================
 *
 * Example:
 *
 * 1,987.28 → 1987.28
 * 599.00   → 599
 *
 * ============================================================
 */

function MSE_INGEST_NormalizeAmount_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return null;

  }


  let text =
    String(
      value
    )
    .trim();


  if (
    text === ""
  ) {

    return null;

  }


  text =
    text.replace(
      /,/g,
      ""
    );


  text =
    text.replace(
      /[^0-9.\-]/g,
      ""
    );


  const amount =
    Number(
      text
    );


  if (
    !isFinite(
      amount
    )
  ) {

    return null;

  }


  return amount;

}


/**
 * ============================================================
 * APPEND OUTPUT ROWS
 * ============================================================
 */

function MSE_INGEST_AppendRows_(
  sheet,
  rows
) {

  if (
    !rows ||
    rows.length === 0
  ) {

    return;

  }


  const startRow =
    Math.max(
      2,
      sheet.getLastRow() + 1
    );


  const requiredRows =
    startRow +
    rows.length -
    1;


  const maxRows =
    sheet.getMaxRows();


  if (
    requiredRows >
    maxRows
  ) {

    sheet.insertRowsAfter(
      maxRows,
      requiredRows -
      maxRows
    );

  }


  sheet
    .getRange(
      startRow,
      1,
      rows.length,
      rows[0].length
    )
    .setValues(
      rows
    );

}


/**
 * ============================================================
 * FINALIZE RAW FILE
 * ============================================================
 */

function MSE_INGEST_FinalizeRawFile_(
  catalog,
  catalogRow,
  col,
  tempSpreadsheetId
) {

  /**
   * ----------------------------------------------------------
   * STATUS
   * ----------------------------------------------------------
   */

  catalog
    .getRange(
      catalogRow,
      col.STATUS + 1
    )
    .setValue(
      "PROCESSED"
    );


  /**
   * ----------------------------------------------------------
   * PROCESSED TIMESTAMP
   * ----------------------------------------------------------
   */

  if (
    col.PROCESSED_TIMESTAMP !== undefined
  ) {

    catalog
      .getRange(
        catalogRow,
        col.PROCESSED_TIMESTAMP + 1
      )
      .setValue(
        new Date()
      );

  }


  /**
   * ----------------------------------------------------------
   * ACTIVE
   * ----------------------------------------------------------
   */

  if (
    col.ACTIVE !== undefined
  ) {

    catalog
      .getRange(
        catalogRow,
        col.ACTIVE + 1
      )
      .setValue(
        true
      );

  }


  /**
   * ----------------------------------------------------------
   * Delete temporary Google Sheet
   * ----------------------------------------------------------
   */

  if (
    tempSpreadsheetId
  ) {

    try {

      DriveApp
        .getFileById(
          tempSpreadsheetId
        )
        .setTrashed(
          true
        );


      Logger.log(
        "Temporary spreadsheet deleted."
      );

    }
    catch (error) {

      Logger.log(
        "WARNING: Could not trash temporary spreadsheet: " +
        error.message
      );

    }

  }


  SpreadsheetApp.flush();


  Logger.log(
    "RAW FILE INGESTION COMPLETE."
  );

}


/**
 * ============================================================
 * CATALOG COLUMN MAP
 * ============================================================
 */

function MSE_INGEST_GetColumnMap_(
  headers
) {

  const map = {};


  headers.forEach(
    function(header, index) {

      const key =
        String(
          header || ""
        )
        .trim()
        .toUpperCase();


      if (key) {

        map[key] =
          index;

      }

    }
  );


  const required = [

    "RAW_FILE_ID",
    "FILE_NAME",
    "FILE_ID",
    "STATUS"

  ];


  required.forEach(
    function(key) {

      if (
        map[key] === undefined
      ) {

        throw new Error(
          "RAW_FILE_CATALOG is missing required column: " +
          key
        );

      }

    }
  );


  return map;

}


/**
 * ============================================================
 * MANUAL PROCESS SPECIFIC RAW FILE
 * ============================================================
 *
 * Useful when you have several RAW files and want to process
 * a particular one.
 *
 * ============================================================
 */

function MSE_INGEST_ProcessRawFileById(
  rawFileId
) {

  if (
    !rawFileId
  ) {

    throw new Error(
      "RAW_FILE_ID is required."
    );

  }


  const ss =
    getMSEDatabase_();


  const catalog =
    ss.getSheetByName(
      "RAW_FILE_CATALOG"
    );


  if (!catalog) {

    throw new Error(
      "RAW_FILE_CATALOG not found."
    );

  }


  const data =
    catalog
      .getDataRange()
      .getValues();


  const headers =
    data[0];


  const col =
    MSE_INGEST_GetColumnMap_(
      headers
    );


  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    if (
      String(
        data[r][col.RAW_FILE_ID] || ""
      ).trim() ===
      String(
        rawFileId
      ).trim()
    ) {

      return MSE_INGEST_ProcessCatalogRow_(
        ss,
        catalog,
        r + 1,
        col,
        Date.now()
      );

    }

  }


  throw new Error(
    "RAW_FILE_ID not found: " +
    rawFileId
  );

}


/**
 * ============================================================
 * CHECK INGESTION STATUS
 * ============================================================
 */

function MSE_INGEST_CheckRawStatus(
  rawFileId
) {

  const ss =
    getMSEDatabase_();


  const catalog =
    ss.getSheetByName(
      "RAW_FILE_CATALOG"
    );


  if (!catalog) {

    throw new Error(
      "RAW_FILE_CATALOG not found."
    );

  }


  const data =
    catalog
      .getDataRange()
      .getValues();


  const headers =
    data[0];


  const col =
    MSE_INGEST_GetColumnMap_(
      headers
    );


  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    if (
      String(
        data[r][col.RAW_FILE_ID] || ""
      ).trim() ===
      String(
        rawFileId
      ).trim()
    ) {

      return {

        found: true,

        rawFileId:
          data[r][col.RAW_FILE_ID],

        fileName:
          data[r][col.FILE_NAME],

        fileId:
          data[r][col.FILE_ID],

        status:
          data[r][col.STATUS],

        totalRows:
          col.TOTAL_ROWS !== undefined
            ? data[r][col.TOTAL_ROWS]
            : "",

        processedTimestamp:
          col.PROCESSED_TIMESTAMP !== undefined
            ? data[r][col.PROCESSED_TIMESTAMP]
            : ""

      };

    }

  }


  return {

    found: false,

    rawFileId:
      rawFileId

  };

}

/**
 * ============================================================
 * TEST: PROCESS CURRENT RAW FILE
 * ============================================================
 *
 * Apps Script Run button cannot pass arguments.
 * This wrapper allows us to run the current registered RAW
 * file directly from the Apps Script editor.
 * ============================================================
 */
function TEST_MSE_INGEST_CURRENT_RAW_FILE() {

  const RAW_FILE_ID =
    "RAW-20260819080335-3092F15B";

  Logger.log(
    "Starting RAW ingestion for: " +
    RAW_FILE_ID
  );

  const result =
    MSE_INGEST_ProcessRawFileById(
      RAW_FILE_ID
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}

/**
 * ============================================================
 * MSE RAW AUTO INGESTION CONTROLLER
 * ============================================================
 *
 * PURPOSE:
 * Automatically continue RAW ingestion without manually
 * running the processor for every 20,000-row batch.
 *
 * IMPORTANT:
 * Existing progress is determined from the RAW index sheet.
 * Therefore, already-processed rows are never intentionally
 * reread.
 *
 * ============================================================
 */


/**
 * ------------------------------------------------------------
 * AUTO INGESTION SETTINGS
 * ------------------------------------------------------------
 */
const MSE_AUTO_INGEST_CONFIG = {

  // Rows processed per batch.
  CHUNK_SIZE: 20000,

  // Stop before Apps Script execution becomes unsafe.
  // 4 minutes gives us a safety margin.
  MAX_RUNTIME_MS: 4 * 60 * 1000,

  // Number of minutes before the next automatic continuation.
  CONTINUE_AFTER_MINUTES: 1

};


/**
 * ============================================================
 * START AUTO INGESTION
 * ============================================================
 *
 * Run this function ONCE.
 *
 * It will process the next registered/processing RAW file
 * and automatically schedule continuation.
 *
 * ============================================================
 */
function MSE_START_AUTO_RAW_INGESTION() {

  const lock =
    LockService.getScriptLock();

  if (
    !lock.tryLock(5000)
  ) {

    Logger.log(
      "Another RAW ingestion process is already running."
    );

    return {
      success: false,
      message:
        "Another RAW ingestion process is already running."
    };

  }


  try {

    Logger.log(
      "=================================================="
    );

    Logger.log(
      "MSE AUTO RAW INGESTION STARTED"
    );

    Logger.log(
      "=================================================="
    );


    const result =
      MSE_AUTO_ProcessLoop_();


    Logger.log(
      JSON.stringify(
        result,
        null,
        2
      )
    );


    return result;

  }
  finally {

    lock.releaseLock();

  }

}


/**
 * ============================================================
 * AUTO PROCESS LOOP
 * ============================================================
 */
function MSE_AUTO_ProcessLoop_() {

  const startTime =
    Date.now();


  const ss =
    getMSEDatabase_();


  const catalog =
    ss.getSheetByName(
      "RAW_FILE_CATALOG"
    );


  if (!catalog) {

    throw new Error(
      "RAW_FILE_CATALOG does not exist."
    );

  }


  const data =
    catalog
      .getDataRange()
      .getValues();


  if (
    data.length <= 1
  ) {

    Logger.log(
      "No RAW files found."
    );

    return {

      success: true,

      status:
        "NO_FILES",

      message:
        "No RAW files found."

    };

  }


  const headers =
    data[0];


  const col =
    MSE_INGEST_GetColumnMap_(
      headers
    );


  /**
   * ----------------------------------------------------------
   * Find the first RAW file requiring processing.
   * ----------------------------------------------------------
   */

  let targetRow =
    null;


  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    const status =
      String(
        data[r][col.STATUS] || ""
      )
      .trim()
      .toUpperCase();


    if (
      status === "REGISTERED" ||
      status === "PROCESSING"
    ) {

      targetRow =
        r + 1;

      break;

    }

  }


  /**
   * ----------------------------------------------------------
   * Nothing left to process.
   * ----------------------------------------------------------
   */

  if (!targetRow) {

    Logger.log(
      "ALL RAW FILES ARE PROCESSED."
    );


    return {

      success: true,

      status:
        "COMPLETE",

      message:
        "All RAW files are processed."

    };

  }


  const rawFileId =
    String(
      catalog
        .getRange(
          targetRow,
          col.RAW_FILE_ID + 1
        )
        .getValue()
    )
    .trim();


  Logger.log(
    "TARGET RAW FILE: " +
    rawFileId
  );


  /**
   * ----------------------------------------------------------
   * Process batches while there is safe execution time.
   * ----------------------------------------------------------
   */

  let batches =
    0;


  let totalRowsThisRun =
    0;


  let lastResult =
    null;


  while (
    Date.now() -
    startTime <
    MSE_AUTO_INGEST_CONFIG.MAX_RUNTIME_MS
  ) {


    /**
     * --------------------------------------------------------
     * Process one chunk.
     * --------------------------------------------------------
     */

    lastResult =
      MSE_AUTO_ProcessOneChunk_(
        ss,
        catalog,
        targetRow,
        col
      );


    batches++;


    if (
      lastResult.processedThisRun
    ) {

      totalRowsThisRun +=
        lastResult.processedThisRun;

    }


    Logger.log(
      JSON.stringify(
        lastResult,
        null,
        2
      )
    );


    /**
     * --------------------------------------------------------
     * File finished?
     * --------------------------------------------------------
     */

    if (
      lastResult.status ===
      "PROCESSED"
    ) {

      Logger.log(
        "RAW FILE COMPLETED: " +
        rawFileId
      );


      /**
       * Check if another RAW file exists.
       */
      const remaining =
        MSE_AUTO_HasPendingFiles_(
          catalog,
          col
        );


      if (remaining) {

        Logger.log(
          "Another RAW file is waiting."
        );


        MSE_AUTO_ScheduleNext_();


        return {

          success: true,

          status:
            "NEXT_FILE_SCHEDULED",

          completedFile:
            rawFileId,

          batches:
            batches,

          rowsThisRun:
            totalRowsThisRun

        };

      }


      return {

        success: true,

        status:
          "COMPLETE",

        completedFile:
          rawFileId,

        batches:
          batches,

        rowsThisRun:
          totalRowsThisRun

      };

    }


    /**
     * --------------------------------------------------------
     * Check execution time.
     * --------------------------------------------------------
     */

    const elapsed =
      Date.now() -
      startTime;


    if (
      elapsed >=
      MSE_AUTO_INGEST_CONFIG.MAX_RUNTIME_MS
    ) {

      Logger.log(
        "Safe execution time reached."
      );

      break;

    }


    /**
     * --------------------------------------------------------
     * Safety:
     * Don't blindly continue if no rows were processed.
     * --------------------------------------------------------
     */

    if (
      !lastResult.processedThisRun
    ) {

      Logger.log(
        "No rows processed in latest batch."
      );

      break;

    }

  }


  /**
   * ----------------------------------------------------------
   * Schedule continuation.
   * ----------------------------------------------------------
   */

  MSE_AUTO_ScheduleNext_();


  return {

    success: true,

    status:
      "CONTINUE_SCHEDULED",

    rawFileId:
      rawFileId,

    batches:
      batches,

    rowsThisRun:
      totalRowsThisRun,

    lastResult:
      lastResult

  };

}


/**
 * ============================================================
 * PROCESS ONE CHUNK
 * ============================================================
 *
 * This function contains the actual chunk processing logic.
 * It intentionally does NOT create a new temporary file if
 * one already exists.
 *
 * ============================================================
 */
function MSE_AUTO_ProcessOneChunk_(
  ss,
  catalog,
  catalogRow,
  col
) {

  const startTime =
    Date.now();


  const rowValues =
    catalog
      .getRange(
        catalogRow,
        1,
        1,
        catalog.getLastColumn()
      )
      .getValues()[0];


  const rawFileId =
    String(
      rowValues[col.RAW_FILE_ID] || ""
    ).trim();


  const fileId =
    String(
      rowValues[col.FILE_ID] || ""
    ).trim();


  const fileName =
    String(
      rowValues[col.FILE_NAME] || ""
    ).trim();


  if (!rawFileId) {

    throw new Error(
      "RAW_FILE_ID missing."
    );

  }


  if (!fileId) {

    throw new Error(
      "FILE_ID missing for " +
      rawFileId
    );

  }


  /**
   * ----------------------------------------------------------
   * Mark PROCESSING
   * ----------------------------------------------------------
   */

  catalog
    .getRange(
      catalogRow,
      col.STATUS + 1
    )
    .setValue(
      "PROCESSING"
    );


  /**
   * ----------------------------------------------------------
   * Get index sheet
   * ----------------------------------------------------------
   */

  const indexSheet =
    MSE_INGEST_GetOrCreateIndexSheet_(
      ss,
      rawFileId
    );


  /**
   * ----------------------------------------------------------
   * Determine checkpoint
   * ----------------------------------------------------------
   */

  const processedRows =
    MSE_INGEST_GetProcessedRows_(
      indexSheet
    );


  /**
   * ----------------------------------------------------------
   * Get temporary Google Sheet
   * ----------------------------------------------------------
   */

  const tempInfo =
    MSE_INGEST_GetOrCreateTemporarySheet_(
      fileId,
      rawFileId
    );


  const tempSS =
    SpreadsheetApp.openById(
      tempInfo.id
    );


  const sourceSheet =
    tempSS.getSheets()[0];


  const totalSourceRows =
    sourceSheet.getLastRow();


  const totalDataRows =
    Math.max(
      0,
      totalSourceRows - 1
    );


  /**
   * ----------------------------------------------------------
   * Nothing remaining
   * ----------------------------------------------------------
   */

  if (
    processedRows >=
    totalDataRows
  ) {

    MSE_INGEST_FinalizeRawFile_(
      catalog,
      catalogRow,
      col,
      tempInfo.id
    );


    return {

      success: true,

      rawFileId:
        rawFileId,

      fileName:
        fileName,

      processedThisRun:
        0,

      processedTotal:
        processedRows,

      totalRows:
        totalDataRows,

      remainingRows:
        0,

      status:
        "PROCESSED"

    };

  }


  /**
   * ----------------------------------------------------------
   * Determine next chunk.
   * ----------------------------------------------------------
   */

  const remainingRows =
    totalDataRows -
    processedRows;


  const chunkSize =
    Math.min(
      MSE_AUTO_INGEST_CONFIG.CHUNK_SIZE,
      remainingRows
    );


  const startSourceRow =
    2 +
    processedRows;


  Logger.log(
    "--------------------------------------------------"
  );

  Logger.log(
    "RAW FILE: " +
    rawFileId
  );

  Logger.log(
    "SOURCE ROW: " +
    startSourceRow
  );

  Logger.log(
    "CHUNK: " +
    chunkSize
  );

  Logger.log(
    "ALREADY PROCESSED: " +
    processedRows
  );


  /**
   * ----------------------------------------------------------
   * Read source chunk
   * ----------------------------------------------------------
   */

  const values =
    sourceSheet
      .getRange(
        startSourceRow,
        1,
        chunkSize,
        7
      )
      .getDisplayValues();


  /**
   * ----------------------------------------------------------
   * Transform
   * ----------------------------------------------------------
   */

  const output =
    [];


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const sourceRow =
      startSourceRow +
      i;


    const parsed =
      MSE_INGEST_ParseRawRow_(
        values[i]
      );


    output.push([

      rawFileId,

      sourceRow,

      parsed.cardNumber,

      parsed.transactionDate,

      parsed.dateKey,

      parsed.registerNo,

      parsed.transactionNo,

      parsed.storeCode,

      parsed.tenderType,

      parsed.tenderAmount,

      parsed.lookupKey,

      parsed.valid
        ? "VALID"
        : parsed.status

    ]);

  }


  /**
   * ----------------------------------------------------------
   * Write chunk
   * ----------------------------------------------------------
   */

  MSE_INGEST_AppendRows_(
    indexSheet,
    output
  );


  const newProcessedRows =
    processedRows +
    values.length;


  const finished =
    newProcessedRows >=
    totalDataRows;


  /**
   * ----------------------------------------------------------
   * Update catalog
   * ----------------------------------------------------------
   */

  catalog
    .getRange(
      catalogRow,
      col.STATUS + 1
    )
    .setValue(
      finished
        ? "PROCESSED"
        : "PROCESSING"
    );


  /**
   * ----------------------------------------------------------
   * Update total rows
   * ----------------------------------------------------------
   */

  if (
    col.TOTAL_ROWS !== undefined
  ) {

    catalog
      .getRange(
        catalogRow,
        col.TOTAL_ROWS + 1
      )
      .setValue(
        totalDataRows
      );

  }


  /**
   * ----------------------------------------------------------
   * Finalize
   * ----------------------------------------------------------
   */

  if (finished) {

    MSE_INGEST_FinalizeRawFile_(
      catalog,
      catalogRow,
      col,
      tempInfo.id
    );

  }


  const duration =
    (
      Date.now() -
      startTime
    ) / 1000;


  return {

    success: true,

    rawFileId:
      rawFileId,

    fileName:
      fileName,

    processedThisRun:
      values.length,

    processedTotal:
      newProcessedRows,

    totalRows:
      totalDataRows,

    remainingRows:
      Math.max(
        0,
        totalDataRows -
        newProcessedRows
      ),

    status:
      finished
        ? "PROCESSED"
        : "PROCESSING",

    durationSeconds:
      Number(
        duration.toFixed(2)
      )

  };

}


/**
 * ============================================================
 * CHECK FOR PENDING FILES
 * ============================================================
 */
function MSE_AUTO_HasPendingFiles_(
  catalog,
  col
) {

  const data =
    catalog
      .getDataRange()
      .getValues();


  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    const status =
      String(
        data[r][col.STATUS] || ""
      )
      .trim()
      .toUpperCase();


    if (
      status === "REGISTERED" ||
      status === "PROCESSING"
    ) {

      return true;

    }

  }


  return false;

}


/**
 * ============================================================
 * SCHEDULE NEXT EXECUTION
 * ============================================================
 */
function MSE_AUTO_ScheduleNext_() {

  /**
   * Remove existing continuation triggers first.
   *
   * This prevents duplicate executions.
   */

  const triggers =
    ScriptApp
      .getProjectTriggers();


  triggers.forEach(
    function(trigger) {

      if (
        trigger
          .getHandlerFunction() ===
        "MSE_AUTO_CONTINUE_RAW_INGESTION"
      ) {

        ScriptApp
          .deleteTrigger(
            trigger
          );

      }

    }
  );


  ScriptApp
    .newTrigger(
      "MSE_AUTO_CONTINUE_RAW_INGESTION"
    )
    .timeBased()
    .after(
      MSE_AUTO_INGEST_CONFIG
        .CONTINUE_AFTER_MINUTES *
        60 *
        1000
    )
    .create();


  Logger.log(
    "Next RAW ingestion execution scheduled."
  );

}


/**
 * ============================================================
 * CONTINUATION HANDLER
 * ============================================================
 */
function MSE_AUTO_CONTINUE_RAW_INGESTION() {

  const lock =
    LockService.getScriptLock();


  if (
    !lock.tryLock(5000)
  ) {

    Logger.log(
      "Another ingestion execution is active."
    );

    return;

  }


  try {

    Logger.log(
      "=================================================="
    );

    Logger.log(
      "AUTOMATIC RAW INGESTION CONTINUATION"
    );

    Logger.log(
      "=================================================="
    );


    const result =
      MSE_AUTO_ProcessLoop_();


    Logger.log(
      JSON.stringify(
        result,
        null,
        2
      )
    );


  }
  catch (error) {

    Logger.log(
      "AUTO INGESTION ERROR:"
    );

    Logger.log(
      error.stack ||
      error.message ||
      error
    );


    /**
     * Schedule another attempt.
     *
     * The RAW catalog remains PROCESSING, so the next
     * execution resumes from the last completed row.
     */

    try {

      MSE_AUTO_ScheduleNext_();

    }
    catch (scheduleError) {

      Logger.log(
        "Could not schedule retry: " +
        scheduleError.message
      );

    }


  }
  finally {

    lock.releaseLock();

  }

}
