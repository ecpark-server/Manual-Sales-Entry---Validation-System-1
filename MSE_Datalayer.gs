/**
 * ============================================================
 * MSE VALIDATOR
 * DATA STRUCTURE INITIALIZATION
 * ============================================================
 *
 * IMPORTANT:
 * This version DOES NOT depend on MSE_CONFIG.
 *
 * If this Apps Script is bound to the MSE DATABASE spreadsheet,
 * leave MSE_DATABASE_FILE_ID blank.
 *
 * If this is a standalone Apps Script project, put the
 * Spreadsheet ID of the MSE DATABASE below.
 * ============================================================
 */

const MSE_DATABASE_FILE_ID = "";


/**
 * ============================================================
 * GET MSE DATABASE
 * ============================================================
 */
function getMSEDatabase_() {

  // ----------------------------------------------------------
  // OPTION 1:
  // Use configured database ID
  // ----------------------------------------------------------
  if (
    typeof MSE_DATABASE_FILE_ID !== "undefined" &&
    String(MSE_DATABASE_FILE_ID).trim() !== ""
  ) {

    return SpreadsheetApp.openById(
      String(MSE_DATABASE_FILE_ID).trim()
    );

  }

  // ----------------------------------------------------------
  // OPTION 2:
  // Use the spreadsheet this script is bound to
  // ----------------------------------------------------------
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();

  if (activeSS) {
    return activeSS;
  }

  throw new Error(
    "MSE Database spreadsheet could not be determined. " +
    "Either bind this script to the MSE database spreadsheet " +
    "or enter its Spreadsheet ID in MSE_DATABASE_FILE_ID."
  );
}


/**
 * ============================================================
 * CREATE / REPAIR MSE DATABASE STRUCTURE
 * ============================================================
 */
function createMSEDataStructure() {

  const startTime = new Date();

  const ss = getMSEDatabase_();

  Logger.log(
    "MSE DATABASE: " +
    ss.getName() +
    " | " +
    ss.getId()
  );


  /**
   * ==========================================================
   * DATABASE DEFINITIONS
   * ==========================================================
   */

  const definitions = {

    // ========================================================
    // REQUEST MASTER
    // ========================================================
    MSE_REQUESTS: [

      "REQUEST_ID",
      "REQUEST_DATE",
      "REQUESTED_BY",
      "STATUS",

      "TOTAL_TRANSACTIONS",
      "TOTAL_TICKETS",

      "VALID_TRANSACTIONS",
      "ERROR_TRANSACTIONS",

      "CREATED_TIMESTAMP",
      "UPDATED_TIMESTAMP"

    ],


    // ========================================================
    // USER-IMPORTED MSE TRANSACTIONS
    // ========================================================
    MSE_TRANSACTIONS: [

      "REQUEST_ID",
      "ROW_ID",

      "TICKET_ID",
      "DOR",
      "STORE",

      "CARD_NUMBER",
      "CONFIGURATION",
      "DOT",

      "STORE_CODE",
      "REGISTER_NO",
      "TRANS_NO",

      "GROSS_AMOUNT",
      "ISSUE_CONCERN",

      // ------------------------------------------------------
      // ORIGINAL USER INPUT
      // ------------------------------------------------------

      "ORIGINAL_CARD_NUMBER",
      "ORIGINAL_DOR",
      "ORIGINAL_STORE_CODE",
      "ORIGINAL_REGISTER_NO",
      "ORIGINAL_TRANS_NO",
      "ORIGINAL_GROSS_AMOUNT",

      "IMPORT_STATUS"

    ],


    // ========================================================
    // VALIDATION RESULT
    // ========================================================
    MSE_VALIDATION: [

      "REQUEST_ID",
      "ROW_ID",
      "TICKET_ID",

      // ------------------------------------------------------
      // USER SIDE
      // ------------------------------------------------------

      "USER_CARD",
      "USER_DATE",
      "USER_STORE",
      "USER_REGISTER",
      "USER_TRANS_NO",
      "USER_GROSS",

      // ------------------------------------------------------
      // SYSTEM SIDE
      // ------------------------------------------------------

      "SYSTEM_CARD",
      "SYSTEM_DATE",
      "SYSTEM_STORE",
      "SYSTEM_REGISTER",
      "SYSTEM_TRANS_NO",

      "SYSTEM_GROSS",
      "SYSTEM_W8",
      "SYSTEM_NET",

      "TENDER_BREAKDOWN",

      // ------------------------------------------------------
      // PARAMETER RESULTS
      // ------------------------------------------------------

      "DATE_RESULT",
      "STORE_RESULT",
      "REGISTER_RESULT",
      "TRANS_RESULT",
      "GROSS_RESULT",
      "CARD_RESULT",

      // ------------------------------------------------------
      // FINAL DECISION
      // ------------------------------------------------------

      "PARAMETER_RESULT",
      "QUALIFICATION",
      "OVERALL_RESULT",
      "RESULT_REASON",

      // ------------------------------------------------------
      // W8 LOGIC
      // ------------------------------------------------------

      "W8_ONLY",

      // ------------------------------------------------------
      // RAW SOURCE REFERENCE
      // ------------------------------------------------------

      "RAW_FILE_ID",
      "RAW_ROW_REFERENCE",

      // ------------------------------------------------------
      // AUDIT
      // ------------------------------------------------------

      "VALIDATED_TIMESTAMP",
      "VALIDATED_BY"

    ],


    // ========================================================
    // USER CORRECTIONS / MANUAL OVERRIDES
    // ========================================================
    MSE_CORRECTIONS: [

      "REQUEST_ID",
      "ROW_ID",

      "FIELD",

      "ORIGINAL_VALUE",
      "CORRECTED_VALUE",

      "REASON",

      "CORRECTED_BY",
      "CORRECTED_TIMESTAMP",

      "APPROVAL_STATUS"

    ],


    // ========================================================
    // RAW FILE MASTER CATALOG
    // ========================================================
    RAW_FILE_CATALOG: [

      "RAW_FILE_ID",

      "FILE_NAME",
      "FILE_ID",
      "FILE_TYPE",

      "PERIOD_FROM",
      "PERIOD_TO",

      "TOTAL_ROWS",

      "STATUS",

      "UPLOADED_BY",
      "UPLOADED_TIMESTAMP",

      "PROCESSED_TIMESTAMP",

      "ERROR_MESSAGE",

      "VERSION",
      "ACTIVE"

    ],


    // ========================================================
    // RAW SYSTEM TRANSACTION INDEX
    // ========================================================
    RAW_TRANSACTION_INDEX: [

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

      "LOOKUP_KEY"

    ]

  };


  /**
   * ==========================================================
   * CREATE / REPAIR EACH SHEET
   * ==========================================================
   */

  Object.entries(definitions).forEach(
    ([sheetName, headers]) => {

      let sheet = ss.getSheetByName(sheetName);


      // ------------------------------------------------------
      // CREATE SHEET IF MISSING
      // ------------------------------------------------------

      if (!sheet) {

        sheet = ss.insertSheet(sheetName);

        Logger.log(
          "Created sheet: " + sheetName
        );

      }


      // ------------------------------------------------------
      // ENSURE ENOUGH COLUMNS
      // ------------------------------------------------------

      const requiredColumns = headers.length;

      const currentColumns =
        sheet.getMaxColumns();

      if (currentColumns < requiredColumns) {

        sheet.insertColumnsAfter(
          currentColumns,
          requiredColumns - currentColumns
        );

      }


      // ------------------------------------------------------
      // WRITE / REPAIR HEADER
      // ------------------------------------------------------

      const existingHeaders =
        sheet
          .getRange(
            1,
            1,
            1,
            requiredColumns
          )
          .getValues()[0];


      const correctedHeaders =
        headers.map(
          (header, index) => {

            const existing =
              existingHeaders[index];

            return existing === ""
              ? header
              : existing === header
                ? existing
                : header;

          }
        );


      sheet
        .getRange(
          1,
          1,
          1,
          requiredColumns
        )
        .setValues([
          correctedHeaders
        ]);


      // ------------------------------------------------------
      // HEADER FORMATTING
      // ------------------------------------------------------

      sheet
        .getRange(
          1,
          1,
          1,
          requiredColumns
        )
        .setFontWeight("bold");


      sheet.setFrozenRows(1);


      Logger.log(
        "Verified sheet: " +
        sheetName +
        " (" +
        requiredColumns +
        " columns)"
      );

    }
  );


  SpreadsheetApp.flush();


  /**
   * ==========================================================
   * COMPLETION
   * ==========================================================
   */

  const duration =
    ((new Date() - startTime) / 1000)
      .toFixed(2);


  const message =
    "MSE database structure created/repaired successfully.\n\n" +

    "Database: " +
    ss.getName() +
    "\n" +

    "Spreadsheet ID: " +
    ss.getId() +
    "\n\n" +

    "Sheets verified:\n" +
    "• MSE_REQUESTS\n" +
    "• MSE_TRANSACTIONS\n" +
    "• MSE_VALIDATION\n" +
    "• MSE_CORRECTIONS\n" +
    "• RAW_FILE_CATALOG\n" +
    "• RAW_TRANSACTION_INDEX\n\n" +

    "Processing time: " +
    duration +
    " seconds";


  Logger.log(message);


  return {
    success: true,
    databaseId: ss.getId(),
    databaseName: ss.getName(),
    durationSeconds: Number(duration),
    message: message
  };

}
