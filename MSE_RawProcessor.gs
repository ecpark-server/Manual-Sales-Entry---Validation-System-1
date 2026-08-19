/**
 * ============================================================
 * MSE VALIDATOR
 * RAW TRANSACTION PROCESSOR
 * ============================================================
 *
 * PURPOSE
 * ------------------------------------------------------------
 * This module prepares RAW SYSTEM transaction data for the
 * MSE validation engine.
 *
 * IMPORTANT:
 * ------------------------------------------------------------
 * 1. RAW Excel files remain in Google Drive.
 * 2. We DO NOT load millions of Excel rows into one Sheet.
 * 3. RAW_TRANSACTION_INDEX is the searchable/index layer.
 * 4. Transaction identity is:
 *
 *      DATE + STORE + REGISTER + TRANSACTION NO
 *
 * 5. CARD NUMBER is NOT part of the primary transaction key.
 *
 * This is intentional because a user's card may be:
 * - outdated
 * - replaced
 * - newly issued
 * - different from the SYSTEM card
 *
 * ============================================================
 */


/**
 * ============================================================
 * NORMALIZE GENERAL VALUE
 * ============================================================
 */
function mseRawNormalize_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value)
    .trim()
    .toUpperCase();
}


/**
 * ============================================================
 * NORMALIZE CARD NUMBER
 *
 * Keeps digits only.
 * ============================================================
 */
function mseRawNormalizeCard_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value)
    .replace(/\D/g, "");
}


/**
 * ============================================================
 * NORMALIZE TRANSACTION DATE
 *
 * SYSTEM examples:
 *
 * 20260430
 * 2026-04-30
 * 04/30/2026
 * ============================================================
 */
function mseRawNormalizeDate_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }


  /**
   * Google Sheets Date object
   */
  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );

  }


  const text =
    String(value)
      .trim();


  /**
   * YYYYMMDD
   */
  if (
    /^\d{8}$/.test(text)
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

  if (match) {

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

  if (match) {

    return (
      match[3] +
      String(match[1]).padStart(2, "0") +
      String(match[2]).padStart(2, "0")
    );

  }


  /**
   * Try normal JavaScript date
   */
  const parsed =
    new Date(text);

  if (
    !isNaN(parsed.getTime())
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
 * NORMALIZE REGISTER NUMBER
 * ============================================================
 */
function mseRawNormalizeRegister_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }

  const result =
    String(value)
      .trim()
      .replace(/^0+/, "");

  return result || "0";
}


/**
 * ============================================================
 * NORMALIZE TRANSACTION NUMBER
 * ============================================================
 */
function mseRawNormalizeTransaction_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }

  const result =
    String(value)
      .trim()
      .replace(/^0+/, "");

  return result || "0";
}


/**
 * ============================================================
 * NORMALIZE STORE CODE
 * ============================================================
 */
function mseRawNormalizeStore_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }

  const result =
    String(value)
      .trim()
      .replace(/^0+/, "");

  return result || "0";
}


/**
 * ============================================================
 * NORMALIZE TENDER TYPE
 * ============================================================
 */
function mseRawNormalizeTender_(value) {

  return mseRawNormalize_(value);
}


/**
 * ============================================================
 * NORMALIZE AMOUNT
 * ============================================================
 */
function mseRawNormalizeAmount_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return 0;
  }


  if (
    typeof value === "number"
  ) {

    return Math.round(
      value * 100
    ) / 100;

  }


  const cleaned =
    String(value)
      .replace(/,/g, "")
      .replace(/[^\d.-]/g, "");


  const amount =
    parseFloat(cleaned);


  if (
    isNaN(amount)
  ) {
    return 0;
  }


  return Math.round(
    amount * 100
  ) / 100;
}


/**
 * ============================================================
 * BUILD PRIMARY TRANSACTION KEY
 *
 * CARD NUMBER IS INTENTIONALLY NOT INCLUDED.
 *
 * PRIMARY IDENTITY:
 *
 * DATE
 * STORE
 * REGISTER
 * TRANSACTION NO
 * ============================================================
 */
function buildMSETransactionKey_(
  transactionDate,
  storeCode,
  registerNo,
  transactionNo
) {

  const date =
    mseRawNormalizeDate_(
      transactionDate
    );

  const store =
    mseRawNormalizeStore_(
      storeCode
    );

  const register =
    mseRawNormalizeRegister_(
      registerNo
    );

  const transaction =
    mseRawNormalizeTransaction_(
      transactionNo
    );


  return [
    date,
    store,
    register,
    transaction
  ].join("|");
}


/**
 * ============================================================
 * BUILD CARD + TRANSACTION LOOKUP KEY
 * ============================================================
 */
function buildMSERawLookupKey_(
  cardNumber,
  transactionDate,
  storeCode,
  registerNo,
  transactionNo
) {

  const card =
    mseRawNormalizeCard_(
      cardNumber
    );


  const transactionKey =
    buildMSETransactionKey_(
      transactionDate,
      storeCode,
      registerNo,
      transactionNo
    );


  return [
    card,
    transactionKey
  ].join("|");
}


/**
 * ============================================================
 * BUILD RAW INDEX RECORD
 *
 * RAW SYSTEM SOURCE:
 *
 * Card Number
 * Date
 * Register Num
 * Trans Num
 * Store Code
 * Tender Type
 * Tender Amount
 * ============================================================
 */
function buildMSERawIndexRecord_(
  rawFileId,
  rawRowId,
  cardNumber,
  transactionDate,
  registerNo,
  transactionNo,
  storeCode,
  tenderType,
  tenderAmount,
  sourceRow,
  sourceFile
) {

  const normalizedCard =
    mseRawNormalizeCard_(
      cardNumber
    );


  const normalizedDate =
    mseRawNormalizeDate_(
      transactionDate
    );


  const normalizedRegister =
    mseRawNormalizeRegister_(
      registerNo
    );


  const normalizedTransaction =
    mseRawNormalizeTransaction_(
      transactionNo
    );


  const normalizedStore =
    mseRawNormalizeStore_(
      storeCode
    );


  const normalizedTender =
    mseRawNormalizeTender_(
      tenderType
    );


  const normalizedAmount =
    mseRawNormalizeAmount_(
      tenderAmount
    );


  const transactionKey =
    buildMSETransactionKey_(
      normalizedDate,
      normalizedStore,
      normalizedRegister,
      normalizedTransaction
    );


  const lookupKey =
    buildMSERawLookupKey_(
      normalizedCard,
      normalizedDate,
      normalizedStore,
      normalizedRegister,
      normalizedTransaction
    );


  return [

    rawFileId,

    rawRowId,

    normalizedCard,

    normalizedDate,

    normalizedDate
      ? normalizedDate.substring(0, 6)
      : "",

    normalizedRegister,

    normalizedTransaction,

    normalizedStore,

    normalizedTender,

    normalizedAmount,

    lookupKey,

    transactionKey,

    sourceRow,

    sourceFile,

    1

  ];

}


/**
 * ============================================================
 * CALCULATE TRANSACTION TOTALS
 *
 * Example:
 *
 * CA = 3,000
 * W8 =   500
 *
 * GROSS = 3,500
 * W8    =   500
 * NET   = 3,000
 * ============================================================
 */
function calculateMSERawTransactionTotal_(
  tenderRows
) {

  let gross = 0;

  let w8 = 0;

  const breakdown = [];


  tenderRows.forEach(
    row => {

      const tender =
        mseRawNormalizeTender_(
          row.tenderType
        );


      const amount =
        mseRawNormalizeAmount_(
          row.tenderAmount
        );


      gross += amount;


      if (
        tender === "W8"
      ) {

        w8 += amount;

      }


      breakdown.push({

        tenderType:
          tender,

        amount:
          amount

      });

    }
  );


  gross =
    Math.round(
      gross * 100
    ) / 100;


  w8 =
    Math.round(
      w8 * 100
    ) / 100;


  const net =
    Math.round(
      (gross - w8) * 100
    ) / 100;


  const w8Only =
    tenderRows.length > 0 &&
    tenderRows.every(
      row =>
        mseRawNormalizeTender_(
          row.tenderType
        ) === "W8"
    );


  return {

    gross:
      gross,

    w8:
      w8,

    net:
      net,

    w8Only:
      w8Only,

    breakdown:
      breakdown

  };

}


/**
 * ============================================================
 * GROUP RAW ROWS BY TRANSACTION
 * ============================================================
 */
function groupMSERawTransactions_(
  rows
) {

  const grouped =
    new Map();


  rows.forEach(
    row => {

      const key =
        buildMSETransactionKey_(
          row.transactionDate,
          row.storeCode,
          row.registerNo,
          row.transactionNo
        );


      if (
        !grouped.has(key)
      ) {

        grouped.set(
          key,
          []
        );

      }


      grouped
        .get(key)
        .push(row);

    }
  );


  return grouped;
}


/**
 * ============================================================
 * CREATE TRANSACTION SUMMARY
 * ============================================================
 */
function createMSERawTransactionSummary_(
  tenderRows
) {

  if (
    !tenderRows ||
    tenderRows.length === 0
  ) {

    return null;

  }


  const first =
    tenderRows[0];


  const totals =
    calculateMSERawTransactionTotal_(
      tenderRows
    );


  return {

    transactionKey:
      buildMSETransactionKey_(
        first.transactionDate,
        first.storeCode,
        first.registerNo,
        first.transactionNo
      ),


    cardNumbers:
      [
        ...new Set(
          tenderRows
            .map(
              row =>
                mseRawNormalizeCard_(
                  row.cardNumber
                )
            )
            .filter(Boolean)
        )
      ],


    transactionDate:
      mseRawNormalizeDate_(
        first.transactionDate
      ),


    storeCode:
      mseRawNormalizeStore_(
        first.storeCode
      ),


    registerNo:
      mseRawNormalizeRegister_(
        first.registerNo
      ),


    transactionNo:
      mseRawNormalizeTransaction_(
        first.transactionNo
      ),


    gross:
      totals.gross,


    w8:
      totals.w8,


    net:
      totals.net,


    w8Only:
      totals.w8Only,


    tenderBreakdown:
      totals.breakdown,


    tenderCount:
      tenderRows.length

  };

}


/**
 * ============================================================
 * VALIDATE RAW TRANSACTION STRUCTURE
 * ============================================================
 */
function validateMSERawTransactionStructure_(
  transaction
) {

  const errors = [];


  if (
    !transaction.transactionDate
  ) {

    errors.push(
      "Transaction Date"
    );

  }


  if (
    !transaction.storeCode
  ) {

    errors.push(
      "Store Code"
    );

  }


  if (
    !transaction.registerNo
  ) {

    errors.push(
      "Register Number"
    );

  }


  if (
    !transaction.transactionNo
  ) {

    errors.push(
      "Transaction Number"
    );

  }


  if (
    transaction.gross === undefined ||
    transaction.gross === null
  ) {

    errors.push(
      "Tender Amount"
    );

  }


  return {

    valid:
      errors.length === 0,

    errors:
      errors

  };

}


/**
 * ============================================================
 * BUILD RAW TRANSACTION REFERENCE
 * ============================================================
 */
function buildMSERawReference_(
  transactionRows
) {

  const summary =
    createMSERawTransactionSummary_(
      transactionRows
    );


  if (!summary) {

    return null;

  }


  const structure =
    validateMSERawTransactionStructure_(
      summary
    );


  return {

    found:
      true,

    validStructure:
      structure.valid,

    structureErrors:
      structure.errors,

    transaction:
      summary

  };

}


/**
 * ============================================================
 * TEST FUNCTION
 *
 * This does NOT touch your RAW files.
 *
 * It tests the W8 calculation using sample data.
 * ============================================================
 */
function testMSERawTransactionProcessor() {

  const sampleRows = [

    {

      cardNumber:
        "70025010",

      transactionDate:
        "20260430",

      storeCode:
        "343",

      registerNo:
        "14",

      transactionNo:
        "5942",

      tenderType:
        "CA",

      tenderAmount:
        "3000.00"

    },

    {

      cardNumber:
        "70025010",

      transactionDate:
        "20260430",

      storeCode:
        "343",

      registerNo:
        "14",

      transactionNo:
        "5942",

      tenderType:
        "W8",

      tenderAmount:
        "500.00"

    }

  ];


  const result =
    buildMSERawReference_(
      sampleRows
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
