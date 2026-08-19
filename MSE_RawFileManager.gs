/**
 * ============================================================
 * MSE VALIDATOR
 * RAW FILE MANAGER
 * ============================================================
 *
 * PURPOSE
 * ------------------------------------------------------------
 * Manages RAW SYSTEM Excel files used by the MSE Validator.
 *
 * EXISTING DRIVE STRUCTURE
 *
 * ACTIVE
 * ARCHIVE
 * ERROR
 * TEMP
 *
 * WORKFLOW
 *
 * TEMP
 *   ↓
 * SCAN
 *   ↓
 * REGISTER
 *   ↓
 * VALIDATE / PROCESS
 *   ↓
 * ACTIVE
 *
 * Failed processing
 *   ↓
 * ERROR
 *
 * Superseded version
 *   ↓
 * ARCHIVE
 *
 * IMPORTANT
 * ------------------------------------------------------------
 * This file DOES NOT declare:
 *
 * MSE_DATABASE_FILE_ID
 * getMSEDatabase_()
 *
 * Those already exist in MSE_DataLayer.gs.
 *
 * ============================================================
 */


/**
 * ============================================================
 * EXISTING RAW DRIVE FOLDERS
 * ============================================================
 *
 * These are FOLDER IDs.
 *
 * They are NOT Excel file IDs.
 * ============================================================
 */

const MSE_RAW_FOLDER_IDS = {

  ACTIVE:
    "1YYA0mb198-JZpbU36Yq9r3XYN4ZZIFer",

  ARCHIVE:
    "1PQYqg00UvlzY2KP0IEEis0-fbADI5IFZ",

  ERROR:
    "1hzqY4IpnvHQdns3zE-pkq35dg0zPrbjX",

  TEMP:
    "1dJO6BOiv7zH180au9Pz37OAdSaLNIajj"

};


/**
 * ============================================================
 * SUPPORTED RAW FILE TYPES
 * ============================================================
 */

const MSE_RAW_SUPPORTED_EXTENSIONS = [

  "xlsx",
  "xls",
  "csv"

];


/**
 * ============================================================
 * GET RAW STORAGE FOLDERS
 * ============================================================
 */
function getMSERawStorage_() {

  const storage = {};


  Object.keys(
    MSE_RAW_FOLDER_IDS
  ).forEach(
    key => {

      const folderId =
        String(
          MSE_RAW_FOLDER_IDS[key]
        ).trim();


      if (!folderId) {

        throw new Error(
          "RAW folder ID is missing for: " +
          key
        );

      }


      try {

        const folder =
          DriveApp.getFolderById(
            folderId
          );


        storage[
          key.toLowerCase()
        ] = folder;


      } catch (error) {

        throw new Error(
          "Unable to access RAW " +
          key +
          " folder.\n\n" +
          "Folder ID:\n" +
          folderId +
          "\n\n" +
          error.message
        );

      }

    }
  );


  return storage;

}


/**
 * ============================================================
 * GENERATE RAW FILE ID
 * ============================================================
 */
function generateMSERawFileId_() {

  const timestamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMddHHmmss"
    );


  const randomPart =
    Utilities.getUuid()
      .replace(
        /-/g,
        ""
      )
      .substring(
        0,
        8
      )
      .toUpperCase();


  return (
    "RAW-" +
    timestamp +
    "-" +
    randomPart
  );

}


/**
 * ============================================================
 * GET CURRENT USER
 * ============================================================
 */
function getMSECurrentUser_() {

  try {

    const email =
      Session
        .getActiveUser()
        .getEmail();


    return (
      email ||
      "UNKNOWN"
    );

  } catch (error) {

    return "UNKNOWN";

  }

}


/**
 * ============================================================
 * GET RAW FILE CATALOG SHEET
 * ============================================================
 */
function getMSERawCatalogSheet_() {

  const ss =
    getMSEDatabase_();


  const sheet =
    ss.getSheetByName(
      "RAW_FILE_CATALOG"
    );


  if (!sheet) {

    throw new Error(
      "RAW_FILE_CATALOG does not exist.\n\n" +
      "Run createMSEDataStructure() first."
    );

  }


  return sheet;

}


/**
 * ============================================================
 * GET RAW TRANSACTION INDEX SHEET
 * ============================================================
 */
function getMSERawIndexSheet_() {

  const ss =
    getMSEDatabase_();


  const sheet =
    ss.getSheetByName(
      "RAW_TRANSACTION_INDEX"
    );


  if (!sheet) {

    throw new Error(
      "RAW_TRANSACTION_INDEX does not exist.\n\n" +
      "Run createMSEDataStructure() first."
    );

  }


  return sheet;

}


/**
 * ============================================================
 * CHECK FILE EXTENSION
 * ============================================================
 */
function getMSERawFileExtension_(
  fileName
) {

  const name =
    String(
      fileName || ""
    ).trim();


  if (!name) {

    return "";

  }


  const parts =
    name.split(".");


  if (
    parts.length < 2
  ) {

    return "";

  }


  return parts[
    parts.length - 1
  ]
    .toLowerCase();

}


/**
 * ============================================================
 * CHECK IF SUPPORTED
 * ============================================================
 */
function isSupportedMSERawFile_(
  extension
) {

  return (
    MSE_RAW_SUPPORTED_EXTENSIONS
      .indexOf(
        String(extension)
          .toLowerCase()
      ) !== -1
  );

}


/**
 * ============================================================
 * FIND RAW FILE IN CATALOG
 * ============================================================
 */
function findMSERawCatalogByFileId_(
  fileId
) {

  const sheet =
    getMSERawCatalogSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return null;

  }


  const columnCount =
    Math.max(
      16,
      sheet.getLastColumn()
    );


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        columnCount
      )
      .getValues();


  const target =
    String(
      fileId
    ).trim();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const existingFileId =
      String(
        data[i][2] || ""
      ).trim();


    if (
      existingFileId ===
      target
    ) {

      return {

        row:
          i + 2,

        data:
          data[i]

      };

    }

  }


  return null;

}


/**
 * ============================================================
 * FIND RAW CATALOG BY RAW FILE ID
 * ============================================================
 */
function findMSERawCatalogByRawFileId_(
  rawFileId
) {

  const sheet =
    getMSERawCatalogSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return null;

  }


  const columnCount =
    Math.max(
      16,
      sheet.getLastColumn()
    );


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        columnCount
      )
      .getValues();


  const target =
    String(
      rawFileId
    ).trim();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const existingRawFileId =
      String(
        data[i][0] || ""
      ).trim();


    if (
      existingRawFileId ===
      target
    ) {

      return {

        row:
          i + 2,

        data:
          data[i]

      };

    }

  }


  return null;

}


/**
 * ============================================================
 * DETERMINE NEXT VERSION
 *
 * Version is determined using:
 *
 * PERIOD_FROM
 * PERIOD_TO
 *
 * If period is not known yet,
 * version starts at 1.
 * ============================================================
 */
function getNextMSERawVersion_(
  periodFrom,
  periodTo
) {

  const sheet =
    getMSERawCatalogSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return 1;

  }


  const columnCount =
    Math.max(
      16,
      sheet.getLastColumn()
    );


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        columnCount
      )
      .getValues();


  const targetFrom =
    String(
      periodFrom || ""
    ).trim();


  const targetTo =
    String(
      periodTo || ""
    ).trim();


  let highestVersion =
    0;


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const existingFrom =
      String(
        data[i][4] || ""
      ).trim();


    const existingTo =
      String(
        data[i][5] || ""
      ).trim();


    if (
      existingFrom ===
      targetFrom &&

      existingTo ===
      targetTo
    ) {

      const version =
        Number(
          data[i][13]
        ) || 0;


      highestVersion =
        Math.max(
          highestVersion,
          version
        );

    }

  }


  return (
    highestVersion + 1
  );

}


/**
 * ============================================================
 * DEACTIVATE PREVIOUS ACTIVE VERSION
 *
 * Used when a new version for the SAME period is registered.
 * ============================================================
 */
function deactivateMSERawVersions_(
  periodFrom,
  periodTo
) {

  const sheet =
    getMSERawCatalogSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return 0;

  }


  const columnCount =
    Math.max(
      16,
      sheet.getLastColumn()
    );


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        columnCount
      )
      .getValues();


  const targetFrom =
    String(
      periodFrom || ""
    ).trim();


  const targetTo =
    String(
      periodTo || ""
    ).trim();


  let count =
    0;


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const existingFrom =
      String(
        data[i][4] || ""
      ).trim();


    const existingTo =
      String(
        data[i][5] || ""
      ).trim();


    const activeValue =
      data[i][14];


    const isActive =
      activeValue === true ||
      String(
        activeValue
      ).toUpperCase() ===
        "TRUE";


    if (
      existingFrom ===
      targetFrom &&

      existingTo ===
      targetTo &&

      isActive
    ) {

      sheet
        .getRange(
          i + 2,
          15
        )
        .setValue(
          false
        );


      count++;

    }

  }


  return count;

}


/**
 * ============================================================
 * SCAN TEMP FOLDER
 *
 * This only discovers files.
 *
 * It DOES NOT:
 * - process files
 * - modify files
 * - move files
 * - create catalog records
 *
 * ============================================================
 */
function scanMSERawTempFolder() {

  const startTime =
    new Date();


  const storage =
    getMSERawStorage_();


  const tempFolder =
    storage.temp;


  const files =
    tempFolder.getFiles();


  const results = [];


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    const fileName =
      file.getName();


    const extension =
      getMSERawFileExtension_(
        fileName
      );


    const supported =
      isSupportedMSERawFile_(
        extension
      );


    const fileSizeMB =
      Math.round(
        (
          file.getSize() /
          1024 /
          1024
        ) * 100
      ) / 100;


    results.push({

      fileId:
        file.getId(),

      fileName:
        fileName,

      extension:
        extension,

      mimeType:
        file.getMimeType(),

      sizeMB:
        fileSizeMB,

      supported:
        supported

    });

  }


  const duration =
    (
      (new Date() - startTime) /
      1000
    ).toFixed(2);


  Logger.log(
    "TEMP RAW FILES FOUND: " +
    results.length
  );


  Logger.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );


  Logger.log(
    "SCAN TIME: " +
    duration +
    " seconds"
  );


  return {

    success:
      true,

    count:
      results.length,

    files:
      results,

    durationSeconds:
      Number(duration)

  };

}


/**
 * ============================================================
 * REGISTER ONE RAW FILE
 *
 * This registers metadata only.
 *
 * It does NOT ingest millions of rows.
 *
 * Period can initially be blank.
 * The ingestion engine will determine the actual period.
 *
 * ============================================================
 */
function registerMSERawFile(
  options
) {

  if (
    !options
  ) {

    throw new Error(
      "RAW file registration options are required."
    );

  }


  const fileId =
    String(
      options.fileId || ""
    ).trim();


  if (
    !fileId
  ) {

    throw new Error(
      "RAW Excel FILE ID is required."
    );

  }


  let file;


  try {

    file =
      DriveApp.getFileById(
        fileId
      );

  } catch (error) {

    throw new Error(
      "Cannot access RAW file.\n\n" +
      "FILE ID:\n" +
      fileId +
      "\n\n" +
      error.message
    );

  }


  const fileName =
    file.getName();


  const extension =
    getMSERawFileExtension_(
      fileName
    );


  if (
    !isSupportedMSERawFile_(
      extension
    )
  ) {

    throw new Error(
      "Unsupported RAW file type: " +
      extension +
      ".\n\n" +
      "Supported types: XLSX, XLS, CSV."
    );

  }


  /**
   * ----------------------------------------------------------
   * Prevent duplicate registration
   * ----------------------------------------------------------
   */

  const existing =
    findMSERawCatalogByFileId_(
      fileId
    );


  if (
    existing
  ) {

    return {

      success:
        false,

      duplicate:
        true,

      message:
        "This RAW file is already registered.",

      catalogRow:
        existing.row,

      rawFileId:
        existing.data[0],

      fileName:
        existing.data[1]

    };

  }


  /**
   * ----------------------------------------------------------
   * Period
   *
   * May initially be blank.
   * ----------------------------------------------------------
   */

  const periodFrom =
    String(
      options.periodFrom || ""
    ).trim();


  const periodTo =
    String(
      options.periodTo || ""
    ).trim();


  /**
   * ----------------------------------------------------------
   * Version
   * ----------------------------------------------------------
   */

  const version =
    getNextMSERawVersion_(
      periodFrom,
      periodTo
    );


  /**
   * ----------------------------------------------------------
   * Generate RAW FILE ID
   * ----------------------------------------------------------
   */

  const rawFileId =
    generateMSERawFileId_();


  /**
   * ----------------------------------------------------------
   * Deactivate old version
   *
   * Only when period is known.
   * ----------------------------------------------------------
   */

  let deactivated =
    0;


  if (
    periodFrom &&
    periodTo
  ) {

    deactivated =
      deactivateMSERawVersions_(
        periodFrom,
        periodTo
      );

  }


  /**
   * ----------------------------------------------------------
   * File information
   * ----------------------------------------------------------
   */

  const fileSizeMB =
    Math.round(
      (
        file.getSize() /
        1024 /
        1024
      ) * 100
    ) / 100;


  const uploadedBy =
    getMSECurrentUser_();


  const uploadedTimestamp =
    new Date();


  /**
   * ----------------------------------------------------------
   * Append catalog record
   * ----------------------------------------------------------
   */

  const sheet =
    getMSERawCatalogSheet_();


  sheet.appendRow([

    rawFileId,

    fileName,

    fileId,

    extension.toUpperCase(),

    periodFrom,

    periodTo,

    0,

    fileSizeMB,

    "REGISTERED",

    uploadedBy,

    uploadedTimestamp,

    "",

    "",

    version,

    true,

    options.notes || ""

  ]);


  return {

    success:
      true,

    rawFileId:
      rawFileId,

    fileId:
      fileId,

    fileName:
      fileName,

    fileType:
      extension.toUpperCase(),

    periodFrom:
      periodFrom,

    periodTo:
      periodTo,

    version:
      version,

    deactivatedPreviousVersions:
      deactivated,

    fileSizeMB:
      fileSizeMB,

    status:
      "REGISTERED"

  };

}


/**
 * ============================================================
 * REGISTER ALL FILES FOUND IN TEMP
 *
 * This is the new main test function.
 *
 * It:
 *
 * 1. Scans TEMP.
 * 2. Finds actual Excel files.
 * 3. Registers each file.
 * 4. Does NOT process the Excel contents yet.
 *
 * ============================================================
 */
function registerAllMSERawTempFiles() {

  const startTime =
    new Date();


  const scan =
    scanMSERawTempFolder();


  if (
    !scan.success
  ) {

    throw new Error(
      "RAW TEMP scan failed."
    );

  }


  const results = [];


  scan.files.forEach(
    item => {

      /**
       * ------------------------------------------------------
       * Skip unsupported files
       * ------------------------------------------------------
       */

      if (
        !item.supported
      ) {

        results.push({

          fileId:
            item.fileId,

          fileName:
            item.fileName,

          status:
            "SKIPPED",

          reason:
            "Unsupported file type"

        });


        return;

      }


      try {

        const result =
          registerMSERawFile({

            fileId:
              item.fileId,

            periodFrom:
              "",

            periodTo:
              "",

            notes:
              "Registered from TEMP folder. Period pending RAW ingestion."

          });


        results.push(
          result
        );


      } catch (error) {

        results.push({

          fileId:
            item.fileId,

          fileName:
            item.fileName,

          status:
            "ERROR",

          error:
            error.message

        });

      }

    }
  );


  const duration =
    (
      (new Date() - startTime) /
      1000
    ).toFixed(2);


  Logger.log(
    "RAW TEMP REGISTRATION COMPLETE"
  );


  Logger.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );


  Logger.log(
    "Processing time: " +
    duration +
    " seconds"
  );


  return {

    success:
      true,

    scanned:
      scan.count,

    results:
      results,

    durationSeconds:
      Number(duration)

  };

}


/**
 * ============================================================
 * UPDATE RAW CATALOG STATUS
 * ============================================================
 */
function updateMSERawCatalogStatus_(
  rawFileId,
  status,
  options
) {

  const found =
    findMSERawCatalogByRawFileId_(
      rawFileId
    );


  if (
    !found
  ) {

    throw new Error(
      "RAW_FILE_ID not found in catalog: " +
      rawFileId
    );

  }


  const sheet =
    getMSERawCatalogSheet_();


  const row =
    found.row;


  const opts =
    options || {};


  /**
   * STATUS
   */
  sheet
    .getRange(
      row,
      9
    )
    .setValue(
      status
    );


  /**
   * TOTAL ROWS
   */
  if (
    opts.totalRows !== undefined
  ) {

    sheet
      .getRange(
        row,
        7
      )
      .setValue(
        opts.totalRows
      );

  }


  /**
   * PERIOD FROM
   */
  if (
    opts.periodFrom !== undefined
  ) {

    sheet
      .getRange(
        row,
        5
      )
      .setValue(
        opts.periodFrom
      );

  }


  /**
   * PERIOD TO
   */
  if (
    opts.periodTo !== undefined
  ) {

    sheet
      .getRange(
        row,
        6
      )
      .setValue(
        opts.periodTo
      );

  }


  /**
   * PROCESSED TIMESTAMP
   */
  if (
    opts.processedTimestamp !== undefined
  ) {

    sheet
      .getRange(
        row,
        12
      )
      .setValue(
        opts.processedTimestamp
      );

  }


  /**
   * ERROR MESSAGE
   */
  if (
    opts.errorMessage !== undefined
  ) {

    sheet
      .getRange(
        row,
        13
      )
      .setValue(
        opts.errorMessage
      );

  }


  /**
   * ACTIVE
   */
  if (
    opts.active !== undefined
  ) {

    sheet
      .getRange(
        row,
        15
      )
      .setValue(
        opts.active
      );

  }


  /**
   * NOTES
   */
  if (
    opts.notes !== undefined
  ) {

    sheet
      .getRange(
        row,
        16
      )
      .setValue(
        opts.notes
      );

  }


  SpreadsheetApp.flush();


  return {

    success:
      true,

    row:
      row,

    rawFileId:
      rawFileId,

    status:
      status

  };

}


/**
 * ============================================================
 * MOVE FILE TO FOLDER
 *
 * IMPORTANT:
 * Google Drive can behave differently depending on whether
 * the file is in My Drive / Shared Drive.
 *
 * We therefore use the safest available method and log errors.
 * ============================================================
 */
function moveMSERawFileToFolder_(
  fileId,
  destinationFolder
) {

  const file =
    DriveApp.getFileById(
      fileId
    );


  if (
    !destinationFolder
  ) {

    throw new Error(
      "Destination RAW folder is missing."
    );

  }


  try {

    /**
     * --------------------------------------------------------
     * Add to destination.
     * --------------------------------------------------------
     */

    destinationFolder
      .addFile(
        file
      );


    return {

      success:
        true,

      fileId:
        fileId,

      destinationFolderId:
        destinationFolder.getId()

    };

  } catch (error) {

    return {

      success:
        false,

      fileId:
        fileId,

      destinationFolderId:
        destinationFolder.getId(),

      error:
        error.message

    };

  }

}


/**
 * ============================================================
 * GET ACTIVE RAW FILES
 * ============================================================
 */
function getActiveMSERawFiles() {

  const sheet =
    getMSERawCatalogSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return [];

  }


  const columnCount =
    Math.max(
      16,
      sheet.getLastColumn()
    );


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        columnCount
      )
      .getValues();


  const results = [];


  data.forEach(
    row => {

      const active =
        row[14];


      const isActive =
        active === true ||
        String(
          active
        ).toUpperCase() ===
          "TRUE";


      if (
        !isActive
      ) {

        return;

      }


      results.push({

        rawFileId:
          row[0],

        fileName:
          row[1],

        fileId:
          row[2],

        fileType:
          row[3],

        periodFrom:
          row[4],

        periodTo:
          row[5],

        totalRows:
          row[6],

        fileSizeMB:
          row[7],

        status:
          row[8],

        uploadedBy:
          row[9],

        uploadedTimestamp:
          row[10],

        processedTimestamp:
          row[11],

        errorMessage:
          row[12],

        version:
          row[13],

        active:
          row[14],

        notes:
          row[15]

      });

    }
  );


  return results;

}


/**
 * ============================================================
 * GET RAW STORAGE SUMMARY
 * ============================================================
 */
function getMSERawStorageSummary() {

  const storage =
    getMSERawStorage_();


  const activeFiles =
    getActiveMSERawFiles();


  return {

    success:
      true,

    folders: {

      active: {

        id:
          storage.active.getId(),

        name:
          storage.active.getName()

      },

      archive: {

        id:
          storage.archive.getId(),

        name:
          storage.archive.getName()

      },

      error: {

        id:
          storage.error.getId(),

        name:
          storage.error.getName()

      },

      temp: {

        id:
          storage.temp.getId(),

        name:
          storage.temp.getName()

      }

    },

    activeFiles:
      activeFiles

  };

}


/**
 * ============================================================
 * TEST RAW STORAGE
 * ============================================================
 *
 * Run this FIRST.
 *
 * ============================================================
 */
function testMSERawStorage() {

  const storage =
    getMSERawStorage_();


  const result = {

    success:
      true,

    active: {

      id:
        storage.active.getId(),

      name:
        storage.active.getName()

    },

    archive: {

      id:
        storage.archive.getId(),

      name:
        storage.archive.getName()

    },

    error: {

      id:
        storage.error.getId(),

      name:
        storage.error.getName()

    },

    temp: {

      id:
        storage.temp.getId(),

      name:
        storage.temp.getName()

    }

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
 * TEST TEMP SCAN
 * ============================================================
 *
 * Run this AFTER testMSERawStorage().
 *
 * ============================================================
 */
function testMSERawTempScan() {

  const result =
    scanMSERawTempFolder();


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
 * TEST REGISTER ALL TEMP FILES
 * ============================================================
 *
 * Run this AFTER placing a sample Excel file inside TEMP.
 *
 * ============================================================
 */
function testRegisterAllMSERawTempFiles() {

  const result =
    registerAllMSERawTempFiles();


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}
