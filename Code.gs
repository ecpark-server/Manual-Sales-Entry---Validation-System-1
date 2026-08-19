/**
 * ============================================================
 * MSE VALIDATOR PORTAL
 * Manual Sales Entry - Validation System
 * ============================================================
 *
 * ARCHITECTURE
 *
 * Code.gs
 * CSS.html
 * Components.html
 * JS.html
 * Index.html
 * NewRequest.html
 * NewRequestJS.html
 *
 * DATABASE SHEETS
 *
 * USERS
 * PERMISSIONS
 * ROLE_PERMISSIONS
 * USER_PERMISSIONS
 *
 * MSE REQUESTS
 * MSE TICKETS
 * MSE TRANSACTIONS
 *
 * ============================================================
 */


/* ============================================================
   CONFIGURATION
   ============================================================ */

const MSE_CONFIG = {

  PROJECT_TITLE:
    "Manual Sales Entry - Validation System",

  PORTAL_NAME:
    "MSE Validator Portal",

  VERSION:
    "1.0.0",

  YEAR:
    "2026",

  PROJECT_LOGO_FILE_ID:
    "1-v3k3loSqVnnAtFeZ4j5FBd0KYUsDzNs",

  PROJECT_LOGO_FILE_NAME:
    "MSE_LOGO.png",

  FOOTER_LOGO_FILE_ID:
    "1hhld8DCiAbS_86k5TDRvjIS0fQHAety0",

  FOOTER_LOGO_FILE_NAME:
    "company_logo.png",

  /* ----------------------------------------------------------
     DATABASE
     ---------------------------------------------------------- */

  USER_SHEET:
    "USERS",

  PERMISSION_SHEET:
    "PERMISSIONS",

  ROLE_PERMISSION_SHEET:
    "ROLE_PERMISSIONS",

  USER_PERMISSION_SHEET:
    "USER_PERMISSIONS",

  REQUEST_SHEET:
    "MSE REQUESTS",

  TICKET_SHEET:
    "MSE TICKETS",

  TRANSACTION_SHEET:
    "MSE TRANSACTIONS",

  /* ----------------------------------------------------------
     SESSION
     ---------------------------------------------------------- */

  SESSION_PREFIX:
    "MSE_SESSION_",

  SESSION_SECONDS:
    21600, // 6 hours

  /* ----------------------------------------------------------
     SECURITY
     ---------------------------------------------------------- */

  PASSWORD_HASH_ALGORITHM:
    Utilities.DigestAlgorithm.SHA_256

};


/* ============================================================
   WEB APP ENTRY POINT
   ============================================================ */

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle(MSE_CONFIG.PORTAL_NAME)
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* ============================================================
   INCLUDE HTML
   ============================================================ */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* ============================================================
   PROJECT LOGO
   ============================================================ */

function getProjectLogo() {

  return getDriveImageAsDataUri_(
    MSE_CONFIG.PROJECT_LOGO_FILE_ID
  );

}


/* ============================================================
   FOOTER LOGO
   ============================================================ */

function getFooterLogo() {

  return getDriveImageAsDataUri_(
    MSE_CONFIG.FOOTER_LOGO_FILE_ID
  );

}


/* ============================================================
   DRIVE IMAGE → DATA URI
   ============================================================ */

function getDriveImageAsDataUri_(fileId) {

  try {

    const file =
      DriveApp.getFileById(fileId);

    const blob =
      file.getBlob();

    const contentType =
      blob.getContentType();

    const bytes =
      blob.getBytes();

    const base64 =
      Utilities.base64Encode(bytes);

    return `data:${contentType};base64,${base64}`;

  } catch (error) {

    console.error(
      "Unable to load image:",
      error
    );

    return "";

  }

}


/* ============================================================
   PORTAL CONFIGURATION
   ============================================================ */

function getPortalConfig() {

  return {

    projectTitle:
      MSE_CONFIG.PROJECT_TITLE,

    portalName:
      MSE_CONFIG.PORTAL_NAME,

    version:
      MSE_CONFIG.VERSION,

    year:
      MSE_CONFIG.YEAR,

    projectLogo:
      getProjectLogo(),

    footerLogo:
      getFooterLogo()

  };

}


/* ============================================================
   DATABASE SHEET HELPER
   ============================================================ */

function getMseSheet_(
  sheetName,
  headers
) {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {

    throw new Error(
      "Unable to access the active spreadsheet."
    );

  }

  let sheet =
    spreadsheet.getSheetByName(
      sheetName
    );

  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(
        sheetName
      );

  }

  if (
    sheet.getLastRow() === 0
  ) {

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

    sheet.setFrozenRows(1);

  } else {

    const currentHeaders =
      sheet
        .getRange(
          1,
          1,
          1,
          Math.max(
            sheet.getLastColumn(),
            headers.length
          )
        )
        .getValues()[0];

    let needsHeaderUpdate =
      false;

    for (
      let i = 0;
      i < headers.length;
      i++
    ) {

      if (
        currentHeaders[i] !==
        headers[i]
      ) {

        needsHeaderUpdate =
          true;

        break;

      }

    }

    if (needsHeaderUpdate) {

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

    }

  }

  return sheet;

}


/* ============================================================
   DATABASE INITIALIZATION
   ============================================================ */

function initializeMseDatabase() {

  const userSheet =
    getMseSheet_(
      MSE_CONFIG.USER_SHEET,
      [
        "User ID",
        "Full Name",
        "Email",
        "Username",
        "Password Hash",
        "Role",
        "Status",
        "Access Level",
        "Created Date",
        "Last Login"
      ]
    );


  const permissionSheet =
    getMseSheet_(
      MSE_CONFIG.PERMISSION_SHEET,
      [
        "PERMISSION_ID",
        "PERMISSION_CODE",
        "MODULE",
        "ACTION",
        "DESCRIPTION"
      ]
    );


  const rolePermissionSheet =
    getMseSheet_(
      MSE_CONFIG.ROLE_PERMISSION_SHEET,
      [
        "ROLE_ID",
        "PERMISSION_ID",
        "ALLOWED"
      ]
    );


  const userPermissionSheet =
    getMseSheet_(
      MSE_CONFIG.USER_PERMISSION_SHEET,
      [
        "USER_ID",
        "PERMISSION_ID",
        "ALLOWED",
        "Updated Date",
        "Updated By"
      ]
    );


  const requestSheet =
    getMseSheet_(
      MSE_CONFIG.REQUEST_SHEET,
      [
        "Request ID",
        "Prepared By",
        "Prepared Date",
        "Status",
        "Submitted Date",
        "Ticket Count",
        "Transaction Count",
        "Created Timestamp"
      ]
    );


  const ticketSheet =
    getMseSheet_(
      MSE_CONFIG.TICKET_SHEET,
      [
        "Request ID",
        "Ticket #",
        "Ticket ID",
        "DOR",
        "Store",
        "Issue",
        "Remarks",
        "Transaction Count",
        "Created Timestamp"
      ]
    );


  const transactionSheet =
    getMseSheet_(
      MSE_CONFIG.TRANSACTION_SHEET,
      [
        "Request ID",
        "Ticket #",
        "Transaction #",
        "Card Number",
        "Created Timestamp"
      ]
    );


  formatMseHeader_(userSheet);
  formatMseHeader_(permissionSheet);
  formatMseHeader_(rolePermissionSheet);
  formatMseHeader_(userPermissionSheet);
  formatMseHeader_(requestSheet);
  formatMseHeader_(ticketSheet);
  formatMseHeader_(transactionSheet);


  return {

    success: true,

    sheets: {

      users:
        userSheet.getName(),

      permissions:
        permissionSheet.getName(),

      rolePermissions:
        rolePermissionSheet.getName(),

      userPermissions:
        userPermissionSheet.getName(),

      requests:
        requestSheet.getName(),

      tickets:
        ticketSheet.getName(),

      transactions:
        transactionSheet.getName()

    }

  };

}


/* ============================================================
   FORMAT HEADER
   ============================================================ */

function formatMseHeader_(sheet) {

  if (
    !sheet ||
    sheet.getLastColumn() === 0
  ) {

    return;

  }

  sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .setFontWeight(
      "bold"
    );

  sheet.setFrozenRows(1);

}


/* ============================================================
   AUTHENTICATION
   ============================================================ */

/**
 * Generate SHA-256 password hash.
 *
 * Use this helper to generate a hash for a password.
 *
 * Example:
 *
 * generatePasswordHash("AdminPassword123")
 */
function generatePasswordHash(password) {

  if (
    password === null ||
    password === undefined
  ) {

    throw new Error(
      "Password is required."
    );

  }

  return hashPassword_(
    String(password)
  );

}


/* ============================================================
   HASH PASSWORD
   ============================================================ */

function hashPassword_(password) {

  const bytes =
    Utilities.computeDigest(
      MSE_CONFIG.PASSWORD_HASH_ALGORITHM,
      String(password),
      Utilities.Charset.UTF_8
    );

  return bytes
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;

      return value
        .toString(16)
        .padStart(2, "0");

    })
    .join("");

}


/* ============================================================
   NORMALIZE USERNAME
   ============================================================ */

function normalizeUsername_(value) {

  return String(
    value || ""
  )
    .trim()
    .toLowerCase();

}


/* ============================================================
   FIND USER
   ============================================================ */

function findMseUser_(username) {

  const normalizedUsername =
    normalizeUsername_(
      username
    );

  if (!normalizedUsername) {

    return null;

  }

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_SHEET
    );

  if (!sheet) {

    throw new Error(
      "USERS sheet has not been initialized."
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return null;

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        10
      )
      .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    if (
      normalizeUsername_(
        row[3]
      ) === normalizedUsername
    ) {

      return {

        rowNumber:
          i + 2,

        userId:
          String(row[0] || "").trim(),

        fullName:
          String(row[1] || "").trim(),

        email:
          String(row[2] || "").trim(),

        username:
          String(row[3] || "").trim(),

        passwordHash:
          String(row[4] || "").trim(),

        role:
          String(row[5] || "").trim().toUpperCase(),

        status:
          String(row[6] || "").trim().toUpperCase(),

        accessLevel:
          String(row[7] || "").trim().toUpperCase(),

        createdDate:
          row[8],

        lastLogin:
          row[9]

      };

    }

  }

  return null;

}


/* ============================================================
   AUTHENTICATE USER
   ============================================================ */

function authenticateMseUser(
  username,
  password
) {

  const user =
    findMseUser_(
      username
    );

  if (!user) {

    throw new Error(
      "Invalid username or password."
    );

  }

  if (
    user.status !== "ACTIVE"
  ) {

    throw new Error(
      "This user account is disabled."
    );

  }

  if (
    !user.passwordHash
  ) {

    throw new Error(
      "This user account does not have a configured password."
    );

  }

  const suppliedHash =
    hashPassword_(
      String(password || "")
    );

  if (
    suppliedHash !==
    user.passwordHash
  ) {

    throw new Error(
      "Invalid username or password."
    );

  }

  const session =
    createMseSession_(
      user
    );

  updateMseLastLogin_(
    user.rowNumber
  );

  return {

    success: true,

    sessionToken:
      session.token,

    expiresIn:
      MSE_CONFIG.SESSION_SECONDS,

    user:
      getSafeUserObject_(
        user
      ),

    permissions:
      getUserPermissions_(
        user
      ),

    modules:
      getUserModules_(
        user
      )

  };

}


/* ============================================================
   SAFE USER OBJECT
   ============================================================ */

function getSafeUserObject_(user) {

  return {

    userId:
      user.userId,

    fullName:
      user.fullName,

    email:
      user.email,

    username:
      user.username,

    role:
      user.role,

    status:
      user.status,

    accessLevel:
      user.accessLevel,

    initials:
      getInitials_(
        user.fullName
      )

  };

}


/* ============================================================
   INITIALS
   ============================================================ */

function getInitials_(name) {

  const value =
    String(
      name || ""
    )
      .trim();

  if (!value) {

    return "U";

  }

  const parts =
    value
      .split(/\s+/)
      .filter(Boolean);

  if (
    parts.length === 1
  ) {

    return parts[0]
      .substring(0, 1)
      .toUpperCase();

  }

  return (
    parts[0].substring(0, 1) +
    parts[parts.length - 1]
      .substring(0, 1)
  )
    .toUpperCase();

}


/* ============================================================
   UPDATE LAST LOGIN
   ============================================================ */

function updateMseLastLogin_(
  rowNumber
) {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_SHEET
    );

  if (!sheet) {

    return;

  }

  sheet
    .getRange(
      rowNumber,
      10
    )
    .setValue(
      new Date()
    );

}


/* ============================================================
   SESSION
   ============================================================ */

function createMseSession_(user) {

  const token =
    Utilities.getUuid() +
    "-" +
    Utilities.getUuid();

  const sessionData = {

    userId:
      user.userId,

    username:
      user.username,

    created:
      Date.now(),

    expires:
      Date.now() +
      (
        MSE_CONFIG.SESSION_SECONDS *
        1000
      )

  };

  CacheService
    .getScriptCache()
    .put(
      MSE_CONFIG.SESSION_PREFIX + token,
      JSON.stringify(sessionData),
      MSE_CONFIG.SESSION_SECONDS
    );

  return {

    token:
      token,

    expires:
      sessionData.expires

  };

}


/* ============================================================
   GET SESSION
   ============================================================ */

function getMseSession_(token) {

  const value =
    String(
      token || ""
    ).trim();

  if (!value) {

    return null;

  }

  const cached =
    CacheService
      .getScriptCache()
      .get(
        MSE_CONFIG.SESSION_PREFIX +
        value
      );

  if (!cached) {

    return null;

  }

  try {

    const session =
      JSON.parse(
        cached
      );

    if (
      session.expires &&
      Date.now() >
      Number(session.expires)
    ) {

      CacheService
        .getScriptCache()
        .remove(
          MSE_CONFIG.SESSION_PREFIX +
          value
        );

      return null;

    }

    return session;

  } catch (error) {

    return null;

  }

}


/* ============================================================
   GET AUTHENTICATED USER
   ============================================================ */

function getAuthenticatedMseUser(
  sessionToken
) {

  const session =
    getMseSession_(
      sessionToken
    );

  if (!session) {

    throw new Error(
      "Your session has expired. Please log in again."
    );

  }

  const user =
    findMseUser_(
      session.username
    );

  if (!user) {

    revokeMseSession_(
      sessionToken
    );

    throw new Error(
      "User account no longer exists."
    );

  }

  if (
    user.status !== "ACTIVE"
  ) {

    revokeMseSession_(
      sessionToken
    );

    throw new Error(
      "Your user account has been disabled."
    );

  }

  return user;

}


/* ============================================================
   SESSION CHECK
   ============================================================ */

function getMseSessionInfo(
  sessionToken
) {

  const user =
    getAuthenticatedMseUser(
      sessionToken
    );

  return {

    success: true,

    authenticated: true,

    user:
      getSafeUserObject_(
        user
      ),

    permissions:
      getUserPermissions_(
        user
      ),

    modules:
      getUserModules_(
        user
      )

  };

}


/* ============================================================
   LOGOUT
   ============================================================ */

function logoutMseUser(
  sessionToken
) {

  if (
    sessionToken
  ) {

    revokeMseSession_(
      sessionToken
    );

  }

  return {

    success: true

  };

}


/* ============================================================
   REVOKE SESSION
   ============================================================ */

function revokeMseSession_(
  sessionToken
) {

  if (!sessionToken) {

    return;

  }

  CacheService
    .getScriptCache()
    .remove(
      MSE_CONFIG.SESSION_PREFIX +
      sessionToken
    );

}


/* ============================================================
   PERMISSIONS
   ============================================================ */

/**
 * Returns role permissions plus user-specific overrides.
 *
 * USER_PERMISSIONS has priority over ROLE_PERMISSIONS.
 *
 * TRUE  = allowed
 * FALSE = denied
 */

function getUserPermissions_(user) {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const permissionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.PERMISSION_SHEET
    );

  const rolePermissionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.ROLE_PERMISSION_SHEET
    );

  const userPermissionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_PERMISSION_SHEET
    );

  if (
    !permissionSheet ||
    !rolePermissionSheet
  ) {

    throw new Error(
      "Permission database has not been initialized."
    );

  }

  const permissions =
    permissionSheet.getLastRow() >= 2
      ? permissionSheet
          .getRange(
            2,
            1,
            permissionSheet.getLastRow() - 1,
            5
          )
          .getValues()
      : [];

  const rolePermissions =
    rolePermissionSheet.getLastRow() >= 2
      ? rolePermissionSheet
          .getRange(
            2,
            1,
            rolePermissionSheet.getLastRow() - 1,
            3
          )
          .getValues()
      : [];

  const userPermissions =
    userPermissionSheet &&
    userPermissionSheet.getLastRow() >= 2
      ? userPermissionSheet
          .getRange(
            2,
            1,
            userPermissionSheet.getLastRow() - 1,
            5
          )
          .getValues()
      : [];


  const result = {};

  /* ----------------------------------------------------------
     ROLE PERMISSIONS
     ---------------------------------------------------------- */

  rolePermissions.forEach(
    function(row) {

      const roleId =
        String(
          row[0] || ""
        )
          .trim()
          .toUpperCase();

      const permissionId =
        String(
          row[1] || ""
        )
          .trim();

      const allowed =
        normalizeBoolean_(
          row[2]
        );

      if (
        roleId ===
        getRoleId_(user.role)
      ) {

        result[
          permissionId
        ] = allowed;

      }

    }
  );


  /* ----------------------------------------------------------
     USER-SPECIFIC OVERRIDES
     ---------------------------------------------------------- */

  userPermissions.forEach(
    function(row) {

      const userId =
        String(
          row[0] || ""
        )
          .trim();

      const permissionId =
        String(
          row[1] || ""
        )
          .trim();

      const allowed =
        normalizeBoolean_(
          row[2]
        );

      if (
        userId ===
        user.userId
      ) {

        result[
          permissionId
        ] = allowed;

      }

    }
  );


  /* ----------------------------------------------------------
     BUILD COMPLETE PERMISSION OBJECT
     ---------------------------------------------------------- */

  const finalPermissions = {};

  permissions.forEach(
    function(row) {

      const permissionId =
        String(
          row[0] || ""
        )
          .trim();

      const permissionCode =
        String(
          row[1] || ""
        )
          .trim()
          .toUpperCase();

      const module =
        String(
          row[2] || ""
        )
          .trim();

      const action =
        String(
          row[3] || ""
        )
          .trim()
          .toUpperCase();

      const description =
        String(
          row[4] || ""
        )
          .trim();

      finalPermissions[
        permissionCode
      ] = {

        permissionId:
          permissionId,

        permissionCode:
          permissionCode,

        module:
          module,

        action:
          action,

        description:
          description,

        allowed:
          result.hasOwnProperty(
            permissionId
          )
            ? !!result[permissionId]
            : false

      };

    }
  );


  /*
   * ADMIN OVERRIDE
   *
   * ADMIN receives full permission access.
   */

  if (
    user.role === "ADMIN"
  ) {

    Object.keys(
      finalPermissions
    ).forEach(
      function(code) {

        finalPermissions[
          code
        ].allowed = true;

      }
    );

  }


  return finalPermissions;

}


/* ============================================================
   ROLE ID
   ============================================================ */

function getRoleId_(role) {

  const normalizedRole =
    String(
      role || ""
    )
      .trim()
      .toUpperCase();

  return (
    "ROLE-" +
    normalizedRole
  );

}


/* ============================================================
   BOOLEAN NORMALIZER
   ============================================================ */

function normalizeBoolean_(value) {

  if (
    value === true
  ) {

    return true;

  }

  if (
    value === false
  ) {

    return false;

  }

  const normalized =
    String(
      value || ""
    )
      .trim()
      .toUpperCase();

  return (
    normalized === "TRUE" ||
    normalized === "YES" ||
    normalized === "1" ||
    normalized === "ALLOWED"
  );

}


/* ============================================================
   MODULE ACCESS
   ============================================================ */

function getUserModules_(user) {

  const permissions =
    getUserPermissions_(
      user
    );

  const modules = {

    Dashboard: false,

    Request: false,

    Validation: false,

    Extract: false,

    Reports: false,

    Users: false,

    Roles: false,

    Permissions: false,

    "Processed Data": false,

    Audit: false,

    Configuration: false

  };


  Object.keys(
    permissions
  ).forEach(
    function(code) {

      const permission =
        permissions[code];

      if (
        permission.allowed
      ) {

        const module =
          permission.module;

        if (
          modules.hasOwnProperty(
            module
          )
        ) {

          modules[module] =
            true;

        }

      }

    }
  );


  return modules;

}


/* ============================================================
   SERVER-SIDE PERMISSION CHECK
   ============================================================ */

/**
 * IMPORTANT:
 *
 * Frontend permissions are for UI visibility only.
 *
 * Every protected server function must use
 * requireMsePermission_().
 */

function requireMsePermission_(
  sessionToken,
  permissionCode
) {

  const user =
    getAuthenticatedMseUser(
      sessionToken
    );

  const permissions =
    getUserPermissions_(
      user
    );

  const code =
    String(
      permissionCode || ""
    )
      .trim()
      .toUpperCase();

  if (
    user.role === "ADMIN"
  ) {

    return user;

  }

  if (
    !permissions[code] ||
    !permissions[code].allowed
  ) {

    throw new Error(
      "Access denied. Required permission: " +
      code
    );

  }

  return user;

}


/* ============================================================
   ADMIN CHECK
   ============================================================ */

function requireMseAdmin_(
  sessionToken
) {

  const user =
    getAuthenticatedMseUser(
      sessionToken
    );

  if (
    user.role !== "ADMIN"
  ) {

    throw new Error(
      "Administrator access is required."
    );

  }

  return user;

}


/* ============================================================
   ADMIN - GET USERS
   ============================================================ */

function getMseUsers(
  sessionToken
) {

  requireMsePermission_(
    sessionToken,
    "USER_VIEW"
  );

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_SHEET
    );

  if (!sheet) {

    throw new Error(
      "USERS sheet has not been initialized."
    );

  }

  if (
    sheet.getLastRow() < 2
  ) {

    return {

      success: true,

      users: []

    };

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        10
      )
      .getValues();

  const users =
    values.map(
      function(row) {

        return {

          userId:
            row[0],

          fullName:
            row[1],

          email:
            row[2],

          username:
            row[3],

          role:
            row[5],

          status:
            row[6],

          accessLevel:
            row[7],

          createdDate:
            row[8],

          lastLogin:
            row[9]

        };

      }
    );

  return {

    success: true,

    users: users

  };

}


/* ============================================================
   ADMIN - UPDATE USER STATUS
   ============================================================ */

function updateMseUserStatus(
  sessionToken,
  userId,
  status
) {

  requireMsePermission_(
    sessionToken,
    "USER_DISABLE"
  );

  const normalizedStatus =
    String(
      status || ""
    )
      .trim()
      .toUpperCase();

  if (
    normalizedStatus !== "ACTIVE" &&
    normalizedStatus !== "DISABLED"
  ) {

    throw new Error(
      "Invalid user status."
    );

  }

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_SHEET
    );

  if (!sheet) {

    throw new Error(
      "USERS sheet not found."
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {

    throw new Error(
      "No users found."
    );

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        10
      )
      .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][0]
      ).trim() ===
      String(
        userId
      ).trim()
    ) {

      sheet
        .getRange(
          i + 2,
          7
        )
        .setValue(
          normalizedStatus
        );

      return {

        success: true,

        userId:
          userId,

        status:
          normalizedStatus

      };

    }

  }

  throw new Error(
    "User not found."
  );

}


/* ============================================================
   ADMIN - SET USER PERMISSION
   ============================================================ */

function setMseUserPermission(
  sessionToken,
  userId,
  permissionId,
  allowed
) {

  const admin =
    requireMseAdmin_(
      sessionToken
    );

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getMseSheet_(
      MSE_CONFIG.USER_PERMISSION_SHEET,
      [
        "USER_ID",
        "PERMISSION_ID",
        "ALLOWED",
        "Updated Date",
        "Updated By"
      ]
    );


  const targetUserId =
    String(
      userId || ""
    )
      .trim();

  const targetPermissionId =
    String(
      permissionId || ""
    )
      .trim();

  if (
    !targetUserId ||
    !targetPermissionId
  ) {

    throw new Error(
      "User ID and Permission ID are required."
    );

  }


  const lastRow =
    sheet.getLastRow();

  if (
    lastRow >= 2
  ) {

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          5
        )
        .getValues();

    for (
      let i = 0;
      i < values.length;
      i++
    ) {

      if (
        String(
          values[i][0]
        ).trim() ===
        targetUserId &&
        String(
          values[i][1]
        ).trim() ===
        targetPermissionId
      ) {

        sheet
          .getRange(
            i + 2,
            3
          )
          .setValue(
            !!allowed
          );

        sheet
          .getRange(
            i + 2,
            4
          )
          .setValue(
            new Date()
          );

        sheet
          .getRange(
            i + 2,
            5
          )
          .setValue(
            admin.username
          );

        return {

          success: true,

          action: "UPDATED",

          userId:
            targetUserId,

          permissionId:
            targetPermissionId,

          allowed:
            !!allowed

        };

      }

    }

  }


  sheet.appendRow([

    targetUserId,

    targetPermissionId,

    !!allowed,

    new Date(),

    admin.username

  ]);


  return {

    success: true,

    action: "CREATED",

    userId:
      targetUserId,

    permissionId:
      targetPermissionId,

    allowed:
      !!allowed

  };

}


/* ============================================================
   ADMIN - GET USER PERMISSIONS
   ============================================================ */

function getMseUserPermissionMatrix(
  sessionToken,
  userId
) {

  requireMseAdmin_(
    sessionToken
  );

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const userSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_SHEET
    );

  const permissionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.PERMISSION_SHEET
    );

  const rolePermissionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.ROLE_PERMISSION_SHEET
    );

  const userPermissionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_PERMISSION_SHEET
    );

  if (
    !userSheet ||
    !permissionSheet ||
    !rolePermissionSheet
  ) {

    throw new Error(
      "Authorization database is incomplete."
    );

  }


  const targetUser =
    findUserById_(
      userId
    );

  if (!targetUser) {

    throw new Error(
      "User not found."
    );

  }


  const permissions =
    permissionSheet.getLastRow() >= 2
      ? permissionSheet
          .getRange(
            2,
            1,
            permissionSheet.getLastRow() - 1,
            5
          )
          .getValues()
      : [];


  const rolePermissions =
    rolePermissionSheet.getLastRow() >= 2
      ? rolePermissionSheet
          .getRange(
            2,
            1,
            rolePermissionSheet.getLastRow() - 1,
            3
          )
          .getValues()
      : [];


  const userPermissions =
    userPermissionSheet &&
    userPermissionSheet.getLastRow() >= 2
      ? userPermissionSheet
          .getRange(
            2,
            1,
            userPermissionSheet.getLastRow() - 1,
            5
          )
          .getValues()
      : [];


  const roleDefaults = {};


  rolePermissions.forEach(
    function(row) {

      if (
        String(
          row[0] || ""
        )
          .trim()
          .toUpperCase() ===
        getRoleId_(
          targetUser.role
        )
      ) {

        roleDefaults[
          String(
            row[1] || ""
          ).trim()
        ] =
          normalizeBoolean_(
            row[2]
          );

      }

    }
  );


  const overrides = {};


  userPermissions.forEach(
    function(row) {

      if (
        String(
          row[0] || ""
        ).trim() ===
        targetUser.userId
      ) {

        overrides[
          String(
            row[1] || ""
          ).trim()
        ] =
          normalizeBoolean_(
            row[2]
          );

      }

    }
  );


  const matrix =
    permissions.map(
      function(row) {

        const permissionId =
          String(
            row[0] || ""
          ).trim();

        const permissionCode =
          String(
            row[1] || ""
          ).trim();

        const roleAllowed =
          roleDefaults.hasOwnProperty(
            permissionId
          )
            ? roleDefaults[permissionId]
            : false;

        const hasOverride =
          overrides.hasOwnProperty(
            permissionId
          );

        const effectiveAllowed =
          hasOverride
            ? overrides[permissionId]
            : roleAllowed;

        return {

          permissionId:
            permissionId,

          permissionCode:
            permissionCode,

          module:
            row[2],

          action:
            row[3],

          description:
            row[4],

          roleAllowed:
            roleAllowed,

          hasOverride:
            hasOverride,

          allowed:
            targetUser.role === "ADMIN"
              ? true
              : effectiveAllowed

        };

      }
    );


  return {

    success: true,

    user:
      getSafeUserObject_(
        targetUser
      ),

    permissions:
      matrix

  };

}


/* ============================================================
   FIND USER BY ID
   ============================================================ */

function findUserById_(
  userId
) {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.USER_SHEET
    );

  if (!sheet) {

    return null;

  }

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {

    return null;

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        10
      )
      .getValues();

  const target =
    String(
      userId || ""
    ).trim();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][0]
      ).trim() === target
    ) {

      return {

        rowNumber:
          i + 2,

        userId:
          String(values[i][0] || "").trim(),

        fullName:
          String(values[i][1] || "").trim(),

        email:
          String(values[i][2] || "").trim(),

        username:
          String(values[i][3] || "").trim(),

        passwordHash:
          String(values[i][4] || "").trim(),

        role:
          String(values[i][5] || "")
            .trim()
            .toUpperCase(),

        status:
          String(values[i][6] || "")
            .trim()
            .toUpperCase(),

        accessLevel:
          String(values[i][7] || "")
            .trim()
            .toUpperCase(),

        createdDate:
          values[i][8],

        lastLogin:
          values[i][9]

      };

    }

  }

  return null;

}


/* ============================================================
   REQUEST ID
   ============================================================ */

function generateServerRequestId_() {

  const now =
    new Date();

  const timezone =
    Session.getScriptTimeZone() ||
    "Asia/Manila";

  const datePart =
    Utilities.formatDate(
      now,
      timezone,
      "yyyyMMdd"
    );

  const timePart =
    Utilities.formatDate(
      now,
      timezone,
      "HHmmss"
    );

  const randomPart =
    String(
      Math.floor(
        1000 +
        Math.random() * 9000
      )
    );

  return (
    "MSE-" +
    datePart +
    "-" +
    timePart +
    "-" +
    randomPart
  );

}


/* ============================================================
   TICKET VALIDATION
   ============================================================ */

function isValidMseTicketId_(
  ticketId
) {

  const value =
    String(
      ticketId || ""
    )
      .trim();

  if (
    !/^\d{10}$/.test(
      value
    )
  ) {

    return false;

  }

  const year =
    Number(
      value.substring(
        0,
        2
      )
    );

  const month =
    Number(
      value.substring(
        2,
        4
      )
    );

  const day =
    Number(
      value.substring(
        4,
        6
      )
    );

  if (
    month < 1 ||
    month > 12
  ) {

    return false;

  }

  const fullYear =
    2000 + year;

  const testDate =
    new Date(
      fullYear,
      month - 1,
      day
    );

  return (

    testDate.getFullYear() ===
      fullYear &&

    testDate.getMonth() ===
      month - 1 &&

    testDate.getDate() ===
      day

  );

}


/* ============================================================
   EXTRACT DOR
   ============================================================ */

function extractMseDor_(
  ticketId
) {

  const value =
    String(
      ticketId || ""
    )
      .trim();

  if (
    !isValidMseTicketId_(
      value
    )
  ) {

    return "";

  }

  const year =
    2000 +
    Number(
      value.substring(
        0,
        2
      )
    );

  const month =
    Number(
      value.substring(
        2,
        4
      )
    );

  const day =
    Number(
      value.substring(
        4,
        6
      )
    );

  return (
    String(month).padStart(
      2,
      "0"
    ) +
    "/" +
    String(day).padStart(
      2,
      "0"
    ) +
    "/" +
    String(year)
  );

}


/* ============================================================
   CARD NORMALIZATION
   ============================================================ */

function normalizeMseCardNumber_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }

  let cleanVal =
    String(value)
      .replace(
        /[^0-9]/g,
        ""
      );

  if (
    cleanVal === ""
  ) {

    return "";

  }


  /* 8000 BPC */

  if (
    cleanVal.startsWith(
      "8000"
    )
  ) {

    if (
      cleanVal.length >= 8
    ) {

      return cleanVal.slice(
        -8
      );

    }

    return cleanVal;

  }


  /* 7000 BPC */

  if (
    cleanVal.startsWith(
      "7000"
    )
  ) {

    return normalizeMseSixteenDigitCard_(
      cleanVal
    );

  }


  /* 5000 RPC */

  if (
    cleanVal.startsWith(
      "5000"
    )
  ) {

    return normalizeMseSixteenDigitCard_(
      cleanVal
    );

  }


  return normalizeMseBaselineCard_(
    cleanVal
  );

}


/* ============================================================
   BASELINE CARD NORMALIZATION
   ============================================================ */

function normalizeMseBaselineCard_(
  cleanVal
) {

  const len =
    cleanVal.length;

  if (
    len > 16
  ) {

    return (
      cleanVal.substring(
        0,
        1
      ) +
      cleanVal.substring(
        len - 15
      )
    );

  }

  if (
    len >= 11 &&
    len <= 15
  ) {

    return (
      cleanVal.substring(
        0,
        1
      ) +
      "0" +
      cleanVal.substring(
        1
      )
    );

  }

  return cleanVal;

}


/* ============================================================
   SIXTEEN DIGIT CARD NORMALIZATION
   ============================================================ */

function normalizeMseSixteenDigitCard_(
  cleanVal
) {

  const normalized =
    normalizeMseBaselineCard_(
      cleanVal
    );

  if (
    normalized.length === 16
  ) {

    return normalized;

  }

  if (
    normalized.length > 16
  ) {

    return (
      normalized.substring(
        0,
        1
      ) +
      normalized.substring(
        normalized.length - 15
      )
    );

  }

  return normalized;

}


/* ============================================================
   SERVER REQUEST VALIDATION
   ============================================================ */

function validateMseRequestServer_(
  requestData
) {

  if (
    !requestData ||
    typeof requestData !== "object"
  ) {

    throw new Error(
      "Invalid request payload."
    );

  }

  const tickets =
    Array.isArray(
      requestData.tickets
    )
      ?
      requestData.tickets
      :
      [];

  if (
    tickets.length === 0
  ) {

    throw new Error(
      "At least one ticket is required."
    );

  }

  let transactionCount =
    0;


  for (
    let i = 0;
    i < tickets.length;
    i++
  ) {

    const ticket =
      tickets[i];

    if (
      !ticket ||
      typeof ticket !== "object"
    ) {

      throw new Error(
        "Ticket #" +
        (i + 1) +
        " contains invalid data."
      );

    }

    const ticketId =
      String(
        ticket.ticketId || ""
      )
      .trim();

    if (
      !isValidMseTicketId_(
        ticketId
      )
    ) {

      throw new Error(
        "Ticket #" +
        (i + 1) +
        " has an invalid Ticket ID."
      );

    }

    if (
      !ticket.store
    ) {

      throw new Error(
        "Please select the Store for Ticket #" +
        (i + 1) +
        "."
      );

    }

    if (
      !ticket.issue
    ) {

      throw new Error(
        "Please select the Issue for Ticket #" +
        (i + 1) +
        "."
      );

    }

    const transactions =
      Array.isArray(
        ticket.transactions
      )
        ?
        ticket.transactions
        :
        [];


    for (
      let j = 0;
      j < transactions.length;
      j++
    ) {

      const transaction =
        transactions[j];

      const cardNumber =
        normalizeMseCardNumber_(
          transaction
            ? transaction.cardNumber
            : ""
        );

      if (
        !cardNumber
      ) {

        throw new Error(
          "Ticket #" +
          (i + 1) +
          ", Transaction #" +
          (j + 1) +
          " requires a Card Number."
        );

      }

      transactionCount++;

    }

  }


  return {

    ticketCount:
      tickets.length,

    transactionCount:
      transactionCount

  };

}


/* ============================================================
   SUBMIT MSE REQUEST
   ============================================================ */

function submitMSERequest(
  requestData
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(
    30000
  );

  try {

    /*
     * --------------------------------------------------------
     * AUTHORIZATION
     *
     * NewRequestJS.html should send:
     *
     * requestData.sessionToken
     * --------------------------------------------------------
     */

    const sessionToken =
      requestData
        ? requestData.sessionToken
        : "";

    const user =
      requireMsePermission_(
        sessionToken,
        "REQUEST_SUBMIT"
      );


    /*
     * --------------------------------------------------------
     * VALIDATE
     * --------------------------------------------------------
     */

    const validation =
      validateMseRequestServer_(
        requestData
      );


    /*
     * --------------------------------------------------------
     * DATABASE
     * --------------------------------------------------------
     */

    const requestSheet =
      getMseSheet_(
        MSE_CONFIG.REQUEST_SHEET,
        [
          "Request ID",
          "Prepared By",
          "Prepared Date",
          "Status",
          "Submitted Date",
          "Ticket Count",
          "Transaction Count",
          "Created Timestamp"
        ]
      );


    const ticketSheet =
      getMseSheet_(
        MSE_CONFIG.TICKET_SHEET,
        [
          "Request ID",
          "Ticket #",
          "Ticket ID",
          "DOR",
          "Store",
          "Issue",
          "Remarks",
          "Transaction Count",
          "Created Timestamp"
        ]
      );


    const transactionSheet =
      getMseSheet_(
        MSE_CONFIG.TRANSACTION_SHEET,
        [
          "Request ID",
          "Ticket #",
          "Transaction #",
          "Card Number",
          "Created Timestamp"
        ]
      );


    /*
     * --------------------------------------------------------
     * REQUEST ID
     * --------------------------------------------------------
     */

    const requestId =
      generateServerRequestId_();

    const now =
      new Date();

    const timezone =
      Session.getScriptTimeZone() ||
      "Asia/Manila";

    const submittedDate =
      Utilities.formatDate(
        now,
        timezone,
        "MM/dd/yyyy HH:mm:ss"
      );


    /*
     * --------------------------------------------------------
     * PREPARED INFORMATION
     *
     * Use authenticated user instead of trusting
     * the browser-provided user.
     * --------------------------------------------------------
     */

    const preparedBy =
      user.fullName ||
      user.username;

    const preparedDate =
      String(
        requestData.preparedDate || ""
      )
      .trim();


    /*
     * --------------------------------------------------------
     * SAVE REQUEST
     * --------------------------------------------------------
     */

    requestSheet.appendRow([

      requestId,

      preparedBy,

      preparedDate,

      "SUBMITTED",

      submittedDate,

      validation.ticketCount,

      validation.transactionCount,

      now

    ]);


    /*
     * --------------------------------------------------------
     * SAVE TICKETS + TRANSACTIONS
     * --------------------------------------------------------
     */

    const tickets =
      requestData.tickets || [];

    const ticketRows =
      [];

    const transactionRows =
      [];


    tickets.forEach(
      function(
        ticket,
        ticketIndex
      ) {

        const ticketNumber =
          ticketIndex + 1;

        const ticketId =
          String(
            ticket.ticketId || ""
          )
          .trim();

        const dor =
          extractMseDor_(
            ticketId
          );

        const store =
          String(
            ticket.store || ""
          )
          .trim();

        const issue =
          String(
            ticket.issue || ""
          )
          .trim();

        const remarks =
          String(
            ticket.remarks || ""
          )
          .trim();

        const transactions =
          Array.isArray(
            ticket.transactions
          )
            ?
            ticket.transactions
            :
            [];


        ticketRows.push([

          requestId,

          ticketNumber,

          ticketId,

          dor,

          store,

          issue,

          remarks,

          transactions.length,

          now

        ]);


        transactions.forEach(
          function(
            transaction,
            transactionIndex
          ) {

            const cardNumber =
              normalizeMseCardNumber_(
                transaction
                  ? transaction.cardNumber
                  : ""
              );


            transactionRows.push([

              requestId,

              ticketNumber,

              transactionIndex + 1,

              cardNumber,

              now

            ]);

          }
        );

      }
    );


    /*
     * --------------------------------------------------------
     * BATCH WRITE TICKETS
     * --------------------------------------------------------
     */

    if (
      ticketRows.length > 0
    ) {

      ticketSheet
        .getRange(
          ticketSheet.getLastRow() + 1,
          1,
          ticketRows.length,
          ticketRows[0].length
        )
        .setValues(
          ticketRows
        );

    }


    /*
     * --------------------------------------------------------
     * BATCH WRITE TRANSACTIONS
     * --------------------------------------------------------
     */

    if (
      transactionRows.length > 0
    ) {

      transactionSheet
        .getRange(
          transactionSheet.getLastRow() + 1,
          1,
          transactionRows.length,
          transactionRows[0].length
        )
        .setValues(
          transactionRows
        );

    }


    return {

      success:
        true,

      requestId:
        requestId,

      status:
        "SUBMITTED",

      submittedDate:
        submittedDate,

      ticketCount:
        validation.ticketCount,

      transactionCount:
        validation.transactionCount,

      preparedBy:
        preparedBy,

      submittedBy:
        user.userId,

      message:
        "MSE Request submitted successfully."

    };


  } catch (error) {

    console.error(
      "submitMSERequest failed:",
      error
    );

    throw new Error(
      error &&
      error.message
        ?
        error.message
        :
        "Unable to submit MSE Request."
    );


  } finally {

    try {

      lock.releaseLock();

    } catch (ignore) {}

  }

}


/* ============================================================
   GET REQUEST BY ID
   ============================================================ */

function getMSERequest(
  requestId,
  sessionToken
) {

  requireMsePermission_(
    sessionToken,
    "REQUEST_VIEW"
  );


  const id =
    String(
      requestId || ""
    )
    .trim();

  if (!id) {

    throw new Error(
      "Request ID is required."
    );

  }


  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const requestSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.REQUEST_SHEET
    );

  const ticketSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.TICKET_SHEET
    );

  const transactionSheet =
    spreadsheet.getSheetByName(
      MSE_CONFIG.TRANSACTION_SHEET
    );


  if (
    !requestSheet ||
    !ticketSheet ||
    !transactionSheet
  ) {

    throw new Error(
      "MSE database sheets have not been initialized."
    );

  }


  const requestData =
    requestSheet
      .getDataRange()
      .getValues();

  let request =
    null;


  for (
    let i = 1;
    i < requestData.length;
    i++
  ) {

    if (
      String(
        requestData[i][0]
      ) === id
    ) {

      request = {

        requestId:
          requestData[i][0],

        preparedBy:
          requestData[i][1],

        preparedDate:
          requestData[i][2],

        status:
          requestData[i][3],

        submittedDate:
          requestData[i][4],

        ticketCount:
          requestData[i][5],

        transactionCount:
          requestData[i][6],

        createdTimestamp:
          requestData[i][7]

      };

      break;

    }

  }


  if (!request) {

    return {

      success:
        false,

      message:
        "Request not found.",

      request:
        null

    };

  }


  const ticketData =
    ticketSheet
      .getDataRange()
      .getValues();

  const tickets =
    [];


  for (
    let i = 1;
    i < ticketData.length;
    i++
  ) {

    if (
      String(
        ticketData[i][0]
      ) !== id
    ) {

      continue;

    }


    tickets.push({

      ticketNumber:
        ticketData[i][1],

      ticketId:
        ticketData[i][2],

      dor:
        ticketData[i][3],

      store:
        ticketData[i][4],

      issue:
        ticketData[i][5],

      remarks:
        ticketData[i][6],

      transactionCount:
        ticketData[i][7],

      createdTimestamp:
        ticketData[i][8],

      transactions:
        []

    });

  }


  const transactionData =
    transactionSheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < transactionData.length;
    i++
  ) {

    if (
      String(
        transactionData[i][0]
      ) !== id
    ) {

      continue;

    }


    const ticketNumber =
      Number(
        transactionData[i][1]
      );


    const transaction = {

      transactionNumber:
        transactionData[i][2],

      cardNumber:
        transactionData[i][3],

      createdTimestamp:
        transactionData[i][4]

    };


    const ticket =
      tickets.find(
        function(item) {

          return Number(
            item.ticketNumber
          ) === ticketNumber;

        }
      );


    if (ticket) {

      ticket.transactions.push(
        transaction
      );

    }

  }


  request.tickets =
    tickets;


  return {

    success:
      true,

    request:
      request

  };

}


/* ============================================================
   DATABASE STATUS
   ============================================================ */

function getMseDatabaseStatus() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {

    return {

      success:
        false,

      message:
        "Active spreadsheet not available."

    };

  }


  return {

    success:
      true,

    userSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.USER_SHEET
      ),

    permissionSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.PERMISSION_SHEET
      ),

    rolePermissionSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.ROLE_PERMISSION_SHEET
      ),

    userPermissionSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.USER_PERMISSION_SHEET
      ),

    requestSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.REQUEST_SHEET
      ),

    ticketSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.TICKET_SHEET
      ),

    transactionSheet:
      !!spreadsheet.getSheetByName(
        MSE_CONFIG.TRANSACTION_SHEET
      ),

    spreadsheetName:
      spreadsheet.getName()

  };

}


/* ============================================================
   SETUP AUTHORIZATION DATABASE
   ============================================================ */

/**
 * Run this once after replacing Code.gs.
 *
 * It creates USER_PERMISSIONS if it does not exist
 * and preserves your existing USERS,
 * PERMISSIONS and ROLE_PERMISSIONS sheets.
 */

function setupMseAuthorization() {

  const result =
    initializeMseDatabase();

  return {

    success:
      true,

    message:
      "MSE authorization database initialized successfully.",

    result:
      result

  };

}
