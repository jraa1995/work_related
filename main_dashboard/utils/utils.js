// ===== UTILITIES
function getProperty(key, defaultValue = null) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value || defaultValue;
  } catch (error) {
    console.error(`Error getting property ${key}:`, error);
    return defaultValue;
  }
}

// SET SCRIPT PROP
function setProperty(key, value) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
    return true;
  } catch (error) {
    console.error(`Error setting property ${key}:`, error);
    return false;
  }
}

// SET MULTIPLE SCRIPT PROP
function setProperties(properties) {
  try {
    PropertiesService.getScriptProperties().setProperties(properties);
    return true;
  } catch (error) {
    console.error("Error setting properties:", error);
    return false;
  }
}

// INITIALIZE SCRIPT PROP W/ DEFAULT VALUES
function initializeProperties() {
  const defaultProperties = {
    ROOT_FOLDER_ID: "",
    SUPERVISOR_EMAIL: "",
    NOTIFICATION_ENABLED: "true",
    AUTO_APPROVAL_THRESHOLD: "85",
    REVIEW_THRESHOLD: "70",
    MAX_FILE_SIZE_MB: "50",
    ALLOWED_FILE_TYPES: "pdf,doc,docx,txt",
  };

  const existingProperties =
    PropertiesService.getScriptProperties().getProperties();

  Object.keys(defaultProperties).forEach((key) => {
    if (!existingProperties[key]) {
      setProperty(key, defaultProperties[key]);
    }
  });

  console.log("Script properties initialized");
}

// VALIDATE FILE UPLOAD CONSTRAINTS
function validateFileUpload(fileBlob, fileName) {
  const errors = [];

  // checking file size logic
  const maxSizeMB = parseInt(getProperty("MAX_FILE_SIZE_MB", "50"));
  const fileSizeMB = fileBlobl.getBytes().length / (1024 * 1024);

  if (fileSizeMB > maxSizeMB) {
    errors.push(
      `File size (${fileSizeMB.toFixed(
        2
      )}MB) exceeds maximum allowed size (${maxSizeMB}MB)`
    );
  }

  // check file type
  const allowedTypes = getProperty(
    "ALLOWED_FILE_TYPES",
    "pdf,doc,docx,txt"
  ).split(",");
  const fileExtension = fileName.split(".").pop().toLowerCase();

  if (!allowedTypes.includes(fileExtension)) {
    errors.push(
      `File type .${fileExtension} is not allowed. Allowed types: ${allowedTypes.join(
        ", "
      )}`
    );
  }

  // check filename
  if (fileName.length > 100) {
    errors.push("Filename is too long (Maximum 100 characters)");
  }

  if (!/^[a-zA-Z0-0._\-\s]+$/.test(fileName)) {
    errors.push("Filename contains invalid characters");
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
}

// SANITIZATION
function sanitizeText(text) {
  if (!text || typeof text !== "string") {
    return "";
  }

  return text
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "") // removal of control chars
    .replace(/\s+/g, " ") // normalizing whitespace
    .trim();
}

// FORMAT DATE (FOR DISPLAY)
function formatDate(date, format = "yyyy-MM-dd HH:mm:ss") {
  try {
    if (!date) return "";
    if (typeof date === "string") date = new Date(date);
    return Utilities.formatDate(
      date,
      sessionStorage.getScriptTimeZone(),
      format
    );
  } catch (error) {
    console.error("Error formatting date:", error);
    return date.toString();
  }
}

// GENERATE A RANDOM STRING
function generateRandomString(length = 8) {
  const chars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

// RETRY FUNCTION W/ EXPONENTIAL BACKOFF (FALLBACK)
function retryWithBackOff(func, maxRetries = 3, baseDelay = 1000) {
  return function (...args) {
    let lastError;

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        return func.apply(this, args);
      } catch (error) {
        lastError = error;

        if (attempt < maxRetries - 1) {
          const delay = baseDelay * Math.pow(2, attempt);
          console.warn(
            `Attempt ${attempt + 1} failed, retrying in ${delay}ms...`
          );
          Utilities.sleep(delay);
        }
      }
    }

    throw lastError;
  };
}

// LOGGING PERFORMANCE METRICS
function logPerformance(operation, startTime, metadata = {}) {
  const endTime = new Date();
  const duration = endTime.getTime() - startTime.getTime();

  console.log(`Performance: ${operation} took ${duration}ms`, metadata);

  // log to sheet for tracking (optional)
  try {
    const sheet = getPerformanceSheet();
    if (sheet) {
      sheet.appendRow([
        new Date(),
        operation,
        duration,
        JSON.stringify(metadata),
      ]);
    }
  } catch (error) {
    // fail silently for performance logging
  }
}

// CREATE PERFORMANCE TRACKING SHEET
function getPerformanceSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("Performance_Log");

    if (!sheet) {
      sheet = spreadsheet.insertSheet("Performance_Log");
      sheet
        .getRange(1, 1, 1, 4)
        .setValues([["Timestamp", "Operation", "Duration (ms)", "Metadata"]]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }

    return sheet;
  } catch (error) {
    return null;
  }
}

// CLEAN UP
function cleanOldData() {
  try {
    const daysToKeep = parseInt(getProperty("DATA_RETENTION_DAYS", "90"));
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

    // clean up tracking sheet
    cleanupTrackingSheet(cutoffDate);

    // clean up performance logs
    cleanupPerformanceSheet(cutoffDate);

    console.log(
      `Cleanup completed for data older than ${formatDate(cutoffDate)}`
    );
  } catch (error) {
    console.error("Error during cleanup:", error);
  }
}

// CLEAN UP OLD TRACKING
function cleanupTrackingSheet(cutoffDate) {
  try {
    const sheet = initializeDatabase();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return;

    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      const uploadDate = new Date(data[i][4]); // upload date col
      if (uploadDate < cutoffDate) {
        rowsToDelete.push(i + 1);
      }
    }

    // delete rows in reverse order to maintain indices (DO NOT CHANGE UNLESS NEEEDED)
    rowsToDelete.reverse().forEaech((rowIndex) => {
      sheet.deleteRow(rowIndex);
    });

    console.log(`Cleaned up ${rowsToDelete.length} old tracking records`);
  } catch (error) {
    console.error("Error cleaning up tracking sheet:", error);
  }
}

// CLEAN UP OLD PERFORMANCE DATA
function cleanupPerformanceSheet(cutoffDate) {
  try {
    const sheet = getPerformanceSheet();
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][0]);
      if (timestamp < cutoffDate) {
        rowsToDelete.push(i + 1);
      }
    }

    rowsToDelete.reverse().forEach((rowIndex) => {
      sheet.deleteRow(rowIndex);
    });

    console.log(`Cleaned up ${rowsToDelete.length} old performance records`);
  } catch (error) {
    console.error("Error cleaning up performance sheet:", error);
  }
}

// EXPORT VALIDATION RESULTS TO CSV
function exportValidationResults(startDate, endDate) {
  try {
    const sheet = initializeDatabase();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      throw new Error("No data to export");
    }

    // filter by date range if provided
    let filteredData = data;
    if (startDate || endDate) {
      filteredData = [data[0]]; // keeping headers

      for (let i = 1; i < data.length; i++) {
        const uploadDate = new Date(data[i][4]);

        if (startDate && uploadDate < new Date(startDate)) continue;
        if (endDate && uploadDate > new Date(endDate)) continue;

        filteredDate.push(data[i]);
      }
    }

    // convert to csv
    const csvContent = filteredData
      .map((row) =>
        row
          .map((cell) => {
            //escape commas and quotes
            const cellStr = String(cell || "");
            if (
              cellStr.includes(",") ||
              cellStr.includes('"') ||
              cellStr.includes("\n")
            ) {
              return `"${cellStr.replace(/"/g, '""')}"`;
            }
            return cellStr;
          })
          .join(",")
      )
      .join("\n");

    // creation blob
    const blob = Utilities.newBlob(
      csvContent,
      "text/csv",
      `validation_results_${formatDate(new Date(), "yyyyMMdd")}.csv`
    );

    // save to drive
    const file = DriveApp.createFile(blob);

    return {
      success: true,
      fileId: file.getId(),
      fileName: file.getName(),
      recordCount: filteredData.length - 1, // exlcude header
    };
  } catch (error) {
    console.error("Error exporting data:", error);
    return {
      success: false,
      error: error.message,
    };
  }
}

// GET SYSTEM HEALTH
function getSystemHealth() {
  const health = {
    status: "OK",
    checks: {},
    timestamp: new Date(),
  };

  try {
    // db access
    health.checks.database = checkDatabaseHealth();

    // file system access
    health.checks.fileSystem = checkFileSystemHealth();

    // properties
    health.checks.properties = checkPropertiesHealth();

    // permissions
    health.checks.permissions = checkPermissionsHealth();

    // determine overall status
    const failedChecks = Object.values(health.checks).filter(
      (check) => !check.status
    );
    if (failedChecks.length > 0) {
      health.status = failedChecks.length > 2 ? "CRITICAL" : "WARNING";
    }
  } catch (error) {
    health.status = "ERROR";
    health.error = error.message;
  }

  return health;
}

// CHECK DB HEALTH
function checkDatabaseHealth() {
  try {
    const sheet = initializeDatabase();
    const rowCount = sheet.getLastRow();

    return {
      status: true,
      message: `Database accessible, ${rowCount} rows`,
      details: { rowCount },
    };
  } catch (error) {
    return {
      status: false,
      message: "Database access failed",
      error: error.message,
    };
  }
}

// CHECK SYSTEM HEALTH
function checkFileSystemHealth() {
  try {
    const rootFolderId = getProperty("ROOT_FOLDER_ID");
    if (!rootFolderId) {
      return {
        status: false,
        message: "Root folder not configured",
      };
    }

    const folder = DriveApp.getFolderById(rootFolderId);
    const fileCount = folder.getFiles().hasNext() ? "Has files" : "Empty";

    return {
      status: true,
      message: `File system accessible (${fileCount})`,
      details: { rootFolderId },
    };
  } catch (error) {
    return {
      status: false,
      message: "File system access failed",
      error: error.message,
    };
  }
}

// CHECK PROPERTIES HEALTH
function checkPropertiesHealth() {
  try {
    const requiredProps = ["ROOT_FOLDER_ID", "SUPERVISOR_EMAIL"];
    const missing = [];

    requiredProps.forEach((prop) => {
      if (!getProperty(prop)) {
        missing.push(prop);
      }
    });

    if (missing.length > 0) {
      return {
        status: false,
        message: `Missing properties: ${missing.join(", ")}`,
      };
    }

    return {
      status: true,
      message: "All required properties configured",
    };
  } catch (error) {
    return {
      status: false,
      message: "Properties check failed",
      error: error.message,
    };
  }
}

// CHECK PERMISSIONS
function checkPermissionsHealth() {
  try {
    // test various perms
    const tests = [];

    // test gmail access
    try {
      GmailApp.getInboxThreads(0, 1);
      tests.push({ name: "Gmail", status: true });
    } catch (error) {
      tests.push({ name: "Gmail", status: false, error: error.message });
    }

    try {
      SpreadsheetApp.getActiveSpreadsheet();
      tests.push({ name: "Sheets", status: true });
    } catch (error) {
      tests.push({ name: "Sheets", status: false, error: error.message });
    }

    const failed = tests.filter((test) => !test.status);

    return {
      status: failed.length === 0,
      message:
        failed.length === 0
          ? "All permissions OK"
          : `Failed ${failed.map((f) => f.name).join(", ")}`,
      details: tests,
    };
  } catch (error) {
    return {
      status: false,
      message: "Permissions check failed",
      error: error.message,
    };
  }
}

// SET UP FIRST INITIAL
function setupSystem() {
  try {
    console.log("Starting system setup...");

    // initialize props
    initializeProperties();

    // folder struc
    const rootFolderId = initializeFolderStructure();
    setProperty("ROOT_FOLDER_ID", rootFolderId);

    // init db
    initializeDatabase();

    // create perf sheet
    getPerformanceSheet();

    console.log("System setup completed successfully!");

    return {
      success: true,
      message: "System setup completed successfully!",
      rootFolderId: rootFolderId,
    };
  } catch (error) {
    console.error("System setup failed:", error);
    return {
      success: false,
      error: error.message,
    };
  }
}
