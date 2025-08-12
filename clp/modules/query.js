/**
 * main function to process excel file
 */
function processLatestCLPData() {
  try {
    // find the CLP_Data.xlsx file
    const excelFile = findCLPDataFile();
    if (!excelFile) {
      Logger.log("CLP_Data.xlsx file not found");
      return;
    }

    // convert excel to g-sheets
    const convertedSheetId = convertExcelToGoogleSheets(excelFile);
    const sourceSpreadsheet = SpreadsheetApp.openById(convertedSheetId);
    const sourceSheet = sourceSpreadsheet.getSheets()[0];

    // dashboard sheet
    const dashboardSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet =
      dashboardSpreadSheet.getSheetByName("Dashboard") ||
      dashboardSpreadSheet.insertSheet("Dashboard");

    // extract and transform
    const sourceData = extractAndTransformData(sourceSheet);

    // populate dashboard with transformed data
    populateDashboard(dashboardSheet, sourceData);

    Logger.log("Dashboard updated successfully");

    // Clean up the temporary converted file
    DriveApp.getFileById(convertedSheetId).setTrashed(true);
  } catch (error) {
    Logger.log("Error processing CLP Data: " + error.toString());
    throw error;
  }
}

/**
 * Debug function to test folder access
 */
function testFolderAccess() {
  try {
    const folderId = "YOUR_FOLDER_ID_HERE"; // Replace with actual folder ID
    Logger.log("Testing folder ID: " + folderId);

    const folder = DriveApp.getFolderById(folderId);
    Logger.log("Folder found: " + folder.getName());

    const files = folder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      Logger.log("File found: " + file.getName());
      fileCount++;
      if (fileCount > 10) break; // Limit output
    }
  } catch (error) {
    Logger.log("Folder access error: " + error.toString());
  }
}

/**
 * Alternative: Search for CLP_Data.xlsx across all accessible files
 */
function findCLPDataFileGlobally() {
  try {
    Logger.log("Searching for CLP_Data.xlsx globally...");
    const files = DriveApp.getFilesByName("CLP_Data.xlsx");

    if (files.hasNext()) {
      const file = files.next();
      Logger.log("Found CLP_Data.xlsx: " + file.getId());
      return file;
    }

    Logger.log("CLP_Data.xlsx not found in accessible files");
    return null;
  } catch (error) {
    Logger.log("Global search error: " + error.toString());
    return null;
  }
}

/**
 * find the CLP_Data.xlsx file specifically
 */
function findCLPDataFile() {
  // Method 1: Try folder-based search
  try {
    const folderId = "1VETKoY5ZkngPULVxG_G8a9vxFEMbarPl"; // Replace with actual folder ID

    if (folderId === "1VETKoY5ZkngPULVxG_G8a9vxFEMbarPl") {
      Logger.log("Please set the actual folder ID in the code");
      return findCLPDataFileGlobally(); // Fall back to global search
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName("CLP_Data.xlsx");

    if (files.hasNext()) {
      return files.next();
    }

    Logger.log(
      "CLP_Data.xlsx not found in specified folder, trying global search..."
    );
    return findCLPDataFileGlobally();
  } catch (error) {
    Logger.log(
      "Folder search failed: " + error.toString() + ". Trying global search..."
    );
    return findCLPDataFileGlobally();
  }
}

/**
 * Alternative method: Search in the same folder as this spreadsheet
 */
function findCLPDataInCurrentFolder() {
  try {
    // Get the current spreadsheet's folder
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const currentFile = DriveApp.getFileById(currentSpreadsheet.getId());
    const parentFolders = currentFile.getParents();

    if (parentFolders.hasNext()) {
      const folder = parentFolders.next();
      Logger.log("Searching in current folder: " + folder.getName());

      const files = folder.getFilesByName("CLP_Data.xlsx");
      if (files.hasNext()) {
        return files.next();
      }
    }

    return null;
  } catch (error) {
    Logger.log("Current folder search error: " + error.toString());
    return null;
  }
}

/**
 * Master search function that tries multiple methods
 */
function findCLPDataFileRobust() {
  // Try method 1: Specific folder
  let file = findCLPDataFile();
  if (file) return file;

  // Try method 2: Current spreadsheet's folder
  file = findCLPDataInCurrentFolder();
  if (file) return file;

  // Try method 3: Global search
  file = findCLPDataFileGlobally();
  if (file) return file;

  Logger.log("CLP_Data.xlsx not found using any method");
  return null;
}

/**
 * conversion of excel file to gsheets using Drive API
 * @param {File} file the Excel File object
 * @return {string} the id of the converted gsheet
 */
function convertExcelToGoogleSheets(file) {
  const blob = file.getBlob();
  const fileName = file.getName().replace(/\.xlsx?$/, "") + " (Converted)";

  // Create file metadata
  const fileMetadata = {
    name: fileName,
    mimeType: "application/vnd.google-apps.spreadsheet",
  };

  // Create the file with conversion
  const convertedFile = Drive.Files.create(fileMetadata, blob, {
    uploadType: "multipart",
  });

  Logger.log("Excel file converted successfully: " + convertedFile.id);
  return convertedFile.id;
}

/**
 * extract and transform data from source
 * @param {Sheet} sourceSheet the source w/ raw data
 * @return {Array<Array>} transformed data
 */
function extractAndTransformData(sourceSheet) {
  // getting all data from the sheet
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  if (lastRow === 0) {
    throw new Error("No data found in the source sheet");
  }

  const sourceData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // get header and col indices
  const headers = sourceData[0];
  Logger.log("Available headers: " + headers.join(", "));

  const columnIndices = {
    employeeEmail: headers.indexOf("Employee Email"), // Fixed: removed 's'
    departmentId: headers.indexOf("Department ID"),
    supervisorEmail: headers.indexOf("Supervisor E-Mail Address"),
    clpStatusRange: headers.indexOf("CLP Status Range"),
    clpNumberPercent: headers.indexOf("CLP Number %"),
    clpsRequired: headers.indexOf("CLPs Required Employee Total"),
    metricClps: headers.indexOf("Metric CLPS"),
    clpsRemaining: headers.indexOf("CLPS Remaining"),
    facCProfessional: headers.indexOf("FAC-C (Professional)"),
    facPpmEmployees: headers.indexOf("FAC P/PM Employees"),
    facCor3: headers.indexOf("z - raw - FAC-COR Level 3"),
    facCor2: headers.indexOf("FAC-COR 2"),
    facCor1: headers.indexOf("FAC-COR 1"),
  };

  // validate cols and provide helpful error messages
  const missingColumns = [];
  for (const [key, index] of Object.entries(columnIndices)) {
    if (index === -1) {
      missingColumns.push(key);
    }
  }

  if (missingColumns.length > 0) {
    throw new Error(
      `Required columns not found: ${missingColumns.join(
        ", "
      )}. Available headers: ${headers.join(", ")}`
    );
  }

  // transform the data
  const transformedData = sourceData
    .slice(1)
    .map((row) => [
      row[columnIndices.employeeEmail],
      row[columnIndices.departmentId],
      row[columnIndices.supervisorEmail],
      row[columnIndices.clpStatusRange],
      row[columnIndices.clpNumberPercent],
      row[columnIndices.clpsRequired],
      row[columnIndices.metricClps],
      row[columnIndices.clpsRemaining],
      row[columnIndices.facCProfessional],
      row[columnIndices.facPpmEmployees],
      row[columnIndices.facCor3],
      row[columnIndices.facCor2],
      row[columnIndices.facCor1],
    ]);

  Logger.log(`Processed ${transformedData.length} rows of data`);
  return transformedData;
}

/**
 * populates the dash
 * @param {Sheet} dashboardSheet the target dash sheet
 * @param {Array<Array>} data the transformed data
 */
function populateDashboard(dashboardSheet, data) {
  dashboardSheet.clearContents();

  // Dashboard Headers
  const headers = [
    "employee",
    "org code",
    "supervisor",
    "range",
    "% earned",
    "required",
    "earned",
    "remaining",
    "FCP",
    "PPM",
    "COR 3",
    "COR 2",
    "COR 1",
  ];

  dashboardSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // add data if there's any
  if (data.length > 0) {
    dashboardSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  // apply formatting
  formatDashboard(dashboardSheet);

  Logger.log(`Dashboard populated with ${data.length} rows`);
}

/**
 * applies formatting
 * @param {Sheet} sheet the dashboard sheet
 */
function formatDashboard(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange
    .setBackground("#4285F4")
    .setFontColor("white")
    .setFontWeight("bold");

  // freeze header row
  sheet.setFrozenRows(1);

  // set alternating row colors (if there's data)
  if (sheet.getLastRow() > 1) {
    const dataRange = sheet.getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      sheet.getLastColumn()
    );
    dataRange.applyRowBanding();
  }

  // auto-resize cols
  for (let i = 1; i <= sheet.getLastColumn(); i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Alternative main function using robust file finding
 */
function processLatestCLPDataRobust() {
  try {
    // find the CLP_Data.xlsx file using multiple methods
    const excelFile = findCLPDataFileRobust();
    if (!excelFile) {
      Logger.log("CLP_Data.xlsx file not found using any method");
      return;
    }

    Logger.log(
      "Found file: " + excelFile.getName() + " (ID: " + excelFile.getId() + ")"
    );

    // convert excel to g-sheets
    const convertedSheetId = convertExcelToGoogleSheets(excelFile);
    const sourceSpreadsheet = SpreadsheetApp.openById(convertedSheetId);
    const sourceSheet = sourceSpreadsheet.getSheets()[0];

    // dashboard sheet
    const dashboardSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet =
      dashboardSpreadSheet.getSheetByName("Dashboard") ||
      dashboardSpreadSheet.insertSheet("Dashboard");

    // extract and transform
    const sourceData = extractAndTransformData(sourceSheet);

    // populate dashboard with transformed data
    populateDashboard(dashboardSheet, sourceData);

    Logger.log("Dashboard updated successfully");

    // Clean up the temporary converted file
    DriveApp.getFileById(convertedSheetId).setTrashed(true);
  } catch (error) {
    Logger.log("Error processing CLP Data: " + error.toString());
    throw error;
  }
}
