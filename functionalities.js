/**
 * functionalities.gs - Dashboard Data Functions
 * This file handles reading processed data for the web app dashboard
 */

/**
 * Get dashboard data for the web app (reads from existing Dashboard sheet)
 * @return {Array<Object>} formatted data for the web app
 */
function getDashboardData() {
  try {
    const dashboardSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = dashboardSpreadSheet.getSheetByName("Dashboard");

    if (!dashboardSheet) {
      Logger.log(
        "Dashboard sheet not found. Please run processLatestCLPDataRobust() first."
      );
      return [];
    }

    const lastRow = dashboardSheet.getLastRow();
    const lastCol = dashboardSheet.getLastColumn();

    if (lastRow <= 1) {
      Logger.log("No data found in Dashboard sheet");
      return [];
    }

    // Get headers to understand the data structure
    const headers = dashboardSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log("Dashboard headers: " + headers.join(", "));

    // Find the correct column indices based on header names
    const columnIndices = {};
    headers.forEach((header, index) => {
      const cleanHeader = header.toString().toLowerCase().trim();
      columnIndices[cleanHeader] = index;
    });

    Logger.log("Column mapping found:");
    Object.keys(columnIndices).forEach((key) => {
      Logger.log(
        `  "${key}" -> Column ${columnIndices[key]} (${String.fromCharCode(
          65 + columnIndices[key]
        )})`
      );
    });

    // Get all data from Dashboard sheet (skip header row)
    const dashboardData = dashboardSheet
      .getRange(2, 1, lastRow - 1, lastCol)
      .getValues();

    // Format data for the web app using CORRECT column mapping
    const formattedData = dashboardData.map((row, index) => {
      const employeeData = {
        employee: row[columnIndices["employee"] || 0] || "",
        orgCode: row[columnIndices["org code"] || 1] || "",
        supervisor: row[columnIndices["supervisor"] || 2] || "",
        range: row[columnIndices["range"] || 3] || "", // THIS is the actual range column
        percentEarned: row[columnIndices["% earned"] || 4] || 0, // THIS is % earned
        required: row[columnIndices["required"] || 5] || 0,
        earned: row[columnIndices["earned"] || 6] || 0,
        remaining: row[columnIndices["remaining"] || 7] || 0,
        fcp: row[columnIndices["fcp"] || 8] || "",
        ppm: row[columnIndices["ppm"] || 9] || "",
        cor3: row[columnIndices["cor 3"] || 10] || "",
        cor2: row[columnIndices["cor 2"] || 11] || "",
        cor1: row[columnIndices["cor 1"] || 12] || "",
      };

      // Log first few rows for debugging with correct mapping
      if (index < 5) {
        Logger.log("Row " + (index + 1) + ":");
        Logger.log("  Employee: " + employeeData.employee);
        Logger.log("  Org Code: " + employeeData.orgCode);
        Logger.log(
          "  Range (from column " +
            (columnIndices["range"] || 3) +
            '): "' +
            employeeData.range +
            '"'
        );
        Logger.log(
          "  % Earned (from column " +
            (columnIndices["% earned"] || 4) +
            "): " +
            employeeData.percentEarned
        );
        Logger.log("  Range type: " + typeof employeeData.range);
        Logger.log("  % Earned type: " + typeof employeeData.percentEarned);
      }

      return employeeData;
    });

    Logger.log(
      "Successfully processed " + formattedData.length + " employee records"
    );

    // Additional verification - count range values
    const rangeCounts = {};
    formattedData.forEach((emp) => {
      const range = emp.range;
      rangeCounts[range] = (rangeCounts[range] || 0) + 1;
    });
    Logger.log("Range value distribution in processed data:");
    Object.keys(rangeCounts).forEach((range) => {
      Logger.log('  "' + range + '": ' + rangeCounts[range] + " employees");
    });

    return formattedData;
  } catch (error) {
    Logger.log("Error getting dashboard data: " + error.toString());
    return [];
  }
}

/**
 * Get business unit statistics for the dashboard
 * @return {Object} business unit stats with completion data
 */
function getBusinessUnitStats() {
  try {
    const data = getDashboardData();
    // Define the list of business units to report on. QF1 has been merged into QF2 (SOS),
    // QFD has been deprecated and is therefore omitted. Any org codes prefixed with QFEE,
    // QFEEB or QFEEC will be remapped to QFA (Army).
    const businessUnits = ["QF2", "QFA", "QFB", "QFC", "QFE"];
    const businessUnitStats = {};

    // Initialize business units
    businessUnits.forEach((unit) => {
      businessUnitStats[unit] = {
        total: 0,
        completed: 0,
        incomplete: 0,
        employees: [],
      };
    });

    // Process each employee record
    data.forEach((record) => {
      const businessUnit = extractBusinessUnit(record.orgCode);

      if (businessUnitStats[businessUnit]) {
        businessUnitStats[businessUnit].total++;
        businessUnitStats[businessUnit].employees.push(record);

        if (isComplete(record.remaining)) {
          businessUnitStats[businessUnit].completed++;
        } else {
          businessUnitStats[businessUnit].incomplete++;
        }
      }
    });

    return businessUnitStats;
  } catch (error) {
    Logger.log("Error getting business unit stats: " + error.toString());
    return {};
  }
}

/**
 * Get overall dashboard statistics
 * @return {Object} overall stats for the dashboard
 */
function getOverallStats() {
  try {
    const data = getDashboardData();
    const totalEmployees = data.length;
    const completedCount = data.filter((record) =>
      isComplete(record.remaining)
    ).length;
    const incompleteCount = totalEmployees - completedCount;
    const completionRate =
      totalEmployees > 0
        ? Math.round((completedCount / totalEmployees) * 100)
        : 0;

    return {
      totalEmployees,
      completedCount,
      incompleteCount,
      completionRate,
    };
  } catch (error) {
    Logger.log("Error getting overall stats: " + error.toString());
    return {
      totalEmployees: 0,
      completedCount: 0,
      incompleteCount: 0,
      completionRate: 0,
    };
  }
}

/**
 * Get employees by business unit
 * @param {string} businessUnit - The business unit code (QF2, QFA, etc.)
 * @return {Array<Object>} employees in the specified business unit
 */
function getEmployeesByBusinessUnit(businessUnit) {
  try {
    const data = getDashboardData();
    return data.filter(
      (record) => extractBusinessUnit(record.orgCode) === businessUnit
    );
  } catch (error) {
    Logger.log(
      `Error getting employees for business unit ${businessUnit}: ` +
        error.toString()
    );
    return [];
  }
}

/**
 * Get incomplete employees (remaining CLPs >= 1)
 * @return {Array<Object>} employees who haven't completed their CLPs
 */
function getIncompleteEmployees() {
  try {
    const data = getDashboardData();
    return data.filter((record) => !isComplete(record.remaining));
  } catch (error) {
    Logger.log("Error getting incomplete employees: " + error.toString());
    return [];
  }
}

/**
 * Get completed employees (remaining CLPs < 1)
 * @return {Array<Object>} employees who have completed their CLPs
 */
function getCompletedEmployees() {
  try {
    const data = getDashboardData();
    return data.filter((record) => isComplete(record.remaining));
  } catch (error) {
    Logger.log("Error getting completed employees: " + error.toString());
    return [];
  }
}

/**
 * Get comprehensive dashboard data (combines all stats)
 * @return {Object} comprehensive dashboard data for the web app
 */
function getComprehensiveDashboardData() {
  try {
    const rawData = getDashboardData();
    const businessUnitStats = getBusinessUnitStats();
    const overallStats = getOverallStats();

    return {
      rawData,
      businessUnitStats,
      overallStats,
      lastUpdated: new Date().toISOString(),
    };
  } catch (error) {
    Logger.log(
      "Error getting comprehensive dashboard data: " + error.toString()
    );
    return {
      rawData: [],
      businessUnitStats: {},
      overallStats: {
        totalEmployees: 0,
        completedCount: 0,
        incompleteCount: 0,
        completionRate: 0,
      },
      lastUpdated: new Date().toISOString(),
    };
  }
}

// Helper functions

/**
 * Extract business unit from org code (first 3 characters)
 * @param {string} orgCode - The organization code
 * @return {string} business unit code
 */
function extractBusinessUnit(orgCode) {
  // Normalize input and handle missing values
  if (!orgCode || typeof orgCode !== "string") return "Unknown";
  const code = orgCode.toUpperCase().trim();

  // Certain Defence (QFEE...) codes belong in the Army section (QFA)
  if (code.startsWith("QFEE") || code.startsWith("QFEEB") || code.startsWith("QFEEC")) {
    return "QFA";
  }

  // Merge QF1 into QF2 (SOS)
  if (code.startsWith("QF1")) {
    return "QF2";
  }

  // Determine the standard business unit prefix
  const prefix = code.substring(0, 3);

  // Skip deprecated QFD entries by returning Unknown so they do not show up in stats
  if (prefix === "QFD") {
    return "Unknown";
  }

  return prefix;
}

/**
 * Check if an employee has completed their CLPs
 * @param {number} remaining - Number of remaining CLPs
 * @return {boolean} true if complete (remaining < 1)
 */
function isComplete(remaining) {
  return remaining < 1;
}

/**
 * Get the last update timestamp of the Dashboard sheet
 * @return {string} ISO timestamp of last update
 */
function getDashboardLastUpdated() {
  try {
    const dashboardSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = dashboardSpreadSheet.getSheetByName("Dashboard");

    if (!dashboardSheet) {
      return null;
    }

    // Get the last modified time of the spreadsheet
    // Note: This returns the last time the spreadsheet was modified, not specifically the Dashboard sheet
    const file = DriveApp.getFileById(dashboardSpreadSheet.getId());
    return file.getLastUpdated().toISOString();
  } catch (error) {
    Logger.log("Error getting last updated time: " + error.toString());
    return new Date().toISOString();
  }
}
