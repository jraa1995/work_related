// ===== DATABASE OPERATIONS

function initializeDatabase() {
  const sheetName = "Document_Tracking";

  try {
    let sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== sheetName) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const existingSheet = spreadsheet.getSheetByName(sheetName);

      if (existingSheet) {
        sheet = existingSheet;
      } else {
        sheet = spreadsheet.insertSheet(sheetName);
      }
    }

    if (sheet.getLastRow() === 0) {
      const headers = [
        "Tracking ID",
        "File Name",
        "File ID",
        "Document Type",
        "Upload Date",
        "Uploaded By",
        "Status",
        "Risk Score",
        "Completeness Score",
        "Compliance Score",
        "Format Score",
        "Issues Count",
        "Issues List",
        "Validation Details",
        "Reviewer",
        "Review Date",
        "Review Comments",
        "Final Status",
      ];

      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    return sheet;
  } catch (error) {
    console.error("Error initializing database:", error);
    throw new Error("Failed to initialize tracking database");
  }
}

// ===== SAVE VALIDATION RESULTS

function saveToTrackingSheet(data) {
  const sheet = initializeDatabase();
  const trackingId = generateTrackingId();

  const row = [
    trackingId,
    data.fileName,
    data.fileId,
    data.documentType,
    data.uploadDate,
    data.uploadedBy,
    data.status,
    data.riskScore,
    data.validationResults.completenessScore,
    data.validationResults.complianceScore,
    data.validationResults.formatScore,
    data.validationResults.issues.length,
    data.validationResults.issues.join("; "),
    JSON.stringify(data.validationResults.validationDetails),
    "", // reviewer (to be assigned)
    "", // review date (to be filled)
    "", // review comments (to be filled)
    data.status, // final status
  ];

  sheet.appendRow(row);

  console.log(`Saved tracking record: ${trackingId}`);
  return trackingId;
}

// ===== GENERATE UNIQUE ID TRACKING

function generateTrackingId() {
  const timestamp = Utilities.formatDate(
    new Date(),
    sessionStorage.getScriptTimeZone(),
    "yyyyMMdd"
  );
  const random = Math.floor(Math.random() * 10000).toString.padStart(4, "0");
  return `TRK${timestamp}${random}`;
}

// ===== GET VALIDATION RESULTS

function getDashboardData() {
  try {
    const sheet = initializeDatabase();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return {
        recentDocuments: [],
        statistics: {
          totalProcessed: 0,
          avgRiskScore: 0,
          statusBreakdown: {},
          typeBreakdown: {},
        },
      };
    }

    const data = sheet
      .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .getValues();

    const recentDocuments = data
      .slice(-10)
      .reverse()
      .map((row) => ({
        trackingId: row[0],
        fileName: row[1],
        documentType: row[3],
        uploadDate: row[4],
        status: row[6],
        riskScore: row[7],
      }));

    const statistics = calculateStatistics(data);

    return {
      recentDocuments: recentDocuments,
      statistics: statistics,
    };
  } catch (error) {
    console.error("Error getting dashboard data:", error);
    return {
      recentDocuments: [],
      statistics: {
        totalProcessed: 0,
        avgRiskScore: 0,
        statusBreakdown: {},
        typeBreakdown: {},
      },
    };
  }
}

// ===== CALCULATED DASHBOARD STATS

function calculateStatistics(data) {
  const stats = {
    totalProcessed: data.length,
    avgRiskScore: 0,
    statusBreakdown: {},
    typeBreakdown: {},
  };

  if (data.length === 0) return stats;

  let totalRiskScore = 0;

  data.forEach((row) => {
    const status = row[6] || "UNKNOWN";
    const docType = row[3] || "UNKNOWN";
    const riskScore = row[7] || 0;

    // status breakdown
    stats.statusBreakdown[status] = (stats.statusBreakdown[status] || 0) + 1;

    // type breakdown
    stats.typeBreakdown[docType] = (stats.typeBreakdown[docType] || 0) + 1;

    // risk score sum
    totalRiskScore += riskScore;
  });

  stats.avgRiskScore = Math.round(totalRiskScore / data.length);

  return stats;
}
