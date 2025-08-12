/**
 * Code.gs - Web App Entry Point
 * This file serves the HTML web app and handles web app routing
 */

/**
 * Serves the HTML web app when accessed via URL
 * @param {Object} e - The event parameter containing request information
 * @return {HtmlOutput} The HTML page to display
 */
function doGet(e) {
  try {
    // Create HTML output from the dashboard template
    const htmlOutput = HtmlService.createTemplateFromFile("dashboard");

    // Set page title and other metadata
    const html = htmlOutput
      .evaluate()
      .setTitle("CLP Dashboard")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");

    return html;
  } catch (error) {
    Logger.log("Error in doGet: " + error.toString());

    // Return error page if something goes wrong
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Error - CLP Dashboard</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              padding: 50px;
              background-color: #f8f9fa;
            }
            .error-container {
              background: white;
              padding: 30px;
              border-radius: 10px;
              box-shadow: 0 2px 10px rgba(0,0,0,0.1);
              max-width: 500px;
              margin: 0 auto;
            }
            .error-title {
              color: #dc3545;
              font-size: 24px;
              margin-bottom: 15px;
            }
            .error-message {
              color: #6c757d;
              font-size: 16px;
              line-height: 1.5;
            }
          </style>
        </head>
        <body>
          <div class="error-container">
            <h1 class="error-title">⚠️ Dashboard Error</h1>
            <p class="error-message">
              There was an error loading the CLP Dashboard. Please check the console logs and ensure all required files are properly set up.
            </p>
            <p class="error-message">
              <strong>Error:</strong> ${error.toString()}
            </p>
          </div>
        </body>
      </html>
    `);
  }
}

/**
 * Includes HTML files (for modular HTML development)
 * @param {string} filename - The name of the file to include
 * @return {string} The content of the file
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log("Error including file " + filename + ": " + error.toString());
    return "<!-- Error loading " + filename + " -->";
  }
}

/**
 * Test function to check if the web app is working
 * @return {string} Test message
 */
function testWebApp() {
  try {
    const data = getDashboardData();
    return {
      success: true,
      message: "Web app is working correctly!",
      dataCount: data.length,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    return {
      success: false,
      message: "Error testing web app: " + error.toString(),
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * Get web app URL for easy access
 * @return {string} The web app URL
 */
function getWebAppUrl() {
  try {
    // Note: This will only work after the web app is deployed
    const scriptId = ScriptApp.getScriptId();
    return `https://script.google.com/macros/s/${scriptId}/exec`;
  } catch (error) {
    Logger.log("Error getting web app URL: " + error.toString());
    return "Deploy the web app first to get the URL";
  }
}

/**
 * Initialize the web app (run this once after setup)
 */
function initializeWebApp() {
  try {
    Logger.log("Initializing CLP Dashboard Web App...");

    // Check if Dashboard sheet exists
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let dashboardSheet = spreadsheet.getSheetByName("Dashboard");

    if (!dashboardSheet) {
      Logger.log("Dashboard sheet not found. Creating it...");
      dashboardSheet = spreadsheet.insertSheet("Dashboard");
    }

    // Check if there's data in the Dashboard sheet
    const lastRow = dashboardSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log(
        "No data in Dashboard sheet. Run processLatestCLPDataRobust() first."
      );
      return {
        success: false,
        message:
          "No data found. Please run processLatestCLPDataRobust() to populate the Dashboard sheet first.",
      };
    }

    Logger.log("Web app initialized successfully!");
    Logger.log("Dashboard sheet has " + (lastRow - 1) + " rows of data.");

    return {
      success: true,
      message: "Web app initialized successfully!",
      dataRows: lastRow - 1,
    };
  } catch (error) {
    Logger.log("Error initializing web app: " + error.toString());
    return {
      success: false,
      message: "Error initializing web app: " + error.toString(),
    };
  }
}

/**
 * Handle POST requests (if needed for future functionality)
 * @param {Object} e - The event parameter containing request information
 * @return {ContentService.TextOutput} JSON response
 */
function doPost(e) {
  try {
    // Parse the POST data
    const data = JSON.parse(e.postData.contents);

    // Handle different POST actions
    switch (data.action) {
      case "refreshData":
        // Refresh the dashboard data
        processLatestCLPDataRobust();
        return ContentService.createTextOutput(
          JSON.stringify({
            success: true,
            message: "Data refreshed successfully",
          })
        ).setMimeType(ContentService.MimeType.JSON);

      default:
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Unknown action",
          })
        ).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
