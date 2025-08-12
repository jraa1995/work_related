/**
 * Enhanced main.js for TAR validation system
 */

// When running in a Node.js environment (e.g. during unit testing) the global
// CONFIG object provided by Apps Script is not available. Import the
// configuration explicitly so that dependent functions have access to
// settings like API keys and validation rules. This code will be ignored
// in the Apps Script runtime where `module` is undefined.
if (typeof module !== 'undefined' && module.exports) {
  // eslint-disable-next-line global-require
  var CONFIG = require('./config.js');

  // Import utility functions when running under Node. In the Apps Script
  // environment these functions are available globally and require() is
  // undefined. We destructure only the functions used within this file.
  const utils = require('./utils.js');
  var average = utils.average;
  var fetchPerDiemByCityState = utils.fetchPerDiemByCityState;
  var validateFormData = utils.validateFormData;
  var calculateExpectedCosts = utils.calculateExpectedCosts;
  var generateValidationReport = utils.generateValidationReport;
  var logValidationResults = utils.logValidationResults;
  var extractTextFromPDF = utils.extractTextFromPDF;
  var extractGSAFormData = utils.extractGSAFormData;
}

// Define a fallback Logger for Node.js. In Apps Script the global Logger
// object will already exist, so this definition will be ignored.
if (typeof Logger === 'undefined') {
  var Logger = {
    log: function () {},
    warn: function () {},
    error: function () {},
  };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Enhanced TAR Validator")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Enhanced TAR validation with document extraction and comprehensive validation
 */
function validateTarWithPerDiem(data) {
  try {
    Logger.log("Starting TAR validation process...");

    // Initialize result object
    const result = {
      success: false,
      extractedData: {},
      validationReport: null,
      errors: [],
      warnings: [],
    };

    // Handle document extraction if PDF content provided
    if (data.pdfContent) {
      Logger.log("Extracting data from PDF...");
      const extractedText = extractTextFromPDF(data.pdfContent);

      if (extractedText) {
        result.extractedData = extractGSAFormData(extractedText);
        Logger.log("PDF extraction successful");
      } else {
        result.warnings.push("PDF extraction failed - using manual input");
      }
    }

    // Merge extracted data with manual input (manual input takes precedence)
    const mergedData = {
      ...result.extractedData,
      ...data,
      // Handle numeric conversions
      duration: parseInt(data.duration) || 1,
      totalCost: parseFloat(data.totalCost) || 0,
    };

    // Map manual form fields to expected names used by validation rules. When a
    // value is missing in extractedData but provided by the user we copy it
    // over to the canonical field. For example `traveler` becomes
    // `travelerName` and `purpose` becomes `travelPurpose`. We also build
    // a dutyStation from city/state when not present.
    if (!mergedData.travelerName && mergedData.traveler) {
      mergedData.travelerName = mergedData.traveler;
    }
    if (!mergedData.travelPurpose && mergedData.purpose) {
      mergedData.travelPurpose = mergedData.purpose;
    }
    if (!mergedData.dutyStation && mergedData.city && mergedData.state) {
      mergedData.dutyStation = `${mergedData.city}, ${mergedData.state}`;
    }
    if (!mergedData.contactNumber && mergedData.poc) {
      mergedData.contactNumber = mergedData.poc;
    }
    // Use totalCost as estimatedCost fallback
    if (!mergedData.estimatedCost && typeof mergedData.totalCost === 'number') {
      mergedData.estimatedCost = mergedData.totalCost;
    }

    // If dutyStation is still missing but an itinerary is provided, use the
    // first itinerary location as the duty station. This ensures trips using
    // itinerary data still satisfy the required dutyStation field. Use
    // fallback city/state values if available.
    if (
      !mergedData.dutyStation &&
      Array.isArray(mergedData.itinerary) &&
      mergedData.itinerary.length > 0
    ) {
      const firstStop = mergedData.itinerary[0];
      if (firstStop && firstStop.city && firstStop.state) {
        mergedData.dutyStation = `${firstStop.city}, ${firstStop.state}`;
      }
    }

    // If city/state are missing but dutyStation is provided (e.g. "City, ST"),
    // attempt to parse city and state from dutyStation. This supports
    // scenarios where the user enters a duty station manually but omits
    // separate city/state fields. We only parse if both city and state are
    // empty and dutyStation contains a comma.
    if (
      (!mergedData.city || !mergedData.state) &&
      typeof mergedData.dutyStation === 'string' &&
      mergedData.dutyStation.includes(',')
    ) {
      const parts = mergedData.dutyStation.split(',');
      if (parts.length >= 2) {
        const cityPart = parts[0].trim();
        const statePart = parts[1].trim().substring(0, 2).toUpperCase();
        if (!mergedData.city) mergedData.city = cityPart;
        if (!mergedData.state) mergedData.state = statePart;
      }
    }

    // Validate form data
    const validation = validateFormData(mergedData);
    if (!validation.isValid) {
      result.errors = validation.errors;
      return result;
    }

    // Calculate expected costs
    let expectedCosts;

    if (mergedData.itinerary && mergedData.itinerary.length > 0) {
      // Use extracted itinerary data
      expectedCosts = calculateExpectedCosts(mergedData.itinerary);
    } else {
      // Use manual city/state input
      const city = mergedData.city || "Unknown";
      const state = mergedData.state || "Unknown";
      const duration = mergedData.duration;

      const rateData = fetchPerDiemByCityState(city, state);
      let mie = CONFIG.DEFAULT_MIE;
      let avgLodging = CONFIG.DEFAULT_LODGING;

      if (rateData) {
        mie = parseFloat(rateData.Meals) || CONFIG.DEFAULT_MIE;
        avgLodging = average([
          rateData.Jan,
          rateData.Feb,
          rateData.Mar,
          rateData.Apr,
          rateData.May,
          rateData.Jun,
          rateData.Jul,
          rateData.Aug,
          rateData.Sep,
          rateData.Oct,
          rateData.Nov,
          rateData.Dec,
        ]);
      } else {
        result.warnings.push(
          "Unable to fetch GSA rates - using default values"
        );
      }

      const dailyTotal = mie + avgLodging;
      expectedCosts = {
        totalExpected: dailyTotal * duration,
        breakdown: [
          {
            location: `${city}, ${state}`,
            date: mergedData.tripDate || new Date().toISOString().split("T")[0],
            mie: mie,
            lodging: avgLodging,
            total: dailyTotal,
          },
        ],
      };
    }

    // Generate comprehensive validation report
    const validationReport = generateValidationReport(
      mergedData,
      expectedCosts,
      mergedData.totalCost
    );

    // Log results for audit trail
    logValidationResults(validationReport);

    // Determine final validation status
    const isValid =
      validationReport.validation.isWithinBuffer &&
      validationReport.validation.isWithinDeviation;

    result.success = true;
    result.validationReport = validationReport;
    result.isValid = isValid;
    result.expectedCost = expectedCosts.totalExpected.toFixed(2);
    result.claimedCost = mergedData.totalCost.toFixed(2);
    result.duration = mergedData.duration;
    result.variance = validationReport.validation.variance.toFixed(2);
    result.variancePercent = validationReport.validation.variancePercent;
    result.breakdown = expectedCosts.breakdown;

    result.message = isValid
      ? "✅ Trip cost is within acceptable per diem range."
      : `⚠️ Trip cost validation failed. Variance: ${result.variancePercent}%`;

    Logger.log("TAR validation completed successfully");
    return result;
  } catch (error) {
    Logger.log(`TAR validation error: ${error.message}`);
    return {
      success: false,
      errors: [`System error: ${error.message}`],
      isValid: false,
      message: "❌ Validation failed due to system error.",
    };
  }
}

/**
 * Test function for GSA API connectivity
 */
function testGSAAPI() {
  try {
    const testCity = "Washington";
    const testState = "DC";
    const result = fetchPerDiemByCityState(testCity, testState);

    if (result) {
      Logger.log(`GSA API test successful for ${testCity}, ${testState}`);
      Logger.log(`M&IE: $${result.Meals}, Lodging rates available`);
      return {
        success: true,
        data: result,
        message: "GSA API connectivity confirmed",
      };
    } else {
      Logger.log("GSA API test failed - no data returned");
      return {
        success: false,
        message: "GSA API test failed - no data returned",
      };
    }
  } catch (error) {
    Logger.log(`GSA API test error: ${error.message}`);
    return {
      success: false,
      message: `GSA API test error: ${error.message}`,
    };
  }
}

/**
 * Function to get GSA per diem rates for a specific location
 */
function getPerDiemRates(city, state) {
  try {
    const rateData = fetchPerDiemByCityState(city, state);

    if (rateData) {
      const avgLodging = average([
        rateData.Jan,
        rateData.Feb,
        rateData.Mar,
        rateData.Apr,
        rateData.May,
        rateData.Jun,
        rateData.Jul,
        rateData.Aug,
        rateData.Sep,
        rateData.Oct,
        rateData.Nov,
        rateData.Dec,
      ]);

      return {
        success: true,
        data: {
          city: city,
          state: state,
          meals: parseFloat(rateData.Meals) || CONFIG.DEFAULT_MIE,
          lodging: avgLodging,
          total:
            (parseFloat(rateData.Meals) || CONFIG.DEFAULT_MIE) + avgLodging,
          monthlyRates: {
            Jan: rateData.Jan,
            Feb: rateData.Feb,
            Mar: rateData.Mar,
            Apr: rateData.Apr,
            May: rateData.May,
            Jun: rateData.Jun,
            Jul: rateData.Jul,
            Aug: rateData.Aug,
            Sep: rateData.Sep,
            Oct: rateData.Oct,
            Nov: rateData.Nov,
            Dec: rateData.Dec,
          },
        },
      };
    } else {
      return {
        success: false,
        message: `No GSA rates found for ${city}, ${state}`,
        data: {
          city: city,
          state: state,
          meals: CONFIG.DEFAULT_MIE,
          lodging: CONFIG.DEFAULT_LODGING,
          total: CONFIG.DEFAULT_MIE + CONFIG.DEFAULT_LODGING,
          usingDefaults: true,
        },
      };
    }
  } catch (error) {
    Logger.log(`Error fetching per diem rates: ${error.message}`);
    return {
      success: false,
      message: `Error fetching rates: ${error.message}`,
    };
  }
}

/**
 * Function to export validation results
 */
function exportValidationResults(format = "json") {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty(
      "VALIDATION_LOG_SHEET_ID"
    );

    if (!spreadsheetId) {
      return {
        success: false,
        message: "No validation data available for export",
      };
    }

    const sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
    const data = sheet.getDataRange().getValues();

    if (format === "csv") {
      const csvContent = data.map((row) => row.join(",")).join("\n");
      const blob = Utilities.newBlob(
        csvContent,
        "text/csv",
        "tar_validation_log.csv"
      );

      return {
        success: true,
        data: blob,
        message: "CSV export ready",
      };
    } else {
      const jsonData = data.slice(1).map((row) => ({
        timestamp: row[0],
        traveler: row[1],
        authNumber: row[2],
        expectedCost: row[3],
        claimedCost: row[4],
        variance: row[5],
        variancePercent: row[6],
        status: row[7],
      }));

      return {
        success: true,
        data: jsonData,
        message: "JSON export ready",
      };
    }
  } catch (error) {
    Logger.log(`Export error: ${error.message}`);
    return {
      success: false,
      message: `Export failed: ${error.message}`,
    };
  }
}

/**
 * Function to clear validation logs
 */
function clearValidationLogs() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty(
      "VALIDATION_LOG_SHEET_ID"
    );

    if (spreadsheetId) {
      const sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
      sheet.clear();

      // Re-add headers
      sheet
        .getRange(1, 1, 1, 8)
        .setValues([
          [
            "Timestamp",
            "Traveler",
            "Auth Number",
            "Expected Cost",
            "Claimed Cost",
            "Variance",
            "Variance %",
            "Status",
          ],
        ]);

      return {
        success: true,
        message: "Validation logs cleared successfully",
      };
    } else {
      return {
        success: false,
        message: "No validation logs found to clear",
      };
    }
  } catch (error) {
    Logger.log(`Clear logs error: ${error.message}`);
    return {
      success: false,
      message: `Failed to clear logs: ${error.message}`,
    };
  }
}

// Export server-side functions for testing in Node.js environments. When
// running in Google Apps Script these exports will be ignored as `module`
// is undefined. Unit tests can import these to validate logic without the
// Apps Script runtime.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    doGet,
    include,
    validateTarWithPerDiem,
    testGSAAPI,
    getPerDiemRates,
    exportValidationResults,
    clearValidationLogs,
  };
}
