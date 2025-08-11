/**
 * main.gs
 *
 * This file contains the server‑side logic for the Enhanced TAR
 * Validation System when deployed as a Google Apps Script web app.
 *
 * It exposes functions for validating TAR requests, fetching per diem
 * rates, exporting validation logs, and clearing logs.  The code has
 * been refactored from a Node.js module and stripped of any
 * references to CommonJS (require/module.exports) so that it runs
 * natively within the Apps Script runtime.  All configuration is
 * pulled from the global CONFIG object defined in config.gs, and
 * utilities are available globally via utils.gs.
 */

/**
 * Serve the main application page.  When a user navigates to the web
 * app URL this function returns the HTML template stored in
 * Index.html.  The page title is set here.  Apps Script will
 * automatically sandbox the HTML in an IFRAME.
 *
 * @return {HtmlOutput} The rendered HTML page
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enhanced TAR Validation System');
}

/**
 * Include an arbitrary HTML file into another template.  This helper
 * allows you to assemble pages from multiple partials.  Usage:
 * <?= include('somePartial'); ?> inside an HTML file.
 *
 * @param {string} filename The name of the HTML file to include
 * @return {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Validate a TAR request using per diem rates and expense rules.
 *
 * @param {Object} data The TAR data to validate.  This object is
 *        expected to contain fields such as city, state, duration,
 *        totalCost, estimatedCost, traveler, etc.  See client.js for
 *        how form data and CSV records are mapped into this format.
 * @return {Object} A result object containing success flag,
 *         validation report and messages.  On error the success flag
 *         will be false and an errors array will describe the problem.
 */
function validateTarWithPerDiem(data) {
  try {
    // Initialize result object
    var result = {
      success: false,
      extractedData: {},
      validationReport: null,
      errors: [],
      warnings: []
    };

    // If a base64 PDF is supplied, attempt to extract data from it.
    if (data.pdfContent) {
      var extractedText = extractTextFromPDF(data.pdfContent);
      if (extractedText) {
        result.extractedData = extractGSAFormData(extractedText);
      } else {
        result.warnings.push('PDF extraction failed - using manual input');
      }
    }

    // Merge extracted values with the provided data.  Manual input
    // takes precedence over extracted values.  Convert strings to
    // appropriate types where necessary.
    var mergedData = Object.assign({}, result.extractedData, data);
    mergedData.duration = parseInt(mergedData.duration, 10) || 1;
    mergedData.totalCost = parseFloat(mergedData.totalCost) || 0;

    // Copy alternative field names into canonical names expected by
    // validation rules.  This ensures that data coming from the form
    // or from a CSV record is normalised.
    if (!mergedData.travelerName && mergedData.traveler) {
      mergedData.travelerName = mergedData.traveler;
    }
    if (!mergedData.travelPurpose && mergedData.purpose) {
      mergedData.travelPurpose = mergedData.purpose;
    }
    if (!mergedData.dutyStation && mergedData.city && mergedData.state) {
      mergedData.dutyStation = mergedData.city + ', ' + mergedData.state;
    }
    if (!mergedData.contactNumber && mergedData.poc) {
      mergedData.contactNumber = mergedData.poc;
    }
    // Use totalCost as a fallback for estimatedCost if not provided.
    if (!mergedData.estimatedCost && typeof mergedData.totalCost === 'number') {
      mergedData.estimatedCost = mergedData.totalCost;
    }
    // If city/state missing but dutyStation provided, parse out city/state
    if ((!mergedData.city || !mergedData.state) && mergedData.dutyStation) {
      var parts = mergedData.dutyStation.split(',');
      if (parts.length >= 2) {
        var cityPart = parts[0].trim();
        var statePart = parts[1].trim().substring(0, 2).toUpperCase();
        if (!mergedData.city) mergedData.city = cityPart;
        if (!mergedData.state) mergedData.state = statePart;
      }
    }

    // Validate the merged data structure.  If validation fails we
    // immediately return with the error messages.
    var validation = validateFormData(mergedData);
    if (!validation.isValid) {
      result.errors = validation.errors;
      return result;
    }

    // Compute expected per diem costs.  If an itinerary is provided
    // then calculate based on the itinerary; otherwise fetch GSA rates
    // for the city/state and multiply by duration.
    var expectedCosts;
    if (Array.isArray(mergedData.itinerary) && mergedData.itinerary.length > 0) {
      expectedCosts = calculateExpectedCosts(mergedData.itinerary);
    } else {
      var city = mergedData.city || 'Unknown';
      var state = mergedData.state || 'Unknown';
      var duration = mergedData.duration;
      var rateData = fetchPerDiemByCityState(city, state);
      var mie = CONFIG.DEFAULT_MIE;
      var avgLodging = CONFIG.DEFAULT_LODGING;
      if (rateData) {
        mie = parseFloat(rateData.Meals) || CONFIG.DEFAULT_MIE;
        avgLodging = average([
          rateData.Jan, rateData.Feb, rateData.Mar, rateData.Apr, rateData.May,
          rateData.Jun, rateData.Jul, rateData.Aug, rateData.Sep, rateData.Oct,
          rateData.Nov, rateData.Dec
        ]);
      } else {
        result.warnings.push('Unable to fetch GSA rates - using default values');
      }
      var dailyTotal = mie + avgLodging;
      expectedCosts = {
        totalExpected: dailyTotal * duration,
        breakdown: [
          {
            location: city + ', ' + state,
            date: mergedData.tripDate || new Date().toISOString().split('T')[0],
            mie: mie,
            lodging: avgLodging,
            total: dailyTotal
          }
        ]
      };
    }

    // Generate a comprehensive validation report.  This checks
    // variance between claimed and expected costs, applies buffer and
    // deviation thresholds, and returns arrays of messages,
    // warnings and errors.
    var validationReport = generateValidationReport(
      mergedData,
      expectedCosts,
      mergedData.totalCost
    );
    logValidationResults(validationReport);
    var isValid =
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
      ? '✅ Trip cost is within acceptable per diem range.'
      : '⚠️ Trip cost validation failed. Variance: ' + result.variancePercent + '%';

    return result;
  } catch (err) {
    return {
      success: false,
      errors: ['System error: ' + err.message],
      isValid: false,
      message: '❌ Validation failed due to system error.'
    };
  }
}

/**
 * Test GSA API connectivity.  Returns per diem rate information for
 * Washington, DC if the API call succeeds.
 *
 * @return {Object} An object with success flag and rate data or error
 */
function testGSAAPI() {
  try {
    var result = fetchPerDiemByCityState('Washington', 'DC');
    if (result) {
      return {
        success: true,
        data: result,
        message: 'GSA API connectivity confirmed'
      };
    } else {
      return {
        success: false,
        message: 'GSA API test failed - no data returned'
      };
    }
  } catch (err) {
    return {
      success: false,
      message: 'GSA API test error: ' + err.message
    };
  }
}

/**
 * Fetch GSA per diem rates for a city/state and compute average
 * monthly lodging.  This function is exposed to the client for
 * on‑demand lookups.
 *
 * @param {string} city The destination city
 * @param {string} state The destination state (2‑letter code)
 * @return {Object} An object containing rate information or error
 */
function getPerDiemRates(city, state) {
  try {
    var rate = fetchPerDiemByCityState(city, state);
    if (rate) {
      var avgLodging = average([
        rate.Jan, rate.Feb, rate.Mar, rate.Apr, rate.May, rate.Jun,
        rate.Jul, rate.Aug, rate.Sep, rate.Oct, rate.Nov, rate.Dec
      ]);
      return {
        success: true,
        data: {
          city: city,
          state: state,
          meals: parseFloat(rate.Meals) || CONFIG.DEFAULT_MIE,
          lodging: avgLodging,
          total: (parseFloat(rate.Meals) || CONFIG.DEFAULT_MIE) + avgLodging,
          monthlyRates: {
            Jan: rate.Jan, Feb: rate.Feb, Mar: rate.Mar, Apr: rate.Apr,
            May: rate.May, Jun: rate.Jun, Jul: rate.Jul, Aug: rate.Aug,
            Sep: rate.Sep, Oct: rate.Oct, Nov: rate.Nov, Dec: rate.Dec
          }
        }
      };
    } else {
      return {
        success: false,
        message: 'No GSA rates found for ' + city + ', ' + state,
        data: {
          city: city,
          state: state,
          meals: CONFIG.DEFAULT_MIE,
          lodging: CONFIG.DEFAULT_LODGING,
          total: CONFIG.DEFAULT_MIE + CONFIG.DEFAULT_LODGING,
          usingDefaults: true
        }
      };
    }
  } catch (err) {
    return {
      success: false,
      message: 'Error fetching rates: ' + err.message
    };
  }
}

/**
 * Export validation results from the log sheet.  The logs are
 * maintained in a Spreadsheet whose ID is stored in Script
 * Properties (VALIDATION_LOG_SHEET_ID).  You can call this
 * function from the client to download logs as CSV or JSON.
 *
 * @param {string} format 'csv' or 'json'
 * @return {Object} An object containing the exported data or an error
 */
function exportValidationResults(format) {
  try {
    var sheetId = PropertiesService.getScriptProperties().getProperty('VALIDATION_LOG_SHEET_ID');
    if (!sheetId) {
      return { success: false, message: 'No validation data available for export' };
    }
    var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    var data = sheet.getDataRange().getValues();
    if (format === 'csv') {
      var csv = data.map(function (row) { return row.join(','); }).join('\n');
      var blob = Utilities.newBlob(csv, 'text/csv', 'tar_validation_log.csv');
      return { success: true, data: blob, message: 'CSV export ready' };
    } else {
      var json = data.slice(1).map(function (row) {
        return {
          timestamp: row[0],
          traveler: row[1],
          authNumber: row[2],
          expectedCost: row[3],
          claimedCost: row[4],
          variance: row[5],
          variancePercent: row[6],
          status: row[7]
        };
      });
      return { success: true, data: json, message: 'JSON export ready' };
    }
  } catch (err) {
    return { success: false, message: 'Export failed: ' + err.message };
  }
}

/**
 * Clear all validation logs from the spreadsheet identified by
 * VALIDATION_LOG_SHEET_ID.  The header row is preserved.
 *
 * @return {Object} Status message indicating success or failure
 */
function clearValidationLogs() {
  try {
    var sheetId = PropertiesService.getScriptProperties().getProperty('VALIDATION_LOG_SHEET_ID');
    if (!sheetId) {
      return { success: false, message: 'No validation logs found to clear' };
    }
    var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    sheet.clear();
    sheet.getRange(1, 1, 1, 8).setValues([
      ['Timestamp','Traveler','Auth Number','Expected Cost','Claimed Cost','Variance','Variance %','Status']
    ]);
    return { success: true, message: 'Validation logs cleared successfully' };
  } catch (err) {
    return { success: false, message: 'Failed to clear logs: ' + err.message };
  }
}