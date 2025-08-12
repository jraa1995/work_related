/**
 * Enhanced utility functions for TAR validation system
 *
 * This module consolidates common utility functions used throughout the
 * application. It is written to be compatible with both Google Apps
 * Script and Node.js environments. When executed within Apps Script,
 * global objects such as `UrlFetchApp`, `Utilities`, `DriveApp`,
 * `DocumentApp`, and `Logger` are available. In a Node.js environment
 * (used primarily for unit tests), these services do not exist. To
 * support local testing the module conditionally imports the CONFIG
 * object and defines a no-op Logger when needed. Functions that rely
 * on Apps Script services will gracefully return `null` in Node.js.
 */

// Import configuration when running under Node.js. In Apps Script
// `module` is undefined and this block is ignored.
if (typeof module !== 'undefined' && module.exports) {
  // eslint-disable-next-line global-require
  var CONFIG = require('./config.js');
}

// Provide a stub Logger for Node.js. When running in Apps Script the
// real Logger object is available globally.
if (typeof Logger === 'undefined') {
  var Logger = {
    log: function () {},
    warn: function () {},
    error: function () {},
  };
}

/**
 * Utility to average numeric strings from GSA rate response
 */
function average(values) {
  /**
   * Convert a GSA rate value into a number. Rates may be provided as a
   * single number or as a range (e.g. "130-150"). When a range is
   * encountered we return the midpoint of the range. If a value cannot
   * be parsed it is ignored.
   */
  function parseRate(value) {
    if (value === undefined || value === null) return NaN;
    // If the value is a string containing a hyphen, treat it as a range
    if (typeof value === 'string' && value.includes('-')) {
      const [low, high] = value
        .split('-')
        .map((p) => parseFloat(p.trim()));
      if (!isNaN(low) && !isNaN(high)) {
        return (low + high) / 2;
      }
    }
    const n = parseFloat(value);
    return isNaN(n) ? NaN : n;
  }

  const nums = values
    .map((v) => parseRate(v))
    .filter((n) => !isNaN(n));
  const total = nums.reduce((a, b) => a + b, 0);
  return nums.length ? total / nums.length : 0;
}

/**
 * Fetch per diem rate from GSA API by city/state with enhanced error handling
 */
function fetchPerDiemByCityState(city, state, year = CONFIG.YEAR) {
  // Validate input to avoid calling methods on undefined/null. If either city or
  // state is missing we cannot perform a lookup and will return null.
  if (!city || !state) {
    return null;
  }
  const cleanedCity = encodeURIComponent(String(city).trim().replace(/[.'-]/g, ' '));
  const cleanedState = String(state).trim().toUpperCase();
  const url = `${CONFIG.GSA_BASE_URL}/rates/city/${cleanedCity}/state/${cleanedState}/year/${year}?api_key=${CONFIG.GSA_API_KEY}`;

  try {
    // Use Apps Script UrlFetchApp when available
    if (typeof UrlFetchApp !== 'undefined') {
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'User-Agent': 'TAR-Validation-System/1.0',
        },
        muteHttpExceptions: true,
      });

      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        return data.length ? data[0] : null;
      }
      // Log API error details when running in Apps Script
      if (typeof Logger !== 'undefined') {
        Logger.log(
          `GSA API Error: ${response.getResponseCode()} - ${response.getContentText()}`
        );
      }
      return null;
    } else if (typeof fetch !== 'undefined') {
      // In Node.js we can attempt to perform the request with fetch.
      // Note that fetch is asynchronous; however our function is synchronous by
      // design. To keep compatibility we perform a synchronous HTTP request
      // using child_process and curl. If curl is unavailable or the request
      // fails we simply return null.
      const execSync = require('child_process').execSync;
      try {
        const curlCmd = `curl -s -H "Accept: application/json" -A "TAR-Validation-System/1.0" "${url}"`;
        const output = execSync(curlCmd, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
        const data = JSON.parse(output);
        return data.length ? data[0] : null;
      } catch (e) {
        return null;
      }
    }
    return null;
  } catch (error) {
    // Log when running in Apps Script
    if (typeof Logger !== 'undefined') {
      Logger.log(`GSA fetch error: ${error.message}`);
    }
    return null;
  }
}

/**
 * Enhanced PDF text extraction using Google Apps Script
 */
function extractTextFromPDF(base64Data) {
  try {
    // In Google Apps Script use Drive and Document services
    if (typeof Utilities !== 'undefined' && typeof DriveApp !== 'undefined') {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64Data),
        'application/pdf',
        'temp_tar_document.pdf'
      );
      const tempFile = DriveApp.createFile(blob);
      const resource = {
        title: 'temp_conversion',
        mimeType: MimeType.GOOGLE_DOCS,
      };
      const docFile = Drive.Files.copy(resource, tempFile.getId(), {
        convert: true,
      });
      const doc = DocumentApp.openById(docFile.id);
      const text = doc.getBody().getText();
      // Clean up temporary files
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      DriveApp.getFileById(docFile.id).setTrashed(true);
      return text;
    }
    // In a Node.js environment PDF extraction is not supported in this
    // simplified implementation. Return null to indicate extraction failed.
    return null;
  } catch (error) {
    if (typeof Logger !== 'undefined') {
      Logger.log(`PDF extraction error: ${error.message}`);
    }
    return null;
  }
}

/**
 * Extract GSA form data from document text
 */
function extractGSAFormData(text) {
  if (!text) return {};

  const extracted = {};

  // Extract using configured patterns
  for (const [field, pattern] of Object.entries(CONFIG.GSA_FIELD_PATTERNS)) {
    const match = text.match(pattern);
    if (match && match[1]) {
      extracted[field] = match[1].trim();
    }
  }

  // Extract itinerary data
  extracted.itinerary = extractItineraryData(text);

  // Parse numeric fields
  if (extracted.estimatedCost) {
    extracted.estimatedCost = parseFloat(
      extracted.estimatedCost.replace(/,/g, "")
    );
  }
  if (extracted.perDiem) {
    extracted.perDiem = parseFloat(extracted.perDiem.replace(/,/g, ""));
  }

  return extracted;
}

/**
 * Extract itinerary information from document text
 */
function extractItineraryData(text) {
  const itinerary = [];

  // Look for itinerary section
  const itinerarySection = text.match(
    /AUTHORIZED OFFICIAL ITINERARY([\s\S]*?)(?=\n\n|\n[A-Z])/i
  );
  if (!itinerarySection) return itinerary;

  const section = itinerarySection[1];

  // Extract dates
  const dates = [...section.matchAll(CONFIG.ITINERARY_PATTERNS.datePattern)];

  // Extract city/state pairs
  const locations = [
    ...section.matchAll(CONFIG.ITINERARY_PATTERNS.cityStatePattern),
  ];

  // Combine dates and locations
  for (let i = 0; i < Math.min(dates.length, locations.length); i++) {
    if (dates[i] && locations[i]) {
      itinerary.push({
        date: dates[i][1],
        city: locations[i][1].trim(),
        state: locations[i][2].trim(),
      });
    }
  }

  return itinerary;
}

/**
 * Validate extracted form data
 */
function validateFormData(data) {
  const errors = [];
  // Check required fields
  CONFIG.VALIDATION_RULES.requiredFields.forEach((field) => {
    const value = data[field];
    const isString = typeof value === 'string';
    const isEmptyString = isString && value.trim() === '';
    if (value === undefined || value === null || isEmptyString) {
      errors.push(`Missing required field: ${field}`);
    }
  });

  // Validate vendor code format. Normalize to string to avoid errors when numbers are passed
  if (data.vendorCode) {
    const vendor = String(data.vendorCode).trim();
    if (!CONFIG.VALIDATION_RULES.vendorCodeFormat.test(vendor)) {
      errors.push('Invalid vendor code format');
    }
  }

  // Validate phone number: strip common separators and convert to string
  if (data.contactNumber) {
    const phoneStr = String(data.contactNumber).replace(/[\s\-\(\)]/g, '');
    if (!CONFIG.VALIDATION_RULES.phoneFormat.test(phoneStr)) {
      errors.push('Invalid phone number format');
    }
  }

  // Validate numeric fields
  if (
    data.estimatedCost &&
    (isNaN(data.estimatedCost) || data.estimatedCost <= 0)
  ) {
    errors.push('Invalid estimated cost');
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
}

/**
 * Calculate expected costs based on itinerary
 */
function calculateExpectedCosts(itinerary) {
  let totalExpected = 0;
  const breakdown = [];

  itinerary.forEach((item) => {
    const rateData = fetchPerDiemByCityState(item.city, item.state);
    let mie = CONFIG.DEFAULT_MIE;
    let lodging = CONFIG.DEFAULT_LODGING;

    if (rateData) {
      mie = parseFloat(rateData.Meals) || CONFIG.DEFAULT_MIE;
      lodging = average([
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
    }

    const dailyTotal = mie + lodging;
    totalExpected += dailyTotal;

    breakdown.push({
      location: `${item.city}, ${item.state}`,
      date: item.date,
      mie: mie,
      lodging: lodging,
      total: dailyTotal,
    });
  });

  return {
    totalExpected: totalExpected,
    breakdown: breakdown,
  };
}

/**
 * Format currency values
 */
function formatCurrency(amount) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(amount);
}

/**
 * Generate detailed validation report
 */
function generateValidationReport(extractedData, expectedCosts, claimedCost) {
  const report = {
    timestamp: new Date().toISOString(),
    traveler: extractedData.travelerName || "Unknown",
    authorizationNumber: extractedData.authorizationNumber || "N/A",
    validation: {
      extractedData: extractedData,
      expectedCosts: expectedCosts,
      claimedCost: claimedCost,
      variance: claimedCost - expectedCosts.totalExpected,
      variancePercent: (
        ((claimedCost - expectedCosts.totalExpected) /
          expectedCosts.totalExpected) *
        100
      ).toFixed(2),
      isWithinBuffer:
        Math.abs(claimedCost - expectedCosts.totalExpected) <=
        CONFIG.COST_BUFFER,
      isWithinDeviation:
        (Math.abs(claimedCost - expectedCosts.totalExpected) /
          expectedCosts.totalExpected) *
          100 <=
        CONFIG.MAX_DEVIATION_PERCENT,
    },
    recommendations: [],
  };

  // Add recommendations based on validation results
  if (report.validation.variance > CONFIG.COST_BUFFER) {
    report.recommendations.push(
      "Claimed cost exceeds expected amount by more than acceptable buffer"
    );
  }

  if (
    Math.abs(report.validation.variancePercent) > CONFIG.MAX_DEVIATION_PERCENT
  ) {
    report.recommendations.push(
      `Variance of ${report.validation.variancePercent}% exceeds maximum acceptable deviation`
    );
  }

  if (expectedCosts.breakdown.length === 0) {
    report.recommendations.push(
      "No itinerary data found - manual review required"
    );
  }

  return report;
}

/**
 * Log validation results to spreadsheet (optional)
 */
function logValidationResults(report) {
  try {
    // Create or get existing spreadsheet for logging
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty(
      "VALIDATION_LOG_SHEET_ID"
    );
    let sheet;

    if (spreadsheetId) {
      sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
    } else {
      const newSheet = SpreadsheetApp.create("TAR Validation Log");
      sheet = newSheet.getActiveSheet();
      PropertiesService.getScriptProperties().setProperty(
        "VALIDATION_LOG_SHEET_ID",
        newSheet.getId()
      );

      // Add headers
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
    }

    // Add log entry
    sheet.appendRow([
      report.timestamp,
      report.traveler,
      report.authorizationNumber,
      report.validation.expectedCosts.totalExpected,
      report.validation.claimedCost,
      report.validation.variance,
      report.validation.variancePercent,
      report.validation.isWithinBuffer ? "APPROVED" : "NEEDS REVIEW",
    ]);
  } catch (error) {
    Logger.log(`Logging error: ${error.message}`);
  }
}

// Export utilities for Node.js environments. Google Apps Script doesn't define
// `module` so the following block will be ignored when deployed there. This
// allows unit tests to import and exercise utility functions locally.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    average,
    fetchPerDiemByCityState,
    extractTextFromPDF,
    extractGSAFormData,
    extractItineraryData,
    validateFormData,
    calculateExpectedCosts,
    formatCurrency,
    generateValidationReport,
    logValidationResults,
  };
}
