/**
 * utils.gs
 *
 * Common utility functions used by both the client and server
 * components of the Enhanced TAR Validation System.  These
 * functions handle rate calculations, API calls, PDF extraction,
 * itinerary parsing, form validation, expected cost calculation,
 * report generation, and logging.  All Node.js specific code has
 * been removed so that the functions run natively in the Google
 * Apps Script environment.
 */

/**
 * Average an array of GSA rate values.  Each value may be a
 * number, a string, or a range like "130-150".  Ranges are
 * converted into their midpoints.  Invalid values are ignored.
 *
 * @param {Array} values A collection of rates
 * @return {number} The average of all valid numeric entries
 */
function average(values) {
  function parseRate(value) {
    if (value === undefined || value === null) return NaN;
    if (typeof value === 'string' && value.indexOf('-') !== -1) {
      var parts = value.split('-');
      var low = parseFloat(parts[0].trim());
      var high = parseFloat(parts[1].trim());
      if (!isNaN(low) && !isNaN(high)) {
        return (low + high) / 2;
      }
    }
    var n = parseFloat(value);
    return isNaN(n) ? NaN : n;
  }
  var nums = values.map(parseRate).filter(function (n) { return !isNaN(n); });
  var total = nums.reduce(function (a, b) { return a + b; }, 0);
  return nums.length ? total / nums.length : 0;
}

/**
 * Fetch per diem rate data from the GSA API.  This function
 * requires UrlFetchApp to be available (which it is in Apps
 * Script).  It returns the first rate object from the API response
 * or null on failure.  Only the city, state and year parameters
 * are used.
 *
 * @param {string} city Destination city
 * @param {string} state Destination state
 * @param {string} [year] Fiscal year (defaults to CONFIG.YEAR)
 * @return {Object|null} The rate object or null
 */
function fetchPerDiemByCityState(city, state, year) {
  if (!city || !state) return null;
  year = year || CONFIG.YEAR;
  var cleanedCity = encodeURIComponent(String(city).trim().replace(/[.'-]/g, ' '));
  var cleanedState = String(state).trim().toUpperCase();
  var url = CONFIG.GSA_BASE_URL + '/rates/city/' + cleanedCity + '/state/' + cleanedState + '/year/' + year + '?api_key=' + CONFIG.GSA_API_KEY;
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Accept: 'application/json', 'User-Agent': 'TAR-Validation-System/1.0' },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      return data.length ? data[0] : null;
    }
    Logger.log('GSA API Error: ' + response.getResponseCode() + ' - ' + response.getContentText());
    return null;
  } catch (err) {
    Logger.log('GSA fetch error: ' + err.message);
    return null;
  }
}

/**
 * Extract plain text from a base64 encoded PDF.  In Apps Script this
 * uses DriveApp and DocumentApp to convert the PDF into a Google
 * Doc and read its contents.  Temporary files are created and
 * cleaned up.  Returns null on failure.
 *
 * @param {string} base64Data The PDF encoded as base64
 * @return {string|null} The extracted text or null
 */
function extractTextFromPDF(base64Data) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', 'temp_tar_document.pdf');
    var tempFile = DriveApp.createFile(blob);
    var resource = { title: 'temp_conversion', mimeType: MimeType.GOOGLE_DOCS };
    var docFile = Drive.Files.copy(resource, tempFile.getId(), { convert: true });
    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();
    // Clean up temporary files
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    DriveApp.getFileById(docFile.id).setTrashed(true);
    return text;
  } catch (err) {
    Logger.log('PDF extraction error: ' + err.message);
    return null;
  }
}

/**
 * Extract GSA form data from raw text.  Uses CONFIG.GSA_FIELD_PATTERNS
 * to pull key/value pairs and extracts itinerary information via
 * extractItineraryData().  Numeric fields are parsed into numbers.
 *
 * @param {string} text The document text
 * @return {Object} An object containing extracted fields and itinerary
 */
function extractGSAFormData(text) {
  if (!text) return {};
  var extracted = {};
  // Extract fields based on patterns
  for (var key in CONFIG.GSA_FIELD_PATTERNS) {
    var pattern = CONFIG.GSA_FIELD_PATTERNS[key];
    var match = text.match(pattern);
    if (match && match[1]) {
      extracted[key] = match[1].trim();
    }
  }
  // Extract itinerary
  extracted.itinerary = extractItineraryData(text);
  // Convert numeric fields
  if (extracted.estimatedCost) {
    extracted.estimatedCost = parseFloat(extracted.estimatedCost.replace(/,/g, ''));
  }
  if (extracted.perDiem) {
    extracted.perDiem = parseFloat(extracted.perDiem.replace(/,/g, ''));
  }
  return extracted;
}

/**
 * Parse itinerary information from free text.  Looks for a section
 * labelled "AUTHORIZED OFFICIAL ITINERARY" and then extracts dates
 * and city/state pairs using patterns defined in CONFIG.ITINERARY_PATTERNS.
 *
 * @param {string} text Raw document text
 * @return {Array<Object>} List of itinerary entries with date, city, state
 */
function extractItineraryData(text) {
  var itinerary = [];
  var secMatch = text.match(/AUTHORIZED OFFICIAL ITINERARY([\s\S]*?)(?=\n\n|\n[A-Z])/i);
  if (!secMatch) return itinerary;
  var section = secMatch[1];
  var dates = section.matchAll(CONFIG.ITINERARY_PATTERNS.datePattern);
  var locations = section.matchAll(CONFIG.ITINERARY_PATTERNS.cityStatePattern);
  var dateArr = Array.from(dates);
  var locArr = Array.from(locations);
  for (var i = 0; i < Math.min(dateArr.length, locArr.length); i++) {
    itinerary.push({
      date: dateArr[i][1],
      city: locArr[i][1].trim(),
      state: locArr[i][2].trim()
    });
  }
  return itinerary;
}

/**
 * Validate TAR form data.  Ensures required fields are present and
 * formats vendor codes and phone numbers.  Returns an object with
 * isValid flag and an array of error messages.
 *
 * @param {Object} data The data to validate
 * @return {Object} Validation result
 */
function validateFormData(data) {
  var errors = [];
  // Required fields
  CONFIG.VALIDATION_RULES.requiredFields.forEach(function (field) {
    var value = data[field];
    var isString = typeof value === 'string';
    var isEmpty = isString && value.trim() === '';
    if (value === undefined || value === null || isEmpty) {
      errors.push('Missing required field: ' + field);
    }
  });
  // Vendor code
  if (data.vendorCode) {
    var vendor = String(data.vendorCode).trim();
    if (!CONFIG.VALIDATION_RULES.vendorCodeFormat.test(vendor)) {
      errors.push('Invalid vendor code format');
    }
  }
  // Phone number
  if (data.contactNumber) {
    var phone = String(data.contactNumber).replace(/[\s\-\(\)]/g, '');
    if (!CONFIG.VALIDATION_RULES.phoneFormat.test(phone)) {
      errors.push('Invalid phone number format');
    }
  }
  // Estimated cost
  if (data.estimatedCost && (isNaN(data.estimatedCost) || data.estimatedCost <= 0)) {
    errors.push('Invalid estimated cost');
  }
  return { isValid: errors.length === 0, errors: errors };
}

/**
 * Calculate expected per diem costs from an itinerary.  For each
 * itinerary entry we look up the rate and sum the M&IE and
 * lodging.  Returns the total expected cost and a breakdown.
 *
 * @param {Array} itinerary List of itinerary stops
 * @return {Object} Object with totalExpected and breakdown
 */
function calculateExpectedCosts(itinerary) {
  var totalExpected = 0;
  var breakdown = [];
  itinerary.forEach(function (item) {
    var rate = fetchPerDiemByCityState(item.city, item.state);
    var mie = CONFIG.DEFAULT_MIE;
    var lodging = CONFIG.DEFAULT_LODGING;
    if (rate) {
      mie = parseFloat(rate.Meals) || CONFIG.DEFAULT_MIE;
      lodging = average([
        rate.Jan, rate.Feb, rate.Mar, rate.Apr, rate.May, rate.Jun,
        rate.Jul, rate.Aug, rate.Sep, rate.Oct, rate.Nov, rate.Dec
      ]);
    }
    var dailyTotal = mie + lodging;
    totalExpected += dailyTotal;
    breakdown.push({ location: item.city + ', ' + item.state, date: item.date, mie: mie, lodging: lodging, total: dailyTotal });
  });
  return { totalExpected: totalExpected, breakdown: breakdown };
}

/**
 * Format a number into US currency.  Uses the built‑in Intl API.
 *
 * @param {number} amount The value to format
 * @return {string} Formatted currency string
 */
function formatCurrency(amount) {
  return Utilities.formatString('$%s', amount.toFixed(2));
}

/**
 * Generate a detailed validation report.  Combines expected costs,
 * claimed cost, and extracted data into a structured object with
 * variance calculations and recommendations.  See main.gs for how
 * this report is used.
 *
 * @param {Object} extractedData Data extracted from form or document
 * @param {Object} expectedCosts Output of calculateExpectedCosts()
 * @param {number} claimedCost The claimed trip cost
 * @return {Object} Validation report
 */
function generateValidationReport(extractedData, expectedCosts, claimedCost) {
  var variance = claimedCost - expectedCosts.totalExpected;
  var variancePercent = expectedCosts.totalExpected === 0 ? 0 : ((variance / expectedCosts.totalExpected) * 100);
  var report = {
    timestamp: new Date().toISOString(),
    traveler: extractedData.travelerName || 'Unknown',
    authorizationNumber: extractedData.authorizationNumber || 'N/A',
    validation: {
      extractedData: extractedData,
      expectedCosts: expectedCosts,
      claimedCost: claimedCost,
      variance: variance,
      variancePercent: variancePercent.toFixed(2),
      isWithinBuffer: Math.abs(variance) <= CONFIG.COST_BUFFER,
      isWithinDeviation: Math.abs(variancePercent) <= CONFIG.MAX_DEVIATION_PERCENT
    },
    recommendations: []
  };
  // Build recommendations
  if (report.validation.isWithinBuffer === false) {
    report.recommendations.push('Claimed cost exceeds expected amount by more than acceptable buffer');
  }
  if (report.validation.isWithinDeviation === false) {
    report.recommendations.push('Variance of ' + report.validation.variancePercent + '% exceeds maximum acceptable deviation');
  }
  if (expectedCosts.breakdown.length === 0) {
    report.recommendations.push('No itinerary data found - manual review required');
  }
  return report;
}

/**
 * Log validation results into a spreadsheet.  The spreadsheet ID
 * must be stored in Script Properties under the key
 * VALIDATION_LOG_SHEET_ID.  If the sheet does not exist it will
 * be created and a header row added.
 *
 * @param {Object} report The validation report
 */
function logValidationResults(report) {
  try {
    var sheetId = PropertiesService.getScriptProperties().getProperty('VALIDATION_LOG_SHEET_ID');
    var sheet;
    if (sheetId) {
      sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    }
    if (!sheet) {
      var ss = SpreadsheetApp.create('TAR Validation Logs');
      sheet = ss.getActiveSheet();
      PropertiesService.getScriptProperties().setProperty('VALIDATION_LOG_SHEET_ID', ss.getId());
      sheet.appendRow(['Timestamp','Traveler','Auth Number','Expected Cost','Claimed Cost','Variance','Variance %','Status']);
    }
    var status = report.validation.isWithinBuffer && report.validation.isWithinDeviation ? 'Valid' : 'Invalid';
    sheet.appendRow([
      report.timestamp,
      report.traveler,
      report.authorizationNumber,
      report.validation.expectedCosts.totalExpected.toFixed(2),
      report.validation.claimedCost.toFixed(2),
      report.validation.variance.toFixed(2),
      report.validation.variancePercent,
      status
    ]);
  } catch (err) {
    Logger.log('Logging error: ' + err.message);
  }
}