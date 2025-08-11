/**
 * Enhanced global configuration for TAR validation system
 */

const CONFIG = {
  // GSA API Configuration
  GSA_API_KEY: "UtxSKaLNOnlkWu8z6huLrCKfgkYsMd36OXFiaAf3",
  GSA_BASE_URL: "https://api.gsa.gov/travel/perdiem/v2",
  YEAR: "2025", // fiscal year for lookup

  // Default rates (fallback values)
  DEFAULT_MIE: 79, // default M&IE value fallback
  DEFAULT_LODGING: 150, // default lodging fallback

  // Validation thresholds
  COST_BUFFER: 10, // acceptable buffer in overage (USD)
  MAX_DEVIATION_PERCENT: 15, // max acceptable deviation percentage

  // PDF/Document processing
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB max file size
  SUPPORTED_FORMATS: ["pdf", "docx", "doc"],

  // GSA Form field mappings for extraction
  GSA_FIELD_PATTERNS: {
    authorizationNumber: /authorization\s+number[:\s]*([^\n\r]+)/i,
    travelerName: /traveler[:\s]*([^\n\r]+)/i,
    title: /title[:\s]*([^\n\r]+)/i,
    vendorCode: /pegasys\s+vendor\s+code[:\s]*([^\n\r]+)/i,
    currentAddress: /current\s+residence\s+address[:\s]*([^\n\r]+)/i,
    officeDivision: /office\/service\s+and\s+division[:\s]*([^\n\r]+)/i,
    dutyStation: /official\s+duty\s+station[:\s]*([^\n\r]+)/i,
    contactNumber: /contact\s+telephone\s+number[:\s]*([^\n\r]+)/i,
    travelPurpose: /travel\s+purpose[:\s]*([^\n\r]+)/i,
    briefDescription: /brief\s+description[:\s]*([^\n\r]+)/i,
    estimatedCost: /estimated\s+cost[:\s]*total[:\s]*\$?([0-9,]+\.?\d*)/i,
    perDiem: /per\s+diem[:\s]*\$?([0-9,]+\.?\d*)/i,
    airRail: /air\/rail[:\s]*\$?([0-9,]+\.?\d*)/i,
    other: /other[:\s]*\$?([0-9,]+\.?\d*)/i,
  },

  // Itinerary extraction patterns
  ITINERARY_PATTERNS: {
    datePattern: /(\d{1,2}\/\d{1,2}\/\d{4})/g,
    cityStatePattern: /([A-Za-z\s]+),\s*([A-Z]{2})/g,
    lodgingPattern: /\$?(\d+\.?\d*)/g,
    miePattern: /\$?(\d+\.?\d*)/g,
  },

  // Validation rules
  VALIDATION_RULES: {
    requiredFields: [
      "travelerName",
      "travelPurpose",
      "estimatedCost",
      "dutyStation",
      "contactNumber",
    ],
    dateFormat: /^\d{4}-\d{2}-\d{2}$/,
    phoneFormat: /^[\+]?[1-9][\d]{0,15}$/,
    vendorCodeFormat: /^E\d{8,9}$/,
  },

  // Error messages
  ERROR_MESSAGES: {
    fileTooBig: "File size exceeds maximum limit of 10MB",
    unsupportedFormat:
      "Unsupported file format. Please upload PDF or Word document",
    extractionFailed: "Failed to extract data from document",
    validationFailed: "Document validation failed",
    apiError: "Error fetching GSA per diem rates",
    missingRequiredFields: "Missing required fields",
  },

  // Expense limits for client‑side validations
  // These values mirror policy thresholds enforced in the UI. They are
  // duplicated here so that both Apps Script and client code can share the
  // same constants. Modify values here to adjust validation behaviour.
  EXPENSE_LIMITS: {
    // Maximum daily cost allowed for car rentals; exceeding this triggers an error
    CAR_RENTAL_MAX_DAILY: 75,
    // Threshold beyond which car rental costs require business justification
    CAR_RENTAL_JUSTIFICATION_THRESHOLD: 400,
    // Daily parking cost limit; exceeding this triggers a warning
    PARKING_MAX_DAILY: 25,
    // Maximum conference or training fee; exceeding this triggers an error
    CONFERENCE_FEE_MAX: 2000,
    // Threshold beyond which conference or training fees require pre‑approval
    CONFERENCE_JUSTIFICATION_THRESHOLD: 1000,
    // Threshold beyond which miscellaneous expenses require detailed receipts
    MISC_JUSTIFICATION_THRESHOLD: 200,
    // Acceptable variance between calculated and claimed totals, expressed as a decimal
    TOTAL_VARIANCE_THRESHOLD: 0.15
  },
};
