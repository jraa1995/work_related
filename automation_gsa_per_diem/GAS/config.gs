/**
 * config.gs
 *
 * Defines global configuration constants used by both client and
 * server code.  These values include API keys, default per diem
 * rates, expense thresholds and patterns used for document parsing.
 * By defining CONFIG in a separate file it becomes easy to update
 * thresholds or API settings without touching other code.
 */

const CONFIG = {
  // GSA API configuration
  GSA_API_KEY: 'UtxSKaLNOnlkWu8z6huLrCKfgkYsMd36OXFiaAf3',
  GSA_BASE_URL: 'https://api.gsa.gov/travel/perdiem/v2',
  YEAR: '2025',

  // Default fallback rates when GSA data is unavailable
  DEFAULT_MIE: 79,
  DEFAULT_LODGING: 150,

  // Validation thresholds
  COST_BUFFER: 10, // acceptable absolute buffer in USD
  MAX_DEVIATION_PERCENT: 15, // max allowed deviation percent

  // File upload limits
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  SUPPORTED_FORMATS: ['pdf', 'docx', 'doc'],

  // Patterns for extracting fields from GSA Form 87
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
    other: /other[:\s]*\$?([0-9,]+\.?\d*)/i
  },

  // Patterns used when parsing itineraries from free text
  ITINERARY_PATTERNS: {
    datePattern: /(\d{1,2}\/\d{1,2}\/\d{4})/g,
    cityStatePattern: /([A-Za-z\s]+),\s*([A-Z]{2})/g,
    lodgingPattern: /\$?(\d+\.?\d*)/g,
    miePattern: /\$?(\d+\.?\d*)/g
  },

  // Validation rules and patterns for form data
  VALIDATION_RULES: {
    requiredFields: ['travelerName','travelPurpose','estimatedCost','dutyStation','contactNumber'],
    dateFormat: /^\d{4}-\d{2}-\d{2}$/,
    phoneFormat: /^[\+]?\d{7,15}$/,
    vendorCodeFormat: /^E\d{8,9}$/
  },

  // Standard error messages
  ERROR_MESSAGES: {
    fileTooBig: 'File size exceeds maximum limit of 10MB',
    unsupportedFormat: 'Unsupported file format. Please upload PDF or Word document',
    extractionFailed: 'Failed to extract data from document',
    validationFailed: 'Document validation failed',
    apiError: 'Error fetching GSA per diem rates',
    missingRequiredFields: 'Missing required fields'
  },

  // Expense limits used by client‑side validations
  EXPENSE_LIMITS: {
    CAR_RENTAL_MAX_DAILY: 75,
    CAR_RENTAL_JUSTIFICATION_THRESHOLD: 400,
    PARKING_MAX_DAILY: 25,
    CONFERENCE_FEE_MAX: 2000,
    CONFERENCE_JUSTIFICATION_THRESHOLD: 1000,
    MISC_JUSTIFICATION_THRESHOLD: 200,
    TOTAL_VARIANCE_THRESHOLD: 0.15
  }
};