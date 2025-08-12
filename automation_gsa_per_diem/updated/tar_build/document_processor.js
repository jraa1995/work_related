/**
 * Enhanced document processing utilities for TAR validation system
 * Handles PDF and Word document extraction with improved accuracy
 */

// Provide a stub Logger in environments where the Google Apps Script Logger
// service is unavailable (e.g. Node.js during unit testing). If Logger is
// already defined, this block has no effect.
if (typeof Logger === 'undefined') {
  var Logger = {
    log: function () {},
    warn: function () {},
    error: function () {},
  };
}

// Import utility functions when running under Node.js. In Apps Script
// environment `require` is undefined and this block is ignored. We
// destructure only the functions used within this file.
if (typeof module !== 'undefined' && module.exports) {
  const utils = require('./utils.js');
  // Use the PDF extraction helper from utils when running under Node. The
  // `extractTextFromWord` and other functions are defined in this file
  // below and do not need to be imported.
  var extractTextFromPDF = utils.extractTextFromPDF;
}

/**
 * Main document processing function
 */
function processDocument(base64Data, mimeType, filename) {
  try {
    Logger.log(`Processing document: ${filename} (${mimeType})`);

    let extractedText = "";

    if (mimeType === "application/pdf") {
      extractedText = extractTextFromPDF(base64Data);
    } else if (
      mimeType ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
      mimeType === "application/msword"
    ) {
      extractedText = extractTextFromWord(base64Data);
    } else {
      throw new Error(`Unsupported document type: ${mimeType}`);
    }

    if (!extractedText) {
      throw new Error("No text could be extracted from the document");
    }

    // Clean and process the extracted text
    const cleanedText = cleanExtractedText(extractedText);

    // Extract structured data from the text
    const extractedData = extractGSAFormDataEnhanced(cleanedText);

    // Validate extraction quality
    const quality = validateExtractionQuality(extractedData);

    return {
      success: true,
      extractedText: cleanedText,
      extractedData: extractedData,
      quality: quality,
      metadata: {
        filename: filename,
        mimeType: mimeType,
        extractionMethod: mimeType === "application/pdf" ? "PDF" : "Word",
        textLength: cleanedText.length,
        confidence: quality.confidence,
      },
    };
  } catch (error) {
    Logger.log(`Document processing error: ${error.message}`);
    return {
      success: false,
      error: error.message,
      extractedData: {},
      quality: { confidence: "FAILED", issues: [error.message] },
    };
  }
}

/**
 * Enhanced PDF text extraction with OCR fallback
 */
function extractTextFromPDF(base64Data) {
  try {
    // Method 1: Direct PDF to Google Doc conversion
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/pdf",
      `temp_pdf_${Date.now()}.pdf`
    );

    const tempFile = DriveApp.createFile(blob);

    try {
      // Convert to Google Doc for text extraction
      const resource = {
        title: `temp_conversion_${Date.now()}`,
        mimeType: MimeType.GOOGLE_DOCS,
      };

      const docFile = Drive.Files.copy(resource, tempFile.getId(), {
        convert: true,
      });

      const doc = DocumentApp.openById(docFile.id);
      const extractedText = doc.getBody().getText();

      // Clean up temporary files
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      DriveApp.getFileById(docFile.id).setTrashed(true);

      if (extractedText && extractedText.trim().length > 50) {
        return extractedText;
      }
    } catch (conversionError) {
      Logger.log(`PDF conversion failed: ${conversionError.message}`);
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    }

    // Method 2: OCR fallback using Drive API
    try {
      return extractTextWithOCR(base64Data);
    } catch (ocrError) {
      Logger.log(`OCR extraction failed: ${ocrError.message}`);
      throw new Error("All PDF extraction methods failed");
    }
  } catch (error) {
    Logger.log(`PDF extraction error: ${error.message}`);
    return null;
  }
}

/**
 * OCR text extraction as fallback for PDFs
 */
function extractTextWithOCR(base64Data) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/pdf",
      `ocr_temp_${Date.now()}.pdf`
    );

    const file = DriveApp.createFile(blob);

    // Use Drive API to get OCR text
    const ocrResource = {
      title: `ocr_conversion_${Date.now()}`,
      mimeType: MimeType.GOOGLE_DOCS,
    };

    const ocrFile = Drive.Files.copy(ocrResource, file.getId(), {
      convert: true,
      ocr: true,
      ocrLanguage: "en",
    });

    const doc = DocumentApp.openById(ocrFile.id);
    const text = doc.getBody().getText();

    // Clean up
    DriveApp.getFileById(file.getId()).setTrashed(true);
    DriveApp.getFileById(ocrFile.id).setTrashed(true);

    return text;
  } catch (error) {
    Logger.log(`OCR error: ${error.message}`);
    throw error;
  }
}

/**
 * Enhanced Word document text extraction
 */
function extractTextFromWord(base64Data) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      `temp_word_${Date.now()}.docx`
    );

    const tempFile = DriveApp.createFile(blob);

    try {
      // Convert Word to Google Doc
      const resource = {
        title: `word_conversion_${Date.now()}`,
        mimeType: MimeType.GOOGLE_DOCS,
      };

      const docFile = Drive.Files.copy(resource, tempFile.getId(), {
        convert: true,
      });

      const doc = DocumentApp.openById(docFile.id);
      const extractedText = doc.getBody().getText();

      // Clean up temporary files
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      DriveApp.getFileById(docFile.id).setTrashed(true);

      return extractedText;
    } catch (conversionError) {
      Logger.log(`Word conversion failed: ${conversionError.message}`);
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      throw conversionError;
    }
  } catch (error) {
    Logger.log(`Word extraction error: ${error.message}`);
    return null;
  }
}

/**
 * Clean and normalize extracted text
 */
function cleanExtractedText(text) {
  if (!text) return "";

  return (
    text
      // Normalize whitespace
      .replace(/\s+/g, " ")
      // Remove extra line breaks
      .replace(/\n\s*\n/g, "\n")
      // Clean up common OCR artifacts
      .replace(/[^\w\s\$\.\,\(\)\-\:\/\&\#]/g, "")
      // Normalize currency symbols
      .replace(/\$\s+/g, "$")
      // Fix common OCR mistakes for numbers
      .replace(/[Oo](?=\d)/g, "0")
      .replace(/[Il](?=\d)/g, "1")
      .replace(/[Ss](?=\d)/g, "5")
      .trim()
  );
}

/**
 * Enhanced GSA form data extraction with multiple patterns
 */
function extractGSAFormDataEnhanced(text) {
  if (!text) return {};

  const extracted = {};

  // Enhanced field extraction with multiple pattern attempts
  const fieldExtractions = {
    authorizationNumber: [
      /authorization\s+number[:\s]*([^\n\r]+)/i,
      /auth\s*#[:\s]*([^\n\r]+)/i,
      /authorization[:\s]*([A-Z0-9\-\/]+)/i,
    ],

    travelerName: [
      /traveler[:\s]*([^\n\r]+)/i,
      /employee\s+name[:\s]*([^\n\r]+)/i,
      /name[:\s]*([A-Za-z\s,\.]+)(?=\s|$)/i,
    ],

    title: [
      /title[:\s]*([^\n\r]+)/i,
      /position[:\s]*([^\n\r]+)/i,
      /job\s+title[:\s]*([^\n\r]+)/i,
    ],

    vendorCode: [
      /pegasys\s+vendor\s+code[:\s]*([E]\d{8,9})/i,
      /vendor\s+code[:\s]*([E]\d{8,9})/i,
      /employee\s+id[:\s]*([E]\d{8,9})/i,
    ],

    currentAddress: [
      /current\s+residence\s+address[:\s]*([^\n\r]+)/i,
      /address[:\s]*([^\n\r]+)/i,
      /residence[:\s]*([^\n\r]+)/i,
    ],

    officeDivision: [
      /office\/service\s+and\s+division[:\s]*([^\n\r]+)/i,
      /office[:\s]*([^\n\r]+)/i,
      /division[:\s]*([^\n\r]+)/i,
    ],

    dutyStation: [
      /official\s+duty\s+station[:\s]*([^\n\r]+)/i,
      /duty\s+station[:\s]*([^\n\r]+)/i,
      /work\s+location[:\s]*([^\n\r]+)/i,
    ],

    contactNumber: [
      /contact\s+telephone\s+number[:\s]*([^\n\r]+)/i,
      /phone[:\s]*([^\n\r]+)/i,
      /telephone[:\s]*([^\n\r]+)/i,
      /(\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4})/,
    ],

    travelPurpose: [
      /travel\s+purpose[:\s]*([^\n\r]+)/i,
      /purpose[:\s]*([^\n\r]+)/i,
      /reason\s+for\s+travel[:\s]*([^\n\r]+)/i,
    ],

    briefDescription: [
      /brief\s+description[:\s]*([^\n\r]+)/i,
      /description[:\s]*([^\n\r]+)/i,
      /details[:\s]*([^\n\r]+)/i,
    ],

    estimatedCost: [
      /estimated\s+cost[:\s]*total[:\s]*\$?([0-9,]+\.?\d*)/i,
      /total\s+cost[:\s]*\$?([0-9,]+\.?\d*)/i,
      /amount[:\s]*\$?([0-9,]+\.?\d*)/i,
    ],

    perDiem: [
      /per\s+diem[:\s]*\$?([0-9,]+\.?\d*)/i,
      /meals\s+and\s+incidentals[:\s]*\$?([0-9,]+\.?\d*)/i,
    ],

    airRail: [
      /air\/rail[:\s]*\$?([0-9,]+\.?\d*)/i,
      /transportation[:\s]*\$?([0-9,]+\.?\d*)/i,
      /airfare[:\s]*\$?([0-9,]+\.?\d*)/i,
    ],

    lodging: [
      /lodging[:\s]*\$?([0-9,]+\.?\d*)/i,
      /hotel[:\s]*\$?([0-9,]+\.?\d*)/i,
      /accommodation[:\s]*\$?([0-9,]+\.?\d*)/i,
    ],

    rentalCar: [
      /rental\s+car[:\s]*\$?([0-9,]+\.?\d*)/i,
      /car\s+rental[:\s]*\$?([0-9,]+\.?\d*)/i,
      /vehicle[:\s]*\$?([0-9,]+\.?\d*)/i,
    ],

    miscellaneous: [
      /miscellaneous[:\s]*\$?([0-9,]+\.?\d*)/i,
      /other[:\s]*\$?([0-9,]+\.?\d*)/i,
      /misc[:\s]*\$?([0-9,]+\.?\d*)/i,
    ],
  };

  // Attempt extraction with multiple patterns
  for (const [field, patterns] of Object.entries(fieldExtractions)) {
    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match && match[1] && match[1].trim()) {
        extracted[field] = match[1].trim();
        break; // Use first successful match
      }
    }
  }

  // Extract itinerary with enhanced parsing
  extracted.itinerary = extractItineraryDataEnhanced(text);

  // Parse and validate numeric fields
  [
    "estimatedCost",
    "perDiem",
    "airRail",
    "lodging",
    "rentalCar",
    "miscellaneous",
  ].forEach((field) => {
    if (extracted[field]) {
      const numericValue = parseFloat(extracted[field].replace(/[,$]/g, ""));
      if (!isNaN(numericValue)) {
        extracted[field] = numericValue;
      }
    }
  });

  // Extract dates
  extracted.departureDate = extractDate(text, ["departure", "depart", "leave"]);
  extracted.returnDate = extractDate(text, ["return", "arrive back", "end"]);

  return extracted;
}

/**
 * Extract dates based on context keywords
 */
function extractDate(text, keywords) {
  for (const keyword of keywords) {
    const pattern = new RegExp(`${keyword}[:\\s]*([\\d\\/\\-]+)`, "i");
    const match = text.match(pattern);
    if (match) {
      return normalizeDate(match[1]);
    }
  }
  return null;
}

/**
 * Enhanced itinerary extraction with table detection
 */
function extractItineraryDataEnhanced(text) {
  const itinerary = [];

  // Look for itinerary section with multiple possible headers
  const itinerarySectionPatterns = [
    /AUTHORIZED OFFICIAL ITINERARY([\s\S]*?)(?=\n\n|\n[A-Z]{3,}|$)/i,
    /ITINERARY([\s\S]*?)(?=\n\n|\n[A-Z]{3,}|$)/i,
    /TRAVEL SCHEDULE([\s\S]*?)(?=\n\n|\n[A-Z]{3,}|$)/i,
  ];

  let itinerarySection = null;
  for (const pattern of itinerarySectionPatterns) {
    const match = text.match(pattern);
    if (match) {
      itinerarySection = match[1];
      break;
    }
  }

  if (!itinerarySection) {
    // Try to find any date/location patterns in the text
    return extractItineraryFromGeneralText(text);
  }

  // Enhanced date patterns
  const datePatterns = [
    /(\d{1,2}\/\d{1,2}\/\d{4})/g,
    /(\d{1,2}-\d{1,2}-\d{4})/g,
    /(\d{4}-\d{1,2}-\d{1,2})/g,
  ];

  // Enhanced location patterns
  const locationPatterns = [
    /([A-Za-z\s]+),\s*([A-Z]{2})\b/g,
    /([A-Za-z\s]+)\s+([A-Z]{2})\s/g,
  ];

  let dates = [];
  let locations = [];

  // Extract dates with multiple patterns
  for (const pattern of datePatterns) {
    const matches = [...itinerarySection.matchAll(pattern)];
    if (matches.length > 0) {
      dates = matches.map((m) => m[1]);
      break;
    }
  }

  // Extract locations with multiple patterns
  for (const pattern of locationPatterns) {
    const matches = [...itinerarySection.matchAll(pattern)];
    if (matches.length > 0) {
      locations = matches.map((m) => ({
        city: m[1].trim(),
        state: m[2].trim(),
      }));
      break;
    }
  }

  // Combine dates and locations
  const maxItems = Math.max(dates.length, locations.length);
  for (let i = 0; i < maxItems; i++) {
    const item = {};

    if (i < dates.length) {
      item.date = normalizeDate(dates[i]);
    }

    if (i < locations.length) {
      item.city = locations[i].city;
      item.state = locations[i].state;
    } else if (locations.length > 0) {
      // Use the last known location if we have more dates than locations
      item.city = locations[locations.length - 1].city;
      item.state = locations[locations.length - 1].state;
    }

    if (item.city || item.date) {
      itinerary.push(item);
    }
  }

  return itinerary;
}

/**
 * Extract itinerary from general text when no formal itinerary section exists
 */
function extractItineraryFromGeneralText(text) {
  const itinerary = [];

  // Look for common travel-related keywords with dates and locations
  const travelPatterns = [
    /(?:travel|trip|visit|go)\s+to\s+([A-Za-z\s]+),\s*([A-Z]{2})/gi,
    /(?:from|depart)\s+([A-Za-z\s]+),\s*([A-Z]{2})/gi,
    /(?:arrive|return)\s+([A-Za-z\s]+),\s*([A-Z]{2})/gi,
  ];

  const locations = new Set();

  for (const pattern of travelPatterns) {
    const matches = [...text.matchAll(pattern)];
    matches.forEach((match) => {
      locations.add(
        JSON.stringify({
          city: match[1].trim(),
          state: match[2].trim(),
        })
      );
    });
  }

  // Convert back to objects and add to itinerary
  Array.from(locations).forEach((loc) => {
    const location = JSON.parse(loc);
    itinerary.push(location);
  });

  return itinerary;
}

/**
 * Normalize date formats
 */
function normalizeDate(dateString) {
  if (!dateString) return null;

  try {
    // Handle different date formats
    let normalized = dateString;

    // Convert MM/DD/YYYY to YYYY-MM-DD
    if (/\d{1,2}\/\d{1,2}\/\d{4}/.test(dateString)) {
      const parts = dateString.split("/");
      normalized = `${parts[2]}-${parts[0].padStart(
        2,
        "0"
      )}-${parts[1].padStart(2, "0")}`;
    }

    // Convert MM-DD-YYYY to YYYY-MM-DD
    if (/\d{1,2}-\d{1,2}-\d{4}/.test(dateString)) {
      const parts = dateString.split("-");
      normalized = `${parts[2]}-${parts[0].padStart(
        2,
        "0"
      )}-${parts[1].padStart(2, "0")}`;
    }

    // Validate the date
    const date = new Date(normalized);
    if (isNaN(date.getTime())) {
      return dateString; // Return original if can't parse
    }

    return normalized;
  } catch (error) {
    Logger.log(`Date normalization error: ${error.message}`);
    return dateString;
  }
}

/**
 * Advanced form field detection using context clues
 */
function detectFormFields(text) {
  const fields = {};
  const lines = text.split("\n");

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Look for GSA form structure patterns
    if (line.match(/\d+\.\s*AUTHORIZATION NUMBER/i)) {
      fields.authorizationNumber = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*TRAVELER/i)) {
      fields.travelerName = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*TITLE/i)) {
      fields.title = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*VENDOR CODE/i)) {
      fields.vendorCode = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*CURRENT RESIDENCE ADDRESS/i)) {
      fields.currentAddress = extractFromNextLines(lines, i + 1, 3);
    }

    if (line.match(/\d+\.\s*OFFICE.*DIVISION/i)) {
      fields.officeDivision = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*OFFICIAL DUTY STATION/i)) {
      fields.dutyStation = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*CONTACT TELEPHONE/i)) {
      fields.contactNumber = extractFromNextLines(lines, i + 1, 1);
    }

    if (line.match(/\d+\.\s*TRAVEL PURPOSE/i)) {
      fields.travelPurpose = extractFromNextLines(lines, i + 1, 2);
    }

    if (line.match(/\d+\.\s*BRIEF DESCRIPTION/i)) {
      fields.briefDescription = extractFromNextLines(lines, i + 1, 3);
    }
  }

  return fields;
}

/**
 * Helper function to extract data from following lines
 */
function extractFromNextLines(lines, startIndex, maxLines) {
  const extracted = [];

  for (
    let i = startIndex;
    i < Math.min(startIndex + maxLines, lines.length);
    i++
  ) {
    const line = lines[i].trim();
    if (line && !line.match(/^\d+\./)) {
      extracted.push(line);
    } else if (line.match(/^\d+\./)) {
      // Stop when we hit the next numbered section
      break;
    }
  }

  return extracted.join(" ").trim();
}

/**
 * Validate extracted data quality
 */
function validateExtractionQuality(extractedData) {
  const quality = {
    score: 0,
    maxScore: 0,
    issues: [],
    confidence: "LOW",
  };

  // Check for required fields
  const requiredFields = ["travelerName", "estimatedCost", "travelPurpose"];
  requiredFields.forEach((field) => {
    quality.maxScore += 20;
    if (extractedData[field]) {
      quality.score += 20;
    } else {
      quality.issues.push(`Missing required field: ${field}`);
    }
  });

  // Check for optional but important fields
  const optionalFields = [
    "authorizationNumber",
    "dutyStation",
    "contactNumber",
    "title",
  ];
  optionalFields.forEach((field) => {
    quality.maxScore += 10;
    if (extractedData[field]) {
      quality.score += 10;
    }
  });

  // Check itinerary quality
  quality.maxScore += 20;
  if (extractedData.itinerary && extractedData.itinerary.length > 0) {
    quality.score += 20;
  } else {
    quality.issues.push("No itinerary data extracted");
  }

  // Check numeric fields
  const numericFields = ["estimatedCost", "perDiem", "airRail"];
  numericFields.forEach((field) => {
    quality.maxScore += 5;
    if (extractedData[field] && typeof extractedData[field] === "number") {
      quality.score += 5;
    }
  });

  // Calculate confidence level
  const percentage = (quality.score / quality.maxScore) * 100;
  if (percentage >= 80) {
    quality.confidence = "HIGH";
  } else if (percentage >= 60) {
    quality.confidence = "MEDIUM";
  } else {
    quality.confidence = "LOW";
  }

  return quality;
}

/**
 * Test function to validate the document processor
 */
function testDocumentProcessor() {
  // This function can be used to test the document processor
  // with sample data or uploaded files
  Logger.log("Document processor test function ready");
}

/**
 * Legacy function for backward compatibility
 */
function extractGSAFormData(text) {
  // Wrapper for the enhanced function to maintain backward compatibility
  return extractGSAFormDataEnhanced(text);
}

// Export processor functions for Node.js testing. These exports are ignored
// in the Google Apps Script environment. Only functions that do not depend
// heavily on Apps Script services should be used in local testing.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    processDocument,
    extractTextFromPDF,
    extractTextWithOCR,
    extractTextFromWord,
    cleanExtractedText,
    extractGSAFormDataEnhanced,
    extractDate,
    extractItineraryDataEnhanced,
    extractItineraryFromGeneralText,
    normalizeDate,
    detectFormFields,
    validateExtractionQuality,
    extractGSAFormData,
  };
}
