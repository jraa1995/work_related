/**
 * document_processor.gs
 *
 * Utilities for processing uploaded documents (PDF and Word) within
 * Google Apps Script.  This file mirrors the functionality of
 * document_processor.js but has been stripped of any Node.js
 * specific code.  All functions rely solely on Apps Script
 * services such as DriveApp, DocumentApp, and Utilities.  The
 * primary entry point is processDocument(), which returns the
 * extracted text and metadata.
 */

/**
 * Process an uploaded document.  Accepts base64 encoded data,
 * MIME type and filename.  Determines whether to use PDF or Word
 * extraction and returns cleaned text, extracted data and quality
 * metrics.
 *
 * @param {string} base64Data The document encoded as base64
 * @param {string} mimeType The MIME type of the document
 * @param {string} filename The original filename
 * @return {Object} Result containing success flag, extractedText,
 *         extractedData, quality metrics and metadata
 */
function processDocument(base64Data, mimeType, filename) {
  try {
    var extractedText = '';
    if (mimeType === 'application/pdf') {
      extractedText = extractTextFromPDF(base64Data);
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || mimeType === 'application/msword') {
      extractedText = extractTextFromWord(base64Data);
    } else {
      throw new Error('Unsupported document type: ' + mimeType);
    }
    if (!extractedText) {
      throw new Error('No text could be extracted from the document');
    }
    var cleaned = cleanExtractedText(extractedText);
    var extractedData = extractGSAFormDataEnhanced(cleaned);
    var quality = validateExtractionQuality(extractedData);
    return {
      success: true,
      extractedText: cleaned,
      extractedData: extractedData,
      quality: quality,
      metadata: {
        filename: filename,
        mimeType: mimeType,
        extractionMethod: mimeType === 'application/pdf' ? 'PDF' : 'Word',
        textLength: cleaned.length,
        confidence: quality.confidence
      }
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
      extractedData: {},
      quality: { confidence: 'FAILED', issues: [err.message] }
    };
  }
}

/**
 * Extract text from a PDF document using the Apps Script Drive
 * conversion service.  Falls back to OCR if direct conversion
 * produces insufficient text.
 *
 * @param {string} base64Data Base64 encoded PDF
 * @return {string|null} Extracted text or null on failure
 */
function extractTextFromPDF(base64Data) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', 'temp_pdf_' + Date.now() + '.pdf');
    var tempFile = DriveApp.createFile(blob);
    try {
      var resource = { title: 'temp_conversion_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS };
      var docFile = Drive.Files.copy(resource, tempFile.getId(), { convert: true });
      var doc = DocumentApp.openById(docFile.id);
      var text = doc.getBody().getText();
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      DriveApp.getFileById(docFile.id).setTrashed(true);
      if (text && text.trim().length > 50) {
        return text;
      }
    } catch (convErr) {
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    }
    // Fallback to OCR
    try {
      return extractTextWithOCR(base64Data);
    } catch (ocrErr) {
      throw new Error('All PDF extraction methods failed');
    }
  } catch (err) {
    return null;
  }
}

/**
 * Perform OCR on a PDF using Drive API conversion with OCR enabled.
 *
 * @param {string} base64Data Base64 encoded PDF
 * @return {string} Extracted text from OCR
 */
function extractTextWithOCR(base64Data) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', 'ocr_temp_' + Date.now() + '.pdf');
  var file = DriveApp.createFile(blob);
  var ocrResource = { title: 'ocr_conversion_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS };
  var ocrFile = Drive.Files.copy(ocrResource, file.getId(), { convert: true, ocr: true, ocrLanguage: 'en' });
  var doc = DocumentApp.openById(ocrFile.id);
  var text = doc.getBody().getText();
  DriveApp.getFileById(file.getId()).setTrashed(true);
  DriveApp.getFileById(ocrFile.id).setTrashed(true);
  return text;
}

/**
 * Extract text from a Word document by converting it to a Google Doc.
 *
 * @param {string} base64Data Base64 encoded Word document
 * @return {string|null} Extracted text or null
 */
function extractTextFromWord(base64Data) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'temp_word_' + Date.now() + '.docx');
    var tempFile = DriveApp.createFile(blob);
    try {
      var resource = { title: 'word_conversion_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS };
      var docFile = Drive.Files.copy(resource, tempFile.getId(), { convert: true });
      var doc = DocumentApp.openById(docFile.id);
      var text = doc.getBody().getText();
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      DriveApp.getFileById(docFile.id).setTrashed(true);
      return text;
    } catch (convErr) {
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      return null;
    }
  } catch (err) {
    return null;
  }
}

/**
 * Clean extracted text by normalising line breaks and whitespace.
 *
 * @param {string} text Raw extracted text
 * @return {string} Cleaned text
 */
function cleanExtractedText(text) {
  return text.replace(/\r/g, '').replace(/\n{2,}/g, '\n').trim();
}

/**
 * Enhanced form data extraction.  Attempts multiple patterns per
 * field to handle variations in document layouts.  Also extracts
 * itinerary data using extractItineraryData() from utils.
 *
 * @param {string} text Cleaned document text
 * @return {Object} Extracted data
 */
function extractGSAFormDataEnhanced(text) {
  if (!text) return {};
  var extracted = {};
  var fieldExtractions = {
    authorizationNumber: [
      /authorization\s+number[:\s]*([^\n\r]+)/i,
      /auth\s*#[:\s]*([^\n\r]+)/i,
      /authorization[:\s]*([A-Z0-9\-\/]+)/i
    ],
    travelerName: [
      /traveler[:\s]*([^\n\r]+)/i,
      /employee\s+name[:\s]*([^\n\r]+)/i,
      /name[:\s]*([A-Za-z\s,\.]+)(?=\s|$)/i
    ],
    title: [
      /title[:\s]*([^\n\r]+)/i,
      /position[:\s]*([^\n\r]+)/i,
      /job\s+title[:\s]*([^\n\r]+)/i
    ],
    vendorCode: [
      /pegasys\s+vendor\s+code[:\s]*([E]\d{8,9})/i,
      /vendor\s+code[:\s]*([E]\d{8,9})/i,
      /employee\s+id[:\s]*([E]\d{8,9})/i
    ],
    currentAddress: [
      /current\s+residence\s+address[:\s]*([^\n\r]+)/i,
      /address[:\s]*([^\n\r]+)/i,
      /residence[:\s]*([^\n\r]+)/i
    ],
    officeDivision: [
      /office\/service\s+and\s+division[:\s]*([^\n\r]+)/i,
      /office[:\s]*([^\n\r]+)/i,
      /division[:\s]*([^\n\r]+)/i
    ],
    dutyStation: [
      /official\s+duty\s+station[:\s]*([^\n\r]+)/i,
      /duty\s+station[:\s]*([^\n\r]+)/i,
      /work\s+location[:\s]*([^\n\r]+)/i
    ],
    contactNumber: [
      /contact\s+telephone\s+number[:\s]*([^\n\r]+)/i,
      /phone[:\s]*([^\n\r]+)/i,
      /telephone[:\s]*([^\n\r]+)/i,
      /(\(?[0-9]{3}\)?[-\.\s]?[0-9]{3}[-\.\s]?[0-9]{4})/
    ],
    travelPurpose: [
      /travel\s+purpose[:\s]*([^\n\r]+)/i,
      /purpose[:\s]*([^\n\r]+)/i,
      /reason\s+for\s+travel[:\s]*([^\n\r]+)/i
    ],
    briefDescription: [
      /brief\s+description[:\s]*([^\n\r]+)/i,
      /description[:\s]*([^\n\r]+)/i,
      /details[:\s]*([^\n\r]+)/i
    ],
    estimatedCost: [
      /estimated\s+cost[:\s]*total[:\s]*\$?([0-9,]+\.?\d*)/i,
      /total\s+cost[:\s]*\$?([0-9,]+\.?\d*)/i,
      /amount[:\s]*\$?([0-9,]+\.?\d*)/i
    ],
    perDiem: [
      /per\s+diem[:\s]*\$?([0-9,]+\.?\d*)/i,
      /meals\s+and\s+incidentals[:\s]*\$?([0-9,]+\.?\d*)/i
    ],
    airRail: [
      /air\/rail[:\s]*\$?([0-9,]+\.?\d*)/i,
      /transportation[:\s]*\$?([0-9,]+\.?\d*)/i,
      /airfare[:\s]*\$?([0-9,]+\.?\d*)/i
    ],
    lodging: [
      /lodging[:\s]*\$?([0-9,]+\.?\d*)/i,
      /hotel[:\s]*\$?([0-9,]+\.?\d*)/i,
      /accommodation[:\s]*\$?([0-9,]+\.?\d*)/i
    ],
    rentalCar: [
      /rental\s+car[:\s]*\$?([0-9,]+\.?\d*)/i,
      /car\s+rental[:\s]*\$?([0-9,]+\.?\d*)/i
    ]
  };
  for (var field in fieldExtractions) {
    var patterns = fieldExtractions[field];
    for (var i = 0; i < patterns.length; i++) {
      var m = text.match(patterns[i]);
      if (m && m[1]) {
        extracted[field] = m[1].trim();
        break;
      }
    }
  }
  // Extract itinerary using utils function extractItineraryData()
  extracted.itinerary = extractItineraryData(text);
  // Convert numeric fields to numbers
  var numFields = ['estimatedCost','perDiem','airRail','lodging','rentalCar'];
  numFields.forEach(function (nf) {
    if (extracted[nf]) {
      extracted[nf] = parseFloat(extracted[nf].replace(/,/g, ''));
    }
  });
  return extracted;
}

/**
 * Validate extraction quality.  Determines a confidence rating and
 * identifies any missing critical fields.
 *
 * @param {Object} data Extracted data
 * @return {Object} Quality metrics with confidence and issues
 */
function validateExtractionQuality(data) {
  var issues = [];
  var required = ['travelerName','authorizationNumber','travelPurpose','estimatedCost'];
  required.forEach(function (field) {
    if (!data[field]) {
      issues.push('Missing ' + field);
    }
  });
  var confidence = 'HIGH';
  if (issues.length > 0 && issues.length <= 2) {
    confidence = 'MEDIUM';
  } else if (issues.length > 2) {
    confidence = 'LOW';
  }
  return { confidence: confidence, issues: issues };
}