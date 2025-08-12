const fs = require('fs');
const path = require('path');

// Import the validation function and configuration from the TAR build.
// The tar_build directory contains Node‑compatible modules that export
// server functions when running outside of Apps Script.  We use
// `validateTarWithPerDiem` to perform the same validations that the
// Google Apps Script version does.
const {
  validateTarWithPerDiem,
} = require('./tar_build/main.js');

/**
 * Parse a single CSV line into an array of values.  The parser
 * respects quoted fields (double quotes) and allows commas inside
 * quoted segments.  It does not support escaped quotes within
 * quoted fields, which are not expected in the provided sample.
 *
 * @param {string} line A line of text from the CSV file
 * @returns {string[]} An array of values corresponding to the columns
 */
function parseCsvLine(line) {
  const values = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      // Toggle quote state.  If the next character is another quote
      // (escaped quote), skip it and append a single quote.
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++; // Skip the escaped quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      // Comma outside quotes ends the current field
      values.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  // Push the last field
  values.push(current);
  return values;
}

/**
 * Read and parse a CSV file.  The function returns an array of
 * objects keyed by the header row.  Empty rows are skipped.
 *
 * @param {string} filename Path to the CSV file
 * @returns {Object[]} An array of record objects
 */
function readCsv(filename) {
  const csvContent = fs.readFileSync(filename, 'utf8');
  const lines = csvContent.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (lines.length === 0) return [];
  const header = parseCsvLine(lines[0]);
  const records = [];
  for (let i = 1; i < lines.length; i++) {
    const row = parseCsvLine(lines[i]);
    const record = {};
    for (let j = 0; j < header.length; j++) {
      const key = header[j];
      record[key] = row[j] !== undefined ? row[j] : '';
    }
    records.push(record);
  }
  return records;
}

/**
 * Aggregate multiple CSV rows belonging to the same TAR ID.  This
 * mirrors the logic implemented in the client for bulk processing:
 *
 * - Earliest departure/date submitted becomes the trip start date.
 * - Latest return date becomes the trip end date.
 * - First non‑empty destination and purpose/comment are used.
 * - Total actual cost is the value from the first non‑zero Total
 *   Actual Cost column; otherwise it is the sum of Cost Type Value
 *   across all rows.
 * - Estimated cost is taken from the first non‑zero Total Travel
 *   Estimate; otherwise it falls back to total actual cost.
 *
 * @param {Object[]} records The array of raw CSV record objects
 * @returns {Object[]} An array of aggregated record objects
 */
function aggregateByTarId(records) {
  const grouped = {};
  records.forEach((rec) => {
    const id = rec['TAR ID'] || '';
    if (!id) return;
    if (!grouped[id]) grouped[id] = [];
    grouped[id].push(rec);
  });
  const aggregated = [];
  for (const id of Object.keys(grouped)) {
    const recs = grouped[id];
    const base = recs[0];
    const agg = Object.assign({}, base);
    // Calculate earliest start and latest end dates
    let startDate = null;
    let endDate = null;
    recs.forEach((r) => {
      const sRaw = r['Departure Date'] || r['Date Submitted'] || '';
      const eRaw = r['Return Date'] || '';
      if (sRaw) {
        const sd = new Date(sRaw);
        if (!isNaN(sd)) {
          if (!startDate || sd < startDate) startDate = sd;
        }
      }
      if (eRaw) {
        const ed = new Date(eRaw);
        if (!isNaN(ed)) {
          if (!endDate || ed > endDate) endDate = ed;
        }
      }
    });
    if (startDate) agg['Departure Date'] = startDate.toISOString().split('T')[0];
    if (endDate) agg['Return Date'] = endDate.toISOString().split('T')[0];
    // Use first non‑empty destination
    const destRec = recs.find((r) => {
      const d = r['Destination (City, State, Country)'];
      return d && d.trim().length > 0;
    });
    if (destRec) agg['Destination (City, State, Country)'] = destRec['Destination (City, State, Country)'];
    // Use first non‑empty travel purpose or comments
    const purposeRec = recs.find((r) => r['Travel Purpose'] && r['Travel Purpose'].trim().length > 0);
    const commentRec = recs.find((r) => r['Comments'] && r['Comments'].trim().length > 0);
    agg['Travel Purpose'] = purposeRec ? purposeRec['Travel Purpose'] : (commentRec ? commentRec['Comments'] : '');
    // Compute total actual cost and estimated cost
    const baseActual = parseFloat(base['Total Actual Cost']) || 0;
    const sumCost = recs.reduce((sum, r) => {
      const val = parseFloat(r['Cost Type Value']) || 0;
      return sum + val;
    }, 0);
    const totalActual = baseActual || sumCost;
    agg['Total Actual Cost'] = totalActual ? totalActual.toString() : '0';
    const estRec = recs.find((r) => parseFloat(r['Total Travel Estimate']) > 0);
    agg['Total Travel Estimate'] = estRec ? estRec['Total Travel Estimate'] : (base['Total Travel Estimate'] || '');
    aggregated.push(agg);
  }
  return aggregated;
}

/**
 * Convert an aggregated record into the shape expected by
 * `validateTarWithPerDiem`.  This mirrors the mapRecordToTarData
 * function in the client: parse destination into city/state,
 * derive start/end dates and durations, extract numeric cost values,
 * and prepare metadata fields.
 *
 * @param {Object} record An aggregated CSV record
 * @returns {Object} An object ready to be validated
 */
function mapAggregatedRecord(record) {
  const dest = record['Destination (City, State, Country)'] || '';
  const destParts = dest.split(',');
  const city = destParts[0] ? destParts[0].trim() : '';
  let state = '';
  if (destParts.length > 1) {
    state = destParts[1].trim();
    if (state.length > 2) state = state.substring(0, 2);
    state = state.toUpperCase();
  }
  const parseDate = (d) => {
    if (!d) return '';
    const date = new Date(d);
    if (isNaN(date)) return '';
    return date.toISOString().substring(0, 10);
  };
  const startDate = parseDate(record['Departure Date']);
  const endDate = parseDate(record['Return Date']);
  let duration = 1;
  if (startDate && endDate) {
    const sd = new Date(startDate);
    const ed = new Date(endDate);
    if (ed >= sd) {
      duration = Math.ceil((ed - sd) / (1000 * 60 * 60 * 24)) + 1;
    }
  }
  const toNumber = (val) => {
    const num = parseFloat(String(val).replace(/[^0-9.]/g, ''));
    return isNaN(num) ? 0 : num;
  };
  let totalCost = toNumber(record['Total Actual Cost']);
  if (!totalCost) totalCost = toNumber(record['Total Travel Estimate']);
  const estimatedCost = toNumber(record['Total Travel Estimate']) || totalCost;
  return {
    traveler: record['Traveler'] || record['Creator'] || '',
    city: city,
    state: state,
    tripStartDate: startDate,
    tripEndDate: endDate,
    duration: duration,
    totalCost: totalCost,
    estimatedCost: estimatedCost,
    purpose: record['Travel Purpose'] || record['Comments'] || '',
    // Use the Creator field as the point of contact.  Since
    // validateFormData requires a numeric phone number for
    // contactNumber, we also provide a dummy contactNumber that
    // satisfies the validation pattern.  In a real system this
    // should be replaced with the traveller's phone number or a
    // valid POC.
    poc: record['Creator'] || '',
    contactNumber: '1234567890',
    dutyStation: city && state ? `${city}, ${state}` : '',
    tarId: record['TAR ID'] || '',
    tarTitle: record['TAR Title'] || '',
    contractId: record['Contract Specific ID'] || '',
  };
}

/**
 * Process a CSV file and validate each TAR ID using the TAR
 * validation logic.  Results are printed to the console in a
 * readable format.
 *
 * @param {string} csvPath The path to the CSV file to process
 */
function validateCsvFile(csvPath) {
  const records = readCsv(csvPath);
  if (records.length === 0) {
    console.error('No records found in CSV file');
    return;
  }
  const aggregated = aggregateByTarId(records);
  console.log(`Processing ${aggregated.length} unique TAR IDs...`);
  aggregated.forEach((agg) => {
    const data = mapAggregatedRecord(agg);
    const result = validateTarWithPerDiem(data);
    const status = result && result.success && result.isValid ? '✅' : '❌';
    console.log(`\nTAR ID: ${data.tarId} ${status}`);
    if (result && result.validationReport) {
      const report = result.validationReport;
      const messages = [];
      if (report.errors && report.errors.length > 0) messages.push(...report.errors);
      if (report.warnings && report.warnings.length > 0) messages.push(...report.warnings);
      if (report.validation && report.validation.messages) messages.push(...report.validation.messages);
      if (messages.length === 0) {
        // If the report is invalid but has no specific messages, use the
        // generic message from the result object.  This usually
        // indicates a variance/deviation failure.
        if (result && result.message) {
          console.log(result.message);
        } else {
          console.log('No errors or warnings');
        }
      } else {
        messages.forEach((m) => console.log('- ' + m));
      }
    } else if (result && result.errors && result.errors.length > 0) {
      // The validation failed before generating a report; output the errors directly
      result.errors.forEach((err) => console.log('- ' + err));
    } else {
      console.log('Validation failed or returned no report');
    }
  });
}

// If this script is run directly (e.g. `node bulk_test.js`), process
// the CSV file passed as a command line argument.  If no argument
// supplied, default to the sample file in the project root.
if (require.main === module) {
  const args = process.argv.slice(2);
  const csvPath = args[0] || path.join(__dirname, 'TAR_Bulk_Test - Data.csv');
  if (!fs.existsSync(csvPath)) {
    console.error(`CSV file not found: ${csvPath}`);
    process.exit(1);
  }
  validateCsvFile(csvPath);
}