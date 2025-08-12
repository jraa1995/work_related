// client.js
// Front-end logic for the Enhanced TAR Validation System

/*
  This file contains all of the client-side JavaScript required to power
  the TAR validation interface. It replaces a very large inline script
  that previously lived in index.html. The code is organized into
  logical sections: initialization, funding data lookups, form helpers,
  per diem fetching, expense validations, total calculation, summary
  updates, submission handling, result rendering, and analyst actions.

  By moving this logic into a separate module we keep the HTML clean
  and make it much easier to maintain and test the front‑end code.
*/

// Global state used throughout the application
let fundingData = [];
let selectedTaskOrder = null;
let expenseValidations = {};
let perDiemRates = null;
let currentValidation = null;
let currentFile = null;
let currentBulkFile = null; // holds the CSV file selected for bulk processing

// Initialise the page once the DOM is fully loaded
document.addEventListener('DOMContentLoaded', () => {
  const loadingIcon = document.getElementById('loading-icon');
  const progressBar = document.getElementById('progress-bar');
  if (loadingIcon) loadingIcon.style.display = 'none';
  if (progressBar) progressBar.classList.remove('active');
  setupEventListeners();
  loadFundingData();
});

/**
 * Handle selection of a file from the hidden file input. This function exists
 * to support the inline onchange="handleFileSelect(event)" attribute on the
 * file input in index.html. It mirrors the behaviour of the internal
 * handleSelectedFile() helper used in setupDragAndDrop(). When a user picks a
 * file via the file picker, this function updates the global currentFile
 * reference and displays the file information panel.
 *
 * @param {Event} event The change event from the file input
 */
function handleFileSelect(event) {
  const file = event && event.target && event.target.files ? event.target.files[0] : null;
  if (!file) return;
  currentFile = file;
  const fileInfo = document.getElementById('file-info');
  const fileNameSpan = document.getElementById('file-name');
  const fileSizeSpan = document.getElementById('file-size');
  if (fileInfo && fileNameSpan && fileSizeSpan) {
    fileInfo.classList.remove('hidden');
    fileNameSpan.textContent = file.name;
    // Display size in KB with one decimal place
    const sizeInKB = (file.size / 1024).toFixed(1);
    fileSizeSpan.textContent = `${sizeInKB} KB`;
  }
}

/**
 * Handle selection of a CSV file for bulk processing. When the user picks a
 * file via the file picker in the bulk upload section this function
 * validates the extension, stores the reference in currentBulkFile, displays
 * the file name/size and enables the "Process Bulk" button.
 *
 * @param {Event} event The change event from the CSV file input
 */
function handleBulkFileSelect(event) {
  const file = event && event.target && event.target.files ? event.target.files[0] : null;
  if (!file) return;
  const fileNameLower = file.name.toLowerCase();
  if (!fileNameLower.endsWith('.csv')) {
    showError('Please select a valid CSV file');
    return;
  }
  currentBulkFile = file;
  const infoDiv = document.getElementById('bulk-file-info');
  const nameSpan = document.getElementById('bulk-file-name');
  const sizeSpan = document.getElementById('bulk-file-size');
  if (infoDiv && nameSpan && sizeSpan) {
    nameSpan.textContent = file.name;
    const sizeKB = (file.size / 1024).toFixed(1);
    sizeSpan.textContent = `(${sizeKB} KB)`;
    infoDiv.classList.remove('hidden');
  }
  const btn = document.getElementById('process-bulk-btn');
  if (btn) btn.disabled = false;
  showSuccess('✅ CSV file selected. Ready to process.');
}

/**
 * Read the selected CSV file and begin bulk validation. This function uses
 * FileReader to load the file contents, parses the CSV into an array of
 * records, then processes each record asynchronously. Progress is updated
 * for user feedback and results are displayed in a collapsible list.
 */
function processBulkCsv() {
  if (!currentBulkFile) {
    showError('Please select a CSV file first');
    return;
  }
  updateProgress(10, 'Reading CSV data...');
  const reader = new FileReader();
  reader.onload = function(e) {
    const text = e.target && e.target.result ? e.target.result : '';
    const records = parseCsv(text);
    if (!records || records.length === 0) {
      showError('No valid data found in CSV file');
      updateProgress(0, 'Ready to validate...');
      return;
    }
    processBulkRecords(records);
  };
  reader.onerror = function() {
    showError('Failed to read CSV file');
    updateProgress(0, 'Ready to validate...');
  };
  reader.readAsText(currentBulkFile);
}

/**
 * Split a single CSV line into an array of values. This helper respects
 * quoted values that may contain commas. Leading/trailing quotes are
 * removed and whitespace is trimmed.
 *
 * @param {string} line A single line from a CSV file
 * @returns {string[]} Array of parsed fields
 */
function splitCsvLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current);
  return result.map((s) => s.trim());
}

/**
 * Parse a CSV string into an array of objects keyed by the header row. Empty
 * lines are ignored. If a row has a different number of fields than the
 * header, it is skipped. This parser is intentionally simple and assumes
 * double quotes are used only to wrap fields containing commas.
 *
 * @param {string} text Raw CSV file contents
 * @returns {Object[]} Array of records where keys are header names
 */
function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter((l) => l && l.trim().length > 0);
  if (lines.length < 2) return [];
  const headers = splitCsvLine(lines[0]);
  const records = [];
  for (let i = 1; i < lines.length; i++) {
    const values = splitCsvLine(lines[i]);
    if (values.length !== headers.length) continue;
    const record = {};
    headers.forEach((h, idx) => {
      record[h] = values[idx];
    });
    records.push(record);
  }
  return records;
}

/**
 * Convert a CSV record into a format expected by the validation engine. It
 * extracts key fields such as traveler name, city/state, trip dates and
 * costs. Missing or malformed data is tolerated where possible. Dates are
 * normalised to YYYY-MM-DD ISO format.
 *
 * @param {Object} record The raw CSV row keyed by column names
 * @returns {Object} Normalised TAR data suitable for validateTarWithPerDiem
 */
function mapRecordToTarData(record) {
  const dest = record['Destination (City, State, Country)'] || '';
  const destParts = dest.split(',');
  const city = destParts[0] ? destParts[0].trim() : '';
  let state = '';
  if (destParts.length > 1) {
    state = destParts[1].trim();
    if (state.length > 2) state = state.substring(0, 2);
    state = state.toUpperCase();
  }
  const startRaw = record['Departure Date'] || record['Date Submitted'] || '';
  const endRaw = record['Return Date'] || '';
  const parseDate = (d) => {
    if (!d) return '';
    const date = new Date(d);
    if (isNaN(date)) return '';
    const iso = date.toISOString();
    return iso.substring(0, 10);
  };
  const startDate = parseDate(startRaw);
  const endDate = parseDate(endRaw);
  let duration = 1;
  if (startDate && endDate) {
    try {
      const start = new Date(startDate);
      const end = new Date(endDate);
      if (end >= start) {
        duration = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
      }
    } catch (e) {}
  }
  const toNumber = (val) => {
    const num = parseFloat(String(val).replace(/[^0-9.]/g, ''));
    return isNaN(num) ? 0 : num;
  };
  let totalCost = toNumber(record['Total Actual Cost']);
  if (!totalCost) {
    totalCost = toNumber(record['Total Travel Estimate']);
  }
  const costTypeValue = toNumber(record['Cost Type Value']);
  if (!totalCost && costTypeValue) {
    totalCost = costTypeValue;
  }
  let estimatedCost = toNumber(record['Total Travel Estimate']);
  if (!estimatedCost) estimatedCost = totalCost;
  const data = {
    traveler: record['Traveler'] || record['Creator'] || '',
    city: city,
    state: state,
    tripStartDate: startDate,
    tripEndDate: endDate,
    duration: duration,
    totalCost: totalCost,
    estimatedCost: estimatedCost,
    purpose: record['Travel Purpose'] || record['Comments'] || '',
    poc: record['Creator'] || '',
    dutyStation: city && state ? `${city}, ${state}` : '',
    tarId: record['TAR ID'] || '',
    tarTitle: record['TAR Title'] || '',
    contractId: record['Contract Specific ID'] || '',
    costCategory: record['Cost Category'] || '',
    costType: record['Cost Type'] || '',
  };
  return data;
}

/**
 * Add a single result row to the bulk results display. Each row shows the
 * TAR ID and a summary of success or errors/warnings. It formats the
 * validation report into a user-friendly string.
 *
 * @param {string} tarId The TAR ID associated with this result
 * @param {Object} res The result returned by validateTarWithPerDiem
 */
function addBulkResultRow(tarId, res) {
  const container = document.getElementById('bulk-results-content');
  if (!container) return;
  const row = document.createElement('div');
  row.className = 'p-3 border rounded-lg';
  let statusIcon = '✅';
  let details = [];
  // Determine status: show ❌ if the result is invalid or unsuccessful
  if (!res || res.success === false || res.isValid === false) {
    statusIcon = '❌';
  }
  if (res && res.validationReport) {
    const vr = res.validationReport;
    if (vr.errors && vr.errors.length) {
      vr.errors.forEach((err) => details.push(`<span class="text-red-600">${err}</span>`));
    }
    if (vr.warnings && vr.warnings.length) {
      vr.warnings.forEach((warn) => details.push(`<span class="text-yellow-600">${warn}</span>`));
    }
    if (vr.messages && vr.messages.length) {
      vr.messages.forEach((msg) => details.push(`<span class="text-green-600">${msg}</span>`));
    }
  } else if (res && res.errors && res.errors.length) {
    res.errors.forEach((err) => details.push(`<span class="text-red-600">${err}</span>`));
  } else if (!res) {
    details.push('No result returned');
  }
  row.innerHTML = `<div class="font-medium mb-1">TAR ID: ${tarId} ${statusIcon}</div><div class="text-sm space-y-1">${details.join('<br>')}</div>`;
  container.appendChild(row);
}

/**
 * Process an array of CSV records by mapping each to a TAR data object and
 * invoking the validation engine (server-side via google.script.run if
 * available, or a simple offline fallback). Progress updates are shown
 * after each record. Once all records are processed the results container
 * is revealed.
 *
 * @param {Object[]} records Array of parsed CSV objects
 */
async function processBulkRecords(records) {
  // Consolidate multiple CSV rows by unique TAR ID so that each TAR is
  // validated once. Each TAR ID may have multiple rows representing
  // different cost lines; we aggregate start/end dates, destination,
  // purpose/comments, and sums of cost values. See mapRecordToTarData
  // for how the resulting object is interpreted.

  // Build a grouping dictionary keyed by TAR ID
  const grouped = {};
  records.forEach((rec) => {
    const id = rec['TAR ID'] || '';
    if (!id) return;
    if (!grouped[id]) grouped[id] = [];
    grouped[id].push(rec);
  });
  // Convert groups into aggregated records
  const aggregated = Object.keys(grouped).map((id) => {
    const recs = grouped[id];
    const base = recs[0];
    const agg = Object.assign({}, base);
    // Earliest departure/date submitted and latest return
    let startDate = null;
    let endDate = null;
    recs.forEach((r) => {
      const s = r['Departure Date'] || r['Date Submitted'] || '';
      const e = r['Return Date'] || '';
      if (s) {
        const sd = new Date(s);
        if (!isNaN(sd)) {
          if (!startDate || sd < startDate) startDate = sd;
        }
      }
      if (e) {
        const ed = new Date(e);
        if (!isNaN(ed)) {
          if (!endDate || ed > endDate) endDate = ed;
        }
      }
    });
    if (startDate) agg['Departure Date'] = startDate.toISOString().split('T')[0];
    if (endDate) agg['Return Date'] = endDate.toISOString().split('T')[0];
    // Use first non-empty destination and purpose/comment
    const destRec = recs.find((r) => r['Destination (City, State, Country)'] && r['Destination (City, State, Country)'].trim().length > 0);
    if (destRec) agg['Destination (City, State, Country)'] = destRec['Destination (City, State, Country)'];
    const purposeRec = recs.find((r) => r['Travel Purpose'] && r['Travel Purpose'].trim().length > 0);
    const commentRec = recs.find((r) => r['Comments'] && r['Comments'].trim().length > 0);
    agg['Travel Purpose'] = purposeRec ? purposeRec['Travel Purpose'] : (commentRec ? commentRec['Comments'] : '');
    // Aggregate cost values: sum Cost Type Value across all rows if Total Actual Cost missing
    const baseTotal = parseFloat(base['Total Actual Cost']) || 0;
    const sumCost = recs.reduce((sum, r) => {
      const val = parseFloat(r['Cost Type Value']) || 0;
      return sum + val;
    }, 0);
    const totalActual = baseTotal || sumCost;
    agg['Total Actual Cost'] = totalActual ? totalActual.toString() : '0';
    // Estimated cost: use first non-zero Total Travel Estimate
    const estRec = recs.find((r) => parseFloat(r['Total Travel Estimate']) > 0);
    agg['Total Travel Estimate'] = estRec ? estRec['Total Travel Estimate'] : (base['Total Travel Estimate'] || '');
    return agg;
  });

  // Prepare result display
  const resultsSection = document.getElementById('bulk-results');
  const container = document.getElementById('bulk-results-content');
  if (container) container.innerHTML = '';
  if (resultsSection) resultsSection.classList.remove('hidden');
  const total = aggregated.length;
  let processed = 0;
  updateProgress(20, `Processing ${total} record${total === 1 ? '' : 's'}...`);
  // Process each aggregated TAR
  for (const agg of aggregated) {
    const tarId = agg['TAR ID'] || '';
    const data = mapRecordToTarData(agg);
    if (typeof google !== 'undefined' && google.script && google.script.run) {
      await new Promise((resolve) => {
        google.script.run
          .withSuccessHandler((res) => {
            addBulkResultRow(tarId, res);
            processed += 1;
            updateProgress(20 + (processed / total) * 80, `Processed ${processed}/${total} record${total === 1 ? '' : 's'}...`);
            resolve();
          })
          .withFailureHandler((err) => {
            addBulkResultRow(tarId, { success: false, errors: [err && err.message ? err.message : 'Unknown error'] });
            processed += 1;
            updateProgress(20 + (processed / total) * 80, `Processed ${processed}/${total} record${total === 1 ? '' : 's'}...`);
            resolve();
          })
          .validateTarWithPerDiem(data);
      });
    } else {
      // When offline (e.g. running locally without Apps Script), perform a
      // simplified validation directly in the client.  This uses default
      // per diem rates defined in CONFIG to calculate an expected
      // cost and compares it to the claimed total.  The result
      // mirrors the structure returned by the server function.
      const offlineResult = offlineValidateTar(data);
      addBulkResultRow(tarId, offlineResult);
      processed += 1;
      updateProgress(20 + (processed / total) * 80, `Processed ${processed}/${total} record${total === 1 ? '' : 's'}...`);
    }
  }
  updateProgress(0, 'Ready to validate...');
  showSuccess('✅ Bulk processing complete');
}

/**
 * Register all event handlers needed by the UI.
 */
function setupEventListeners() {
  // Auto-uppercase state abbreviation
  const stateInput = document.querySelector('input[name="state"]');
  if (stateInput) {
    stateInput.addEventListener('input', (e) => {
      e.target.value = e.target.value.toUpperCase();
    });
  }

  // Calculate trip duration whenever dates change
  const startInput = document.querySelector('input[name="tripStartDate"]');
  const endInput = document.querySelector('input[name="tripEndDate"]');
  if (startInput && endInput) {
    startInput.addEventListener('change', calculateDuration);
    endInput.addEventListener('change', calculateDuration);
  }

  // Set up the task order autocomplete
  setupTaskOrderAutocomplete();

  // Enable drag and drop for file uploads
  setupDragAndDrop();

  // Expense validation hooks
  const carInput = document.querySelector('input[name="carRental"]');
  const parkInput = document.querySelector('input[name="parking"]');
  // Use the updated input name for conference fees to match index.html
  const conferenceInput = document.querySelector('input[name="conferenceFee"]');
  const miscInput = document.querySelector('input[name="miscellaneous"]');
  if (carInput) carInput.addEventListener('input', () => validateCarRental(carInput));
  if (parkInput) parkInput.addEventListener('input', () => validateParking(parkInput));
  if (conferenceInput) conferenceInput.addEventListener('input', () => validateConferenceFee(conferenceInput));
  if (miscInput) miscInput.addEventListener('input', () => validateMiscellaneous(miscInput));
}

/**
 * Update the progress bar and text indicator. A non‑zero percentage
 * activates the animated background to provide feedback during long
 * running operations such as document extraction or API calls.
 * @param {number} percent The percentage complete (0–100)
 * @param {string} text A short message to accompany the progress bar
 */
function updateProgress(percent, text) {
  const bar = document.getElementById('progress-bar');
  const label = document.getElementById('progress-text');
  if (bar) {
    bar.style.width = `${percent}%`;
    if (percent > 0 && percent < 100) {
      bar.classList.add('active');
    } else {
      bar.classList.remove('active');
    }
  }
  if (label) {
    label.textContent = text || '';
  }
}

// Notification helpers. These functions simply wrap the progress
// indicator to provide simple success/warning/error messages. For a
// more sophisticated UI you could replace these with modal toasts or
// alerts, but keeping everything tied to the progress bar means the
// user always knows what the system is doing.
function showSuccess(msg) {
  updateProgress(100, msg);
  setTimeout(() => updateProgress(0, 'Ready to validate...'), 2000);
}

function showError(msg) {
  updateProgress(0, msg);
}

function showWarning(msg) {
  updateProgress(0, msg);
}

/**
 * Retrieve funding data from the backend. In a real Apps Script
 * deployment this calls a server function via google.script.run.
 * When testing locally or if the backend call fails a fallback
 * dataset is used instead. Successful calls update the global
 * fundingData array and display a success message.
 */
function loadFundingData() {
  updateProgress(10, 'Loading funding data...');
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler((res) => {
        if (res && res.success) {
          fundingData = res.data || [];
          showSuccess(`Loaded ${fundingData.length} funding records`);
        } else {
          showError(res && res.message ? res.message : 'Failed to load funding data');
          loadFallbackFundingData();
        }
      })
      .withFailureHandler((err) => {
        console.error('Funding data error:', err);
        showError('Failed to load funding data');
        loadFallbackFundingData();
      })
      .getFundingData();
  } else {
    // Fallback for local testing
    loadFallbackFundingData();
  }
}

/**
 * Provide a small set of funding entries for offline/demo use. This
 * should closely match the structure returned from the backend.
 */
function loadFallbackFundingData() {
  fundingData = [
    { awardPiid: 'TO-2025-DTI-001', pm: 'Sarah Johnson', funding: 125000.0 },
    { awardPiid: 'TO-2025-DTI-002', pm: 'Michael Chen', funding: 87500.0 },
    { awardPiid: 'TO-2025-CLD-003', pm: 'Jennifer Davis', funding: 210000.0 },
    { awardPiid: 'TO-2025-SEC-004', pm: 'Robert Wilson', funding: 156000.0 },
    { awardPiid: 'TO-2025-NET-005', pm: 'Lisa Anderson', funding: 98750.0 },
    { awardPiid: 'TO-2025-APP-006', pm: 'David Brown', funding: 175000.0 },
    { awardPiid: 'TO-2025-INF-007', pm: 'Maria Garcia', funding: 134500.0 },
    { awardPiid: 'TO-2025-SUP-008', pm: 'James Miller', funding: 67200.0 },
  ];
  showSuccess('Demo funding data loaded');
}

/**
 * Set up the task order autocomplete. As the user types into the
 * task order field this will show matching PIIDs or project managers
 * from the fundingData array. Selecting a suggestion populates the
 * input and updates the funding display.
 */
function setupTaskOrderAutocomplete() {
  const input = document.getElementById('taskOrderInput');
  const suggestions = document.getElementById('taskOrderSuggestions');
  if (!input || !suggestions) return;
  input.addEventListener('input', (e) => {
    const query = e.target.value.trim().toLowerCase();
    if (query.length < 2) {
      hideSuggestions();
      return;
    }
    const matches = fundingData.filter(
      (item) => item.awardPiid.toLowerCase().includes(query) || item.pm.toLowerCase().includes(query)
    );
    showSuggestions(matches, query);
  });
  input.addEventListener('blur', () => {
    setTimeout(hideSuggestions, 200);
  });
  input.addEventListener('focus', () => {
    const query = input.value.trim().toLowerCase();
    if (query.length >= 2) {
      const matches = fundingData.filter(
        (item) => item.awardPiid.toLowerCase().includes(query) || item.pm.toLowerCase().includes(query)
      );
      showSuggestions(matches, query);
    }
  });
}

/**
 * Populate the suggestions drop‑down for task orders.
 * @param {Array} matches Array of funding items to display
 * @param {string} query The current search query to highlight
 */
function showSuggestions(matches, query) {
  const suggestions = document.getElementById('taskOrderSuggestions');
  if (!suggestions) return;
  if (matches.length === 0) {
    suggestions.innerHTML = '<div class="p-3 text-sm text-gray-500">No matching task orders found</div>';
    suggestions.classList.remove('hidden');
    return;
  }
  const html = matches
    .map((item) => {
      const highlightPiid = highlightMatch(item.awardPiid, query);
      const highlightPm = highlightMatch(item.pm, query);
      return `<div class="suggestion-item p-3 hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-b-0" onclick="selectTaskOrder('${item.awardPiid}', '${item.pm}', ${item.funding})">
      <div class="flex justify-between items-start">
        <div>
          <div class="font-medium text-gray-900">${highlightPiid}</div>
          <div class="text-sm text-gray-600">PM: ${highlightPm}</div>
        </div>
        <div class="text-right">
          <div class="text-sm font-semibold text-green-600">${item.funding.toLocaleString()}</div>
          <div class="text-xs text-gray-500">Available</div>
        </div>
      </div>
    </div>`;
    })
    .join('');
  suggestions.innerHTML = html;
  suggestions.classList.remove('hidden');
}

/**
 * Highlight the query portion within a given string for the suggestions list.
 * @param {string} text The full text
 * @param {string} query The search term
 * @returns {string} The HTML with highlighted query
 */
function highlightMatch(text, query) {
  if (!query) return text;
  const regex = new RegExp(`(${query})`, 'gi');
  return text.replace(regex, '<span class="bg-yellow-200">$1</span>');
}

/**
 * Hide the suggestions list. Called when the input loses focus or the
 * user selects a suggestion.
 */
function hideSuggestions() {
  const suggestions = document.getElementById('taskOrderSuggestions');
  if (suggestions) suggestions.classList.add('hidden');
}

/**
 * Handle selection of a task order from the suggestions. Updates
 * selectedTaskOrder, populates form fields, displays available
 * funding and triggers funding validation.
 * @param {string} awardPiid The PIID of the selected task order
 * @param {string} pm The project manager
 * @param {number} funding The available funding
 */
function selectTaskOrder(awardPiid, pm, funding) {
  selectedTaskOrder = { awardPiid, pm, funding };
  const orderInput = document.getElementById('taskOrderInput');
  const authorityInput = document.getElementById('authorityInput');
  if (orderInput) orderInput.value = awardPiid;
  if (authorityInput) authorityInput.value = pm;
  updateFundingDisplay(funding);
  hideSuggestions();
  // Validate funding if total cost is already entered
  const totalCostInput = document.querySelector('[name="totalCost"]');
  if (totalCostInput && totalCostInput.value) {
    validateFundingAvailability(parseFloat(totalCostInput.value));
  }
  showSuccess(`Task Order ${awardPiid} selected. PM: ${pm}`);
}

/**
 * Update the funding display panel with the selected task order's
 * available funding and project manager.
 * @param {number} funding The available funding amount
 */
function updateFundingDisplay(funding) {
  const info = document.getElementById('fundingInfo');
  const avail = document.getElementById('availableFunding');
  const pmSpan = document.getElementById('projectManager');
  if (info && avail && pmSpan) {
    avail.textContent = funding.toLocaleString();
    pmSpan.textContent = selectedTaskOrder.pm;
    info.classList.remove('hidden');
  }
}

/**
 * Check whether the requested trip cost exceeds the available funding
 * on the selected task order and update the UI accordingly. Sets
 * entries in expenseValidations for funding so they appear in the
 * summary.
 * @param {number} requestedAmount The claimed total trip cost
 */
function validateFundingAvailability(requestedAmount) {
  if (!selectedTaskOrder) return;
  const available = selectedTaskOrder.funding;
  const warningDiv = document.getElementById('fundingWarning');
  const fundingInfo = document.getElementById('fundingInfo');
  if (requestedAmount > available) {
    if (warningDiv) {
      warningDiv.innerHTML = `<div class="flex items-center text-red-600">
        <svg class="w-4 h-4 mr-1" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clip-rule="evenodd"></path>
        </svg>
        <span class="font-medium">⚠️ FUNDING EXCEEDED: Requested ${requestedAmount.toLocaleString()} exceeds available ${available.toLocaleString()}</span>
      </div>`;
      warningDiv.classList.remove('hidden');
    }
    if (fundingInfo) {
      fundingInfo.className = 'mt-2 p-3 bg-red-50 border border-red-200 rounded-lg';
    }
    expenseValidations.funding = {
      isValid: false,
      level: 'error',
      message: `Trip cost exceeds available funding by ${(requestedAmount - available).toLocaleString()}`,
    };
  } else if (requestedAmount > available * 0.8) {
    const remaining = available - requestedAmount;
    if (warningDiv) {
      warningDiv.innerHTML = `<div class="flex items-center text-yellow-600">
        <svg class="w-4 h-4 mr-1" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-8 2a1 1 0 100 2 1 1 0 000-2zm0-6a1 1 0 00-1 1v2a1 1 0 102 0V7a1 1 0 00-1-1z" clip-rule="evenodd"></path>
        </svg>
        <span class="font-medium">⚠️ High funding usage: ${requestedAmount.toLocaleString()} (remaining: ${remaining.toLocaleString()})</span>
      </div>`;
      warningDiv.classList.remove('hidden');
    }
    if (fundingInfo) {
      fundingInfo.className = 'mt-2 p-3 bg-yellow-50 border border-yellow-200 rounded-lg';
    }
    expenseValidations.funding = {
      isValid: true,
      level: 'warning',
      message: `High funding usage: ${requestedAmount.toLocaleString()} requested, ${remaining.toLocaleString()} remaining`,
    };
  } else {
    if (warningDiv) warningDiv.classList.add('hidden');
    if (fundingInfo) {
      fundingInfo.className = 'mt-2 p-3 bg-blue-50 border border-blue-200 rounded-lg';
    }
    expenseValidations.funding = {
      isValid: true,
      level: 'success',
      message: 'Funding within available limits',
    };
  }
  updateValidationSummary();
}

/**
 * Initialise drag and drop functionality for document upload. This
 * covers both drag events on the drop zone and the hidden file input.
 */
function setupDragAndDrop() {
  const dropZone = document.getElementById('file-drop-zone');
  const fileInput = document.getElementById('tarFile');
  const fileInfo = document.getElementById('file-info');
  const fileNameSpan = document.getElementById('file-name');
  const fileSizeSpan = document.getElementById('file-size');
  if (!dropZone || !fileInput) return;
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
  });
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) {
      fileInput.files = e.dataTransfer.files;
      handleSelectedFile(file);
    }
  });
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleSelectedFile(file);
  });
  function handleSelectedFile(file) {
    currentFile = file;
    if (fileInfo && fileNameSpan && fileSizeSpan) {
      fileInfo.classList.remove('hidden');
      fileNameSpan.textContent = file.name;
      const sizeInKB = (file.size / 1024).toFixed(1);
      fileSizeSpan.textContent = `${sizeInKB} KB`;
    }
  }
}

/**
 * Trigger document extraction. If running inside Apps Script it will
 * call a backend function named extractAndPopulateFields; otherwise
 * a simulated extraction will populate some example fields.
 */
function extractDataFromPdf() {
  if (!currentFile) {
    showError('No file selected');
    return;
  }
  updateProgress(30, 'Extracting data from file...');
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    const reader = new FileReader();
    reader.onload = function (e) {
      const content = e.target.result;
      const base64 = content.split(',')[1];
      google.script.run
        .withSuccessHandler((res) => {
          if (res && res.success) {
            populateFormWithExtractedData(res.data || {});
            showSuccess('Data extracted successfully');
          } else {
            showError(res && res.message ? res.message : 'Extraction failed');
          }
        })
        .withFailureHandler((err) => {
          console.error('Extraction error:', err);
          showError('Extraction failed');
        })
        .extractAndPopulateFields({ fileName: currentFile.name, content: base64 });
    };
    reader.readAsDataURL(currentFile);
  } else {
    // Fallback: simulate extraction with dummy data
    setTimeout(() => {
      const data = {
        traveler: 'John Smith',
        city: 'New York',
        state: 'NY',
        tripStartDate: '2025-05-01',
        tripEndDate: '2025-05-05',
        duration: 5,
        lodgingCost: '205.00',
        mealsCost: '79.00',
        perDiemTotal: '1420.00',
        totalCost: '3000.00',
        taskOrder: 'TO-2025-DTI-001',
      };
      populateFormWithExtractedData(data);
      showSuccess('Sample data extracted');
    }, 2000);
  }
}

/**
 * Populate form fields based on extracted data. Fields that exist in
 * the form will be updated directly. Task order selection is handled
 * specially to update the funding display and validations.
 * @param {Object} data Key/value pairs extracted from the document
 */
function populateFormWithExtractedData(data) {
  Object.keys(data).forEach((key) => {
    const field = document.querySelector(`[name="${key}"]`);
    if (field) {
      field.value = data[key];
      // Trigger validations where appropriate
      if (field.name === 'carRental') validateCarRental(field);
      if (field.name === 'parking') validateParking(field);
      // Use the correct input name for conference/training fees as defined in the form
      if (field.name === 'conferenceFee') validateConferenceFee(field);
      if (field.name === 'miscellaneous') validateMiscellaneous(field);
    }
  });
  // Handle task order selection separately
  if (data.taskOrder) {
    const match = fundingData.find((item) => item.awardPiid === data.taskOrder);
    if (match) {
      selectTaskOrder(match.awardPiid, match.pm, match.funding);
    } else {
      document.getElementById('taskOrderInput').value = data.taskOrder;
      showError(`Task order ${data.taskOrder} not found in funding data`);
    }
  }
  calculateDuration();
  setTimeout(() => {
    fetchPerDiemRates();
    validateAndCalculate();
  }, 300);
}

/**
 * Calculate the number of days between the start and end date. The
 * duration includes both start and end dates (i.e. inclusive of both
 * nights). Updates the duration input directly.
 */
function calculateDuration() {
  const start = document.querySelector('[name="tripStartDate"]').value;
  const end = document.querySelector('[name="tripEndDate"]').value;
  const durationInput = document.querySelector('[name="duration"]');
  if (!start || !end || !durationInput) return;
  const startDate = new Date(start);
  const endDate = new Date(end);
  if (startDate && endDate && startDate <= endDate) {
    const diffTime = Math.abs(endDate - startDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    durationInput.value = diffDays;
    durationInput.style.backgroundColor = '#f0f9ff';
  } else {
    durationInput.value = '';
    showError('End date must be after start date');
  }
}

/**
 * Fetch per diem rates either from the backend (if available) or
 * calculate fallback values based on the defaults in config.js. The
 * per diem values populate the M&IE, lodging and per diem total fields
 * and update the global perDiemRates object.
 */
function fetchPerDiemRates() {
  const city = document.querySelector('[name="city"]').value;
  const state = document.querySelector('[name="state"]').value;
  const duration = parseInt(document.querySelector('[name="duration"]').value) || 1;
  if (!city || !state) return;
  // Use Apps Script backend when available
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    updateProgress(30, 'Fetching GSA per diem rates...');
    google.script.run
      .withSuccessHandler((res) => {
        if (res && res.success) {
          const rates = res.data;
          const mealsTotal = rates.meals * duration;
          const lodgingTotal = rates.lodging * duration;
          perDiemRates = {
            meals: rates.meals,
            lodging: rates.lodging,
            mealsTotal,
            lodgingTotal,
            perDiemTotal: mealsTotal + lodgingTotal,
          };
          document.querySelector('[name="mealsCost"]').value = mealsTotal.toFixed(2);
          document.querySelector('[name="lodgingCost"]').value = lodgingTotal.toFixed(2);
          document.querySelector('[name="perDiemTotal"]').value = (mealsTotal + lodgingTotal).toFixed(2);
          showSuccess('Per diem rates loaded');
        } else {
          simulatePerDiem(duration);
        }
      })
      .withFailureHandler((err) => {
        console.error('Per diem error', err);
        simulatePerDiem(duration);
      })
      .getPerDiemRates(city, state);
  } else {
    simulatePerDiem(duration);
  }
}

/**
 * Populate per diem fields with default values when backend rates are
 * unavailable. Pulls defaults from CONFIG defined in config.js.
 * @param {number} duration The number of trip days
 */
function simulatePerDiem(duration) {
  const mie = CONFIG.DEFAULT_MIE || 79;
  const lod = CONFIG.DEFAULT_LODGING || 150;
  const mealsTotal = mie * duration;
  const lodgingTotal = lod * duration;
  perDiemRates = {
    meals: mie,
    lodging: lod,
    mealsTotal,
    lodgingTotal,
    perDiemTotal: mealsTotal + lodgingTotal,
  };
  document.querySelector('[name="mealsCost"]').value = mealsTotal.toFixed(2);
  document.querySelector('[name="lodgingCost"]').value = lodgingTotal.toFixed(2);
  document.querySelector('[name="perDiemTotal"]').value = (mealsTotal + lodgingTotal).toFixed(2);
  updateProgress(0, 'Ready to validate...');
}

/**
 * Validate car rental expenses against daily limits and justification
 * thresholds. Stores the result in expenseValidations and updates
 * summary and total.
 * @param {HTMLInputElement} input The car rental input element
 */
function validateCarRental(input) {
  const value = parseFloat(input.value) || 0;
  const duration = parseInt(document.querySelector('[name="duration"]').value) || 1;
  const dailyRate = duration > 0 ? value / duration : value;
  const div = document.getElementById('carRentalValidation');
  let validation = { isValid: true, level: 'success', message: '' };
  if (value === 0) {
    validation.message = '';
    input.classList.remove('border-red-500', 'border-yellow-500', 'border-green-500');
  } else if (dailyRate > CONFIG.EXPENSE_LIMITS.CAR_RENTAL_MAX_DAILY) {
    validation.isValid = false;
    validation.level = 'error';
    validation.message = `Daily rate (${dailyRate.toFixed(2)}) exceeds maximum (${CONFIG.EXPENSE_LIMITS.CAR_RENTAL_MAX_DAILY})`;
    input.classList.add('border-red-500');
    input.classList.remove('border-yellow-500', 'border-green-500');
  } else if (value > CONFIG.EXPENSE_LIMITS.CAR_RENTAL_JUSTIFICATION_THRESHOLD) {
    validation.level = 'warning';
    validation.message = `Total cost requires business justification (>${CONFIG.EXPENSE_LIMITS.CAR_RENTAL_JUSTIFICATION_THRESHOLD})`;
    input.classList.add('border-yellow-500');
    input.classList.remove('border-red-500', 'border-green-500');
  } else {
    validation.message = `Car rental cost is within policy limits (${dailyRate.toFixed(2)}/day)`;
    input.classList.add('border-green-500');
    input.classList.remove('border-red-500', 'border-yellow-500');
  }
  if (div) {
    div.textContent = validation.message;
    div.className = `text-xs mt-1 ${validation.level === 'error' ? 'text-red-600' : validation.level === 'warning' ? 'text-yellow-600' : 'text-green-600'}`;
  }
  expenseValidations.carRental = validation;
  updateValidationSummary();
  validateAndCalculate();
}

/**
 * Validate parking costs against daily limits. Parking never results
 * in a hard error but a warning is shown if the daily rate exceeds
 * the limit. Updates summary and totals accordingly.
 * @param {HTMLInputElement} input The parking cost input
 */
function validateParking(input) {
  const value = parseFloat(input.value) || 0;
  const duration = parseInt(document.querySelector('[name="duration"]').value) || 1;
  const dailyRate = duration > 0 ? value / duration : value;
  const div = document.getElementById('parkingValidation');
  let validation = { isValid: true, level: 'success', message: '' };
  if (value === 0) {
    validation.message = '';
    input.classList.remove('border-yellow-500', 'border-green-500');
  } else if (dailyRate > CONFIG.EXPENSE_LIMITS.PARKING_MAX_DAILY) {
    validation.level = 'warning';
    validation.message = `Daily parking (${dailyRate.toFixed(2)}) exceeds recommended limit (${CONFIG.EXPENSE_LIMITS.PARKING_MAX_DAILY})`;
    input.classList.add('border-yellow-500');
    input.classList.remove('border-green-500');
  } else {
    validation.message = `Parking cost is within guidelines (${dailyRate.toFixed(2)}/day)`;
    input.classList.add('border-green-500');
    input.classList.remove('border-yellow-500');
  }
  if (div) {
    div.textContent = validation.message;
    div.className = `text-xs mt-1 ${validation.level === 'warning' ? 'text-yellow-600' : 'text-green-600'}`;
  }
  expenseValidations.parking = validation;
  updateValidationSummary();
  validateAndCalculate();
}

/**
 * Validate conference or training fees against the maximum and
 * justification thresholds. An error is shown if the maximum is
 * exceeded, or a warning if pre‑approval is required.
 * @param {HTMLInputElement} input The conference fee input
 */
function validateConferenceFee(input) {
  const value = parseFloat(input.value) || 0;
  const div = document.getElementById('conferenceValidation');
  let validation = { isValid: true, level: 'success', message: '' };
  if (value === 0) {
    validation.message = '';
    input.classList.remove('border-red-500', 'border-yellow-500', 'border-green-500');
  } else if (value > CONFIG.EXPENSE_LIMITS.CONFERENCE_FEE_MAX) {
    validation.isValid = false;
    validation.level = 'error';
    validation.message = `Conference fee exceeds maximum allowed (${CONFIG.EXPENSE_LIMITS.CONFERENCE_FEE_MAX})`;
    input.classList.add('border-red-500');
    input.classList.remove('border-yellow-500', 'border-green-500');
  } else if (value > CONFIG.EXPENSE_LIMITS.CONFERENCE_JUSTIFICATION_THRESHOLD) {
    validation.level = 'warning';
    validation.message = `Pre-approval required for fees >${CONFIG.EXPENSE_LIMITS.CONFERENCE_JUSTIFICATION_THRESHOLD}`;
    input.classList.add('border-yellow-500');
    input.classList.remove('border-red-500', 'border-green-500');
  } else {
    validation.message = `Conference fee is within policy limits`;
    input.classList.add('border-green-500');
    input.classList.remove('border-red-500', 'border-yellow-500');
  }
  if (div) {
    div.textContent = validation.message;
    div.className = `text-xs mt-1 ${validation.level === 'error' ? 'text-red-600' : validation.level === 'warning' ? 'text-yellow-600' : 'text-green-600'}`;
  }
  expenseValidations.conferenceFee = validation;
  updateValidationSummary();
  validateAndCalculate();
}

/**
 * Validate miscellaneous expenses against the receipts threshold. Misc
 * expenses never result in a hard error; a warning appears when the
 * threshold is exceeded. Updates summary and totals accordingly.
 * @param {HTMLInputElement} input The miscellaneous expense input
 */
function validateMiscellaneous(input) {
  const value = parseFloat(input.value) || 0;
  const div = document.getElementById('miscValidation');
  let validation = { isValid: true, level: 'success', message: '' };
  if (value === 0) {
    validation.message = '';
    input.classList.remove('border-yellow-500', 'border-green-500');
  } else if (value > CONFIG.EXPENSE_LIMITS.MISC_JUSTIFICATION_THRESHOLD) {
    validation.level = 'warning';
    validation.message = `Detailed receipts required for misc expenses >${CONFIG.EXPENSE_LIMITS.MISC_JUSTIFICATION_THRESHOLD}`;
    input.classList.add('border-yellow-500');
    input.classList.remove('border-green-500');
  } else {
    validation.message = `Miscellaneous expense is within policy limits`;
    input.classList.add('border-green-500');
    input.classList.remove('border-yellow-500');
  }
  if (div) {
    div.textContent = validation.message;
    div.className = `text-xs mt-1 ${validation.level === 'warning' ? 'text-yellow-600' : 'text-green-600'}`;
  }
  expenseValidations.miscellaneous = validation;
  updateValidationSummary();
  validateAndCalculate();
}

/**
 * Summarise all expenses and per diem values to compute the total
 * claimed cost. Updates the calculatedTotal field, otherExpensesTotal
 * field and triggers total cost validation if a claimed cost exists.
 */
function validateAndCalculate() {
  const airfare = parseFloat(document.querySelector('[name="airfare"]').value) || 0;
  const carRental = parseFloat(document.querySelector('[name="carRental"]').value) || 0;
  // Select the ground transportation field using its actual name attribute
  const groundTransport = parseFloat(document.querySelector('[name="transportation"]').value) || 0;
  const parking = parseFloat(document.querySelector('[name="parking"]').value) || 0;
  // Pull the conference/training fee using the correct field name
  const conference = parseFloat(document.querySelector('[name="conferenceFee"]').value) || 0;
  const miscellaneous = parseFloat(document.querySelector('[name="miscellaneous"]').value) || 0;
  const lodging = parseFloat(document.querySelector('[name="lodgingCost"]').value) || 0;
  const meals = parseFloat(document.querySelector('[name="mealsCost"]').value) || 0;
  const otherExpensesTotal = conference + miscellaneous;
  const otherField = document.querySelector('[name="otherExpensesTotal"]');
  if (otherField) otherField.value = otherExpensesTotal.toFixed(2);
  const calculatedTotal = airfare + carRental + groundTransport + parking + otherExpensesTotal + lodging + meals;
  const calcField = document.querySelector('[name="calculatedTotal"]');
  if (calcField) calcField.value = calculatedTotal.toFixed(2);
  const claimedTotal = parseFloat(document.querySelector('[name="totalCost"]').value) || 0;
  if (claimedTotal) {
    validateTotalCost();
  }
}

/**
 * Compare the claimed total cost against the calculated total. A
 * warning is shown if the variance exceeds the configured threshold
 * but is still within acceptable limits, and an error if it exceeds
 * the threshold. Updates the validation summary and funding
 * availability check as required.
 */
function validateTotalCost() {
  const calculatedTotal = parseFloat(document.querySelector('[name="calculatedTotal"]').value) || 0;
  const claimedTotal = parseFloat(document.querySelector('[name="totalCost"]').value) || 0;
  const div = document.getElementById('totalCostValidation');
  const variance = claimedTotal - calculatedTotal;
  const variancePercent = calculatedTotal > 0 ? Math.abs(variance) / calculatedTotal : 0;
  let message = '';
  let level = 'success';
  if (variance === 0) {
    message = 'Claimed total matches auto-calculated total';
    level = 'success';
  } else if (variancePercent <= CONFIG.EXPENSE_LIMITS.TOTAL_VARIANCE_THRESHOLD) {
    message = `Claimed total differs by ${variance.toFixed(2)} (${(variancePercent * 100).toFixed(1)}%), within acceptable variance`;
    level = 'warning';
  } else {
    message = `Claimed total differs by ${variance.toFixed(2)} (${(variancePercent * 100).toFixed(1)}%), exceeds acceptable variance`;
    level = 'error';
  }
  if (div) {
    div.textContent = message;
    div.className = `text-sm font-medium ${level === 'error' ? 'text-red-600' : level === 'warning' ? 'text-yellow-600' : 'text-green-600'}`;
  }
  expenseValidations.totalCost = { isValid: level !== 'error', level: level, message: message };
  if (selectedTaskOrder) {
    validateFundingAvailability(claimedTotal);
  }
  updateValidationSummary();
}

/**
 * Fill the claimed total cost with the auto-calculated value. This
 * function is bound to the Auto‑Calculate button in the UI and then
 * triggers validation on the total cost.
 */
function autoCalculateTotal() {
  const calcField = document.querySelector('[name="calculatedTotal"]');
  const totalField = document.querySelector('[name="totalCost"]');
  if (calcField && totalField) {
    totalField.value = calcField.value;
    validateTotalCost();
  }
}

/**
 * Reset all form fields and hide any result/summary panels. Clears
 * internal state used for validations and resets the progress bar.
 */
function clearForm() {
  const form = document.getElementById('tar-form');
  if (form) {
    form.reset();
  }
  // Reset internal state
  expenseValidations = {};
  selectedTaskOrder = null;
  currentFile = null;
  perDiemRates = null;
  currentValidation = null;
  // Hide panels
  const fundingInfo = document.getElementById('fundingInfo');
  if (fundingInfo) fundingInfo.classList.add('hidden');
  const fundingWarning = document.getElementById('fundingWarning');
  if (fundingWarning) fundingWarning.classList.add('hidden');
  const summary = document.getElementById('expenseValidationSummary');
  if (summary) summary.classList.add('hidden');
  const result = document.getElementById('result');
  if (result) result.classList.add('hidden');
  const analyst = document.getElementById('analystReviewSection');
  if (analyst) analyst.classList.add('hidden');
  const breakdown = document.getElementById('breakdown');
  if (breakdown) breakdown.classList.add('hidden');
  // Clear validation messages
  ['carRentalValidation', 'parkingValidation', 'conferenceValidation', 'miscValidation', 'totalCostValidation'].forEach((id) => {
    const elem = document.getElementById(id);
    if (elem) elem.textContent = '';
  });
  updateProgress(0, 'Ready to validate...');
}

/**
 * Build a summary of all validation messages and display them in the
 * expense validation summary section. Hides the section when there
 * are no messages to display.
 */
function updateValidationSummary() {
  const summaryDiv = document.getElementById('expenseValidationSummary');
  const contentDiv = document.getElementById('validationSummaryContent');
  if (!summaryDiv || !contentDiv) return;
  const items = [];
  Object.keys(expenseValidations).forEach((key) => {
    const val = expenseValidations[key];
    if (val && val.message) {
      items.push(`<div class="validation-alert validation-${val.level}">${val.message}</div>`);
    }
  });
  if (items.length > 0) {
    contentDiv.innerHTML = items.join('');
    summaryDiv.classList.remove('hidden');
  } else {
    summaryDiv.classList.add('hidden');
  }
}

/**
 * Test connectivity to the GSA API. When running inside Apps Script
 * this calls the testGSAAPI function; otherwise a warning is shown.
 */
function testGSAConnection() {
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    updateProgress(10, 'Testing GSA API connectivity...');
    google.script.run
      .withSuccessHandler((res) => {
        if (res && res.success) {
          showSuccess(res.message || 'GSA API connectivity confirmed');
        } else {
          showError(res && res.message ? res.message : 'GSA API test failed');
        }
      })
      .withFailureHandler((err) => {
        console.error('GSA API test error:', err);
        showError('GSA API test failed');
      })
      .testGSAAPI();
  } else {
    showWarning('GSA API test not available in this environment');
  }
}

/**
 * Handle form submission. Serializes form data and passes it to
 * validateTarWithPerDiem on the backend. Shows a simulated result
 * when running locally. Provides feedback via the progress bar and
 * toggles the loading icon.
 * @param {Event} event The form submit event
 */
function handleSubmit(event) {
  event.preventDefault();
  const loadingIcon = document.getElementById('loading-icon');
  const submitText = document.getElementById('submit-text');
  if (loadingIcon) loadingIcon.style.display = '';
  if (submitText) submitText.textContent = 'Validating...';
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    const data = collectFormData();
    updateProgress(30, 'Submitting for validation...');
    google.script.run
      .withSuccessHandler((res) => {
        if (loadingIcon) loadingIcon.style.display = 'none';
        if (submitText) submitText.textContent = 'Validate Trip';
        if (res && res.success) {
          currentValidation = res;
          showValidationResults(res);
        } else {
          showError(res && res.message ? res.message : 'Validation failed');
        }
      })
      .withFailureHandler((err) => {
        if (loadingIcon) loadingIcon.style.display = 'none';
        if (submitText) submitText.textContent = 'Validate Trip';
        console.error('Validation error:', err);
        showError('Validation failed');
      })
      .validateTarWithPerDiem(data);
  } else {
    // Simulate a result for offline testing
    if (loadingIcon) loadingIcon.style.display = 'none';
    if (submitText) submitText.textContent = 'Validate Trip';
    const result = {
      success: true,
      isValid: true,
      message: 'Trip is within per diem',
      expectedCost: '1500.00',
      claimedCost: '1500.00',
      variance: '0.00',
      variancePercent: 0,
      breakdown: [],
    };
    currentValidation = result;
    showValidationResults(result);
  }
}

/**
 * Serialize all form data into a plain object. Converts numeric
 * values where appropriate and includes the selected task order.
 * @returns {Object} A data object for validation
 */
function collectFormData() {
  const form = document.getElementById('tar-form');
  const formData = new FormData(form);
  const data = {};
  formData.forEach((val, key) => {
    data[key] = val;
  });
  data.taskOrder = selectedTaskOrder ? selectedTaskOrder.awardPiid : data.taskOrder;
  data.duration = parseInt(data.duration) || 1;
  data.totalCost = parseFloat(data.totalCost) || 0;
  // PDF content can be attached via the Apps Script file object; for local testing we omit it
  data.pdfContent = null;
  return data;
}

/**
 * Render validation results into the result section. Also populates
 * the breakdown and analyst review panels. Provides a simple risk
 * assessment based on variance percent.
 * @param {Object} res The result returned from validateTarWithPerDiem
 */
function showValidationResults(res) {
  const resultSection = document.getElementById('result');
  const content = document.getElementById('result-content');
  if (!resultSection || !content) return;
  const msg = res.message || '';
  const expected = res.expectedCost;
  const claimed = res.claimedCost;
  const variance = res.variance;
  const variancePercent = res.variancePercent;
  let html = `<p class="mb-2">${msg}</p>`;
  html += '<ul class="list-disc ml-5 mb-3">';
  html += `<li><strong>Expected Cost:</strong> $${expected}</li>`;
  html += `<li><strong>Claimed Cost:</strong> $${claimed}</li>`;
  html += `<li><strong>Variance:</strong> $${variance} (${variancePercent}%)</li>`;
  html += '</ul>';
  if (res.breakdown && res.breakdown.length > 0) {
    html += '<div class="mb-3"><strong>Per Diem Breakdown:</strong></div>';
    html += '<div class="overflow-x-auto"><table class="min-w-full border divide-y divide-gray-200 text-sm">';
    html += '<thead><tr><th class="px-3 py-1 text-left font-semibold">Location</th><th class="px-3 py-1 text-left">Date</th><th class="px-3 py-1 text-right">M&IE</th><th class="px-3 py-1 text-right">Lodging</th><th class="px-3 py-1 text-right">Total</th></tr></thead>';
    html += '<tbody>';
    res.breakdown.forEach((item) => {
      html += `<tr><td class="px-3 py-1">${item.location}</td><td class="px-3 py-1">${item.date}</td><td class="px-3 py-1 text-right">${item.mie}</td><td class="px-3 py-1 text-right">${item.lodging}</td><td class="px-3 py-1 text-right">${item.total}</td></tr>`;
    });
    html += '</tbody></table></div>';
  }
  content.innerHTML = html;
  resultSection.classList.remove('hidden');
  // Show per diem breakdown separately
  const breakdownSection = document.getElementById('breakdown');
  const breakdownContent = document.getElementById('breakdown-content');
  if (res.breakdown && res.breakdown.length > 0 && breakdownSection && breakdownContent) {
    breakdownContent.innerHTML = html;
    breakdownSection.classList.remove('hidden');
  }
  // Show analyst review with summaries
  const analystSec = document.getElementById('analystReviewSection');
  if (analystSec) {
    const statusDiv = document.getElementById('validationStatusDisplay');
    const expenseAnalysis = document.getElementById('expenseAnalysisDisplay');
    const riskAssessment = document.getElementById('riskAssessmentDisplay');
    if (statusDiv) {
      statusDiv.textContent = res.isValid ? 'PASS' : 'FAIL';
      statusDiv.className = `p-4 rounded-lg border ${res.isValid ? 'bg-green-50 border-green-200 text-green-700' : 'bg-red-50 border-red-200 text-red-700'}`;
    }
    if (expenseAnalysis) {
      let analysisHtml = '';
      Object.keys(expenseValidations).forEach((key) => {
        const val = expenseValidations[key];
        if (val && val.message) {
          analysisHtml += `<div class="validation-alert validation-${val.level}">${val.message}</div>`;
        }
      });
      expenseAnalysis.innerHTML = analysisHtml || '<p>No detailed analysis available.</p>';
    }
    if (riskAssessment) {
      let risk = 'Low';
      const variancePct = res.variancePercent;
      if (variancePct > CONFIG.EXPENSE_LIMITS.TOTAL_VARIANCE_THRESHOLD * 100) {
        risk = 'High';
      } else if (variancePct > (CONFIG.EXPENSE_LIMITS.TOTAL_VARIANCE_THRESHOLD * 100) / 2) {
        risk = 'Medium';
      }
      riskAssessment.textContent = risk;
    }
    analystSec.classList.remove('hidden');
  }
}

/**
 * Analyst actions: these functions currently show simple messages but
 * could be extended to call backend functions (e.g. to record
 * decisions or generate reports).
 */
function analystApprove() {
  if (!currentValidation) {
    showError('No validation data available for approval');
    return;
  }
  showSuccess('Trip approved for management review');
}

function analystReject() {
  showWarning('Trip rejected - needs revision');
}

function requestMoreInfo() {
  showWarning('Requested additional information from traveler');
}

/**
 * Hide the management modal, used when closing the document preview.
 */
function closeManagementModal() {
  const modal = document.getElementById('managementModal');
  if (modal) modal.classList.add('hidden');
}

/**
 * Download the management report. In Apps Script this would return a
 * blob; here we simply show a success message for demonstration.
 */
function downloadPDF() {
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler((res) => {
        showSuccess('Management report downloaded');
      })
      .withFailureHandler((err) => {
        showError('Download failed');
      })
      .exportValidationResults('pdf');
  } else {
    showWarning('Download not available in this environment');
  }
}

/**
 * Email the management report. As with downloadPDF this is a stub
 * function when running locally.
 */
function emailToManagement() {
  showSuccess('Email sent to management (simulated)');
}

/**
 * Remove the currently selected document from the upload area. This allows
 * users to replace or delete a previously selected file. Resets the
 * hidden file input, clears the global currentFile reference and hides
 * the file info panel. Called when the "Remove File" button in the
 * document upload section is clicked.
 */
function removeSelectedFile() {
  const fileInput = document.getElementById('tarFile');
  currentFile = null;
  if (fileInput) {
    // Reset the file input so the same file can be selected again if needed
    fileInput.value = '';
  }
  const fileInfo = document.getElementById('file-info');
  if (fileInfo) {
    fileInfo.classList.add('hidden');
  }
}

/**
 * Perform a simplified TAR validation on the client when running
 * offline (i.e. not connected to Apps Script).  This function
 * calculates an expected cost using default MIE and lodging rates
 * defined in CONFIG and compares it against the claimed total cost
 * provided in the data.  It returns an object resembling the
 * structure produced by validateTarWithPerDiem() on the server.  The
 * result contains `success`, `isValid`, `message` and a minimal
 * `validationReport` for downstream display.
 *
 * @param {Object} data The mapped TAR data object
 * @returns {Object} An object containing validation result info
 */
function offlineValidateTar(data) {
  try {
    const duration = parseInt(data.duration) || 1;
    const totalCost = parseFloat(data.totalCost) || 0;
    const estimatedCost = parseFloat(data.estimatedCost) || totalCost;
    // Use default M&IE and lodging rates since no API is available
    const mie = CONFIG.DEFAULT_MIE || 0;
    const lodging = CONFIG.DEFAULT_LODGING || 0;
    const expected = (mie + lodging) * duration;
    const variance = totalCost - expected;
    const variancePercent = expected ? (variance / expected) * 100 : 0;
    // Determine validity based on buffer and deviation thresholds
    const buffer = CONFIG.COST_BUFFER || 0;
    const maxDeviation = CONFIG.MAX_DEVIATION_PERCENT || 100;
    const isWithinBuffer = Math.abs(variance) <= buffer;
    const isWithinDeviation = Math.abs(variancePercent) <= maxDeviation;
    const isValid = isWithinBuffer || isWithinDeviation;
    const message = isValid
      ? '✅ Offline validation passed.'
      : `⚠️ Trip cost validation failed. Variance: ${variancePercent.toFixed(2)}%`;
    return {
      success: true,
      isValid: isValid,
      validationReport: {
        validation: {
          variance: variance,
          variancePercent: variancePercent.toFixed(2),
          isWithinBuffer: isWithinBuffer,
          isWithinDeviation: isWithinDeviation,
          messages: [message],
        },
        errors: [],
        warnings: [],
        messages: [message],
      },
      message: message,
    };
  } catch (err) {
    return {
      success: false,
      errors: [err.message || 'Offline validation error'],
    };
  }
}

//
// Expose selected functions globally for inline event handlers
//
// In the HTML markup we use inline attributes like `onsubmit="handleSubmit(event)"`
// and `onchange="handleFileSelect(event)"`. When this script is executed as a
// module in some environments (e.g. bundlers or strict scopes), top‑level
// function declarations are not automatically attached to the `window` object.
// To ensure these functions are available to the inline handlers when
// deployed in Google Apps Script or served directly from the browser, we
// explicitly assign them to `window` if it exists. This prevents ReferenceError
// exceptions like `handleFileSelect is not defined`.
if (typeof window !== 'undefined') {
  // Core form submission and file handling
  window.handleSubmit = handleSubmit;
  window.handleFileSelect = handleFileSelect;
  window.extractDataFromPdf = extractDataFromPdf;
  window.removeSelectedFile = removeSelectedFile;
  // Utility hooks used in inline attributes
  window.calculateDuration = calculateDuration;
  window.fetchPerDiemRates = fetchPerDiemRates;
  window.autoCalculateTotal = autoCalculateTotal;
  window.clearForm = clearForm;
  window.testGSAConnection = testGSAConnection;
  // Expense validators referenced directly from HTML
  window.validateCarRental = validateCarRental;
  window.validateParking = validateParking;
  window.validateConferenceFee = validateConferenceFee;
  window.validateMiscellaneous = validateMiscellaneous;
  window.validateTotalCost = validateTotalCost;
  // Analyst actions and management modal controls
  window.analystApprove = analystApprove;
  window.analystReject = analystReject;
  window.requestMoreInfo = requestMoreInfo;
  window.closeManagementModal = closeManagementModal;
  window.downloadPDF = downloadPDF;
  window.emailToManagement = emailToManagement;

  // Bulk upload handlers
  window.handleBulkFileSelect = handleBulkFileSelect;
  window.processBulkCsv = processBulkCsv;
}