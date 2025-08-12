// ===== DOCUMENT VALDATION

function validateDocument(documentType, content, fileName) {
  console.log(`Validating ${documentType} document: ${fileName}`);

  switch (documentType) {
    case "TAR":
      return validateTAR(content, fileName);
    case "RIP":
      return validateRIP(content, fileName);
    case "INVOICE":
      return validateInvoice(content, fileName);
    default:
      return {
        documentType: documentType,
        completenessScore: 0,
        complianceScore: 0,
        formatScore: 50,
        issues: ["Unknown Document type - manual review required"],
        validationDetails: {},
      };
  }
}

// ===== TAR VALIDATION

function validateTAR(content, fileName) {
  const issues = [];
  const validationDetails = {};

  // required fields for TAR
  const requiredSections = {
    "Contract Information": ["contract number", "contract", "agreement"],
    "Reporting Period": [
      "period",
      "quarter",
      "fy",
      "qtr",
      "month",
      "reporting period",
    ],
    "Executive Summary": ["executive summary", "summary", "overview"],
    "Travel Approach": ["travel approach", "methodology", "technical"],
    "Risk Assessment": ["risk", "risk assessment", "mitigation"],
    "Resource Utilization": ["resource", "personnel", "staff", "hours"],
    "Deliverable Status": ["deliverable", "milestone", "completion"],
  };

  // check for required sections
  let sectionsFound = 0;
  const contentLower = content.toLowerCase();

  Object.keys(requiredSections).forEach((section) => {
    const keywords = requiredSections[section];
    const found = keywords.some((keyword) => contentLower.includes(keyword));
    validationDetails[section] = found;

    if (found) {
      sectionsFound++;
    } else {
      issues.push(`Missing required Sections: ${section}`);
    }
  });

  // format validation
  const formatChecks = validateTARFormat(content, fileName);
  issues.push(...formatChecks.issues);

  // compliance validation
  const complianceChecks = validateTARCompliance(content);
  issues.push(...complianceChecks.issues);

  // calculate scores
  const completenessScore = Math.round(
    (sectionsFound / Object.keys(requiredSections).length) * 100
  );
  const complianceScore = complianceChecks.score;
  const formatScore = formatChecks.score;

  return {
    documentType: "TAR",
    completenessScore: completenessScore,
    complianceScore: complianceScore,
    formatScore: formatScore,
    issues: issues,
    validationDetails: validationDetails,
  };
}

// ===== VALIDATE RIP

function validateRIP(content, fileName) {
  const issues = [];
  const validationDetails = {};

  const requiredSections = {
    "Technical Volume": ["technical volumne", "technical approach", "solution"],
    "Management Volume": [
      "management",
      "management approach",
      "project management",
    ],
    "Cost Volumne": ["cost", "pricing", "cost breakdown", "budget"],
    "Past Performance": ["past performance", "experience", "references"],
    "Key Personnel": ["key personnel", "staff", "team", "resume"],
    "Compliance Matrix": ["compliance", "requirements", "matrix"],
  };

  let sectionsFound = 0;
  const contentLower = content.toLowerCase();

  Object.keys(requiredSections).forEach((section) => {
    const keywords = requiredSections[section];
    const found = keywords.some((keyword) => contentLower.includes(keyword));
    validationDetails[section] = found;

    if (found) {
      sectionsFound++;
    } else {
      issues.push(`Missing required section: ${section}`);
    }
  });

  const formatChecks = validateRIPFormat(content, fileName);
  issues.push(...formatChecks.issues);

  const completenessScore = Math.round(
    (sectionsFound / Object.keys(requiredSections).length) * 100
  );

  return {
    documentType: "RIP",
    completenessScore: completenessScore,
    complianceScore: formatChecks.score,
    formatScore: formatChecks.score,
    issues: issues,
    validationDetails: validationDetails,
  };
}

// ===== VALIDATE INVOICE

function validateInvoice(content, fileName) {
  const issues = [];
  const validationDetails = {};

  const requiredElements = {
    "Invoice Number": ["invoice number", "invoice #", "inv #"],
    "Contract Number": ["contract", "agreement", "po number"],
    "Billing Period": ["billing period", "period", "invoice period"],
    "Labor Hours": ["hours", "labor", "time"],
    "Hourly Rates": ["rate", "billing rate", "hourly"],
    "Total Amount": ["total", "amount due", "invoice total"],
    "Tax Information": ["tax", "gst", "vat"],
  };

  let elementsFound = 0;
  const contentLower = content.toLowerCase();

  Object.keys(requiredElements).forEach((element) => {
    const keywords = requiredElements[element];
    const found = keywords.some((keyword) => contentLower.includes(keyword));
    validationDetails[element] = found;

    if (found) {
      elementsFound++;
    } else {
      issues.push(`Missing required element: ${element}`);
    }
  });

  // financial validation
  const financialChecks = validateInvoiceFinancials(content);
  issues.push(...financialChecks.issues);

  const completenessScore = Math.round(
    (elementsFound / Object.keys(requiredElements).length) * 100
  );

  return {
    documentType: "INVOICE",
    completenessScore: completenessScore,
    complianceScore: financialChecks.score,
    formatScore: 85, // basic format score
    issues: issues,
    validationDetails: validationDetails,
  };
}

// ===== VALIDATE TAR FORMAT

function validateTARFormat(content, fileName) {
  const issues = [];
  let score = 100;

  // check file extension
  if (!fileName.match(/\.(pdf|docx|doc)$/i)) {
    issues.push("Invalid file format - must be PDF or Word document");
    score -= 20;
  }

  // check document length
  const wordCount = content.split(/\s+/).length;
  if (wordCount < 1000) {
    issues.push("Document appears too short for a comprehensive TAR");
    score -= 15;
  }

  // check for tables/structured data
  if (!content.match(/\d+\.\d+|\d+%|table|figure/i)) {
    issues.push("Document may lack required tables or figures");
    score -= 10;
  }

  return { issues, score: Math.max(0, score) };
}

// ===== VALIDATE TAR COMPLIANCE

function validateTARCompliance(content) {
  const issues = [];
  let score = 100;

  const contentLower = content.toLowerCase();

  // for contract reference
  if (!contentLower.match(/contract\s+\w+\d+/)) {
    issues.push("Missing proper contract number reference");
    score -= 20;
  }

  // for reporting period
  if (
    !contentLower.match(
      /(quarter|q[1-4]|january|february|march|april|may|june|july|august|september|october|november|december)/
    )
  ) {
    issues.push("Missing clear reporting period identification");
    score -= 15;
  }

  // for signatures/approvals
  if (!contentLower.match(/(signature|approved|reviewed|certified)/)) {
    issues.push("Missing approval or signature indicators");
    score -= 10;
  }

  return { issues, score: Math.max(0, score) };
}

// ===== RIP FORMAT

function validateRIPFormat(content, fileName) {
  const issues = [];
  let score = 100;

  if (!fileName.match(/\.(pdf|docx|doc)$/i)) {
    issues.push("Invalid file format for RIP submission");
    score -= 25;
  }

  const wordCount = content.split(/\s+/).length;
  if (wordCount < 2000) {
    issues.push("RIP response appears too brief");
    score -= 20;
  }

  return { issues, score: Math.max(0, score) };
}

// ===== INVOICE FINANCIAL DATA

function validateInvoiceFinancials(content) {
  const issues = [];
  let score = 100;

  // currency amounts
  const amounts = content.match(/\$[\d,]+\.?\d*/g);
  if (!amounts || amounts.length === 0) {
    issues.push("No currency amounts found in invoice");
    score -= 30;
  }

  // mathematical calculations
  if (!content.match(/\d+\s*[x*]\s*\d+|\d+\s*hours?/i)) {
    issues.push("Missing calculation details (hours × rate)");
    score -= 20;
  }

  return { issues, score: Math.max(0, score) };
}

// ===== OVERALL RISK
function calculateRiskScore(validationResults) {
  const weights = {
    completenessScore: 0.4,
    complianceScore: 0.3,
    formatScore: 0.2,
    issuesPenalty: 0.1,
  };

  const issuesPenalty = Math.max(0, 100 - validationResults.issues.length * 5);

  const weightedScore =
    validationResults.completenessScore * weights.completenessScore +
    validationResults.complianceScore * weights.complianceScore +
    validationResults.formatScore * weights.formatScore +
    issuesPenalty * weights.issuesPenalty;

  return Math.round(Math.max(0, Math.min(100, weightedScore)));
}
