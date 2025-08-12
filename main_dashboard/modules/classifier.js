// ===== DOCUMENT CLASSIFIER

function classifyDocument(content, fileName) {
  console.log(`Classifying Doc: ${fileName}`);

  const contentLower = content.toLowerCase();
  const fileNameLower = fileName.toLowerCase();

  // TAR Classification text
  const tarKeywords = [
    "travel authorization request",
    "tar",
    "deliverable",
    "contract travel",
    "resource utilization",
    "status report",
    "per diem",
    "rates",
  ];

  // RIP Classification text
  const ripKeywords = [
    "rip",
    "request to initialize a purchase",
    "request for initial purchase",
  ];

  // INVOICE Classification text
  const invoiceKeywords = [
    "invoice",
    "billing",
    "payment request",
    "labor hours",
    "expense report",
    "cost reimbursement",
    "billing period",
    "invoice number",
    "amount due",
    "payment terms",
  ];

  // calculating keyword scores
  const tarScore = claculateKeywordScore(
    contentLower + " " + fileNameLower,
    tarKeywords
  );
  const ripScore = calculateKeywordScore(
    contentLower + " " + fileNameLower,
    ripKeywords
  );
  const invoicescore = calculateKeywordScore(
    contentLower + " " + fileNameLower,
    invoiceKeywords
  );

  console.log(
    `Classification scores - TAR: ${tarScore}, RIP: ${ripScore}, Invoice: ${invoicescore}`
  );

  const maxScore = Math.max(tarScore, ripScore, invoiceScore);

  if (maxScore === 0) {
    return "UNKNOWN";
  } else if (tarScore === maxScore) {
    return "TAR";
  } else if (ripScore === maxScore) {
    return "RIP";
  } else {
    return "INVOICE";
  }
}

// ===== KEYWORD MATCH

function calculateKeywordScore(text, keywords) {
  let score = 0;
  keywords.forEach((keyword) => {
    const regex = new RegExp(keyword, "gi");
    const matches = text.match(regex);
    if (matches) {
      score += matches.length;
    }
  });
  return score;
}
