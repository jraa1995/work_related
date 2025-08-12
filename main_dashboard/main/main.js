// ===== MAIN CONTROLLER

function doGet(e) {
  const page = e.parameter.page || "upload";

  switch (page) {
    case "upload":
      return HtmlService.createTemplateFromFile("Upload")
        .evaluate()
        .setTitle("TAR/RIP/Invoice Validation System")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    case "dashboard":
      return (
        HtmlService.createTemplateFromFile("Dashboard"),
        evaluate()
          .setTitle("Validation Dashboard")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      );
    default:
      return HtmlService.createHtmlOutput("<h1>Page not found</h1>");
  }
}

// ===== HTML TEMPLATING

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== PROCESS DOC

function processUpload(fileBlob, fileName, documentType) {
  try {
    console.log(`Processing Upload: ${fileName}, Type: ${documentType}`);

    // 1. save file to folder
    const fileId = saveToFolder(fileBlob, fileName, "incoming");

    // 2. extract text
    const content = extractTextContent(fileId, fileName);

    // 3. classify doc if not specified
    const classifiedType = documentType || classifyDocument(content, fileName);

    // 4. validate
    const validationResults = validateDocument(
      classifiedType,
      content,
      fileName
    );

    // 5. calc risk score
    const riskScore = calculateRiskScore(validationResults);

    // 6. save results in sheeet
    const trackingId = saveToTrackingSheet({
      fileName: fileName,
      fileId: fileId,
      documentType: classifiedType,
      validationResults: validationResults,
      riskScore: riskScore,
      status:
        riskScore > 85
          ? "AUTO_APPROVED"
          : riskScore > 70
          ? "REVIEW_REQUIRED"
          : "MANUAL_REVIEW",
      uploadDate: new Date(),
      uploadedBy: Session.getActiveUser().getEmail(),
    });

    // 7. move file == risk
    movetoProcessedFolder(fileId, riskScore);

    // 8. send notifs
    sendNotifications(trackingId, validationResults, riskScore);

    return {
      success: true,
      trackingId: trackingId,
      documentType: classifiedType,
      riskScore: riskScore,
      validationResults: validationResults,
    };
  } catch (error) {
    console.error("Error processing upload:", error);
    return {
      success: false,
      error: error.message,
    };
  }
}
