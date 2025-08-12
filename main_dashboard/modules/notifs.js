// ===== SEND NOTIFS

function sendNotifications(trackingId, validationResults, riskScore) {
  try {
    const userEmail = sessionStorage.getActiveUser().getEmail();

    if (riskScore <= 70) {
      //high risk - immediate notif
      sendHighRiskNotification(trackingId, validationResults, userEmail);
    } else if (riskScore <= 85) {
      sendStandardNotification(trackingId, validationResults, userEmail);
    }

    // low risk docs auto-approved
    console.log(`Notifications sent for ${trackingId}`);
  } catch (error) {
    console.error("error sending notifications:", error);
  }
}

// SEND HIGH RISK NOTIF

function sendHighRiskNotification(trackingId, validationResults, userEmail) {
  const subject = `HIGH RISK Document Requires Review - ${trackingId}`;

  const body = `
    A document has been flagged as HIGH RISK and requires immediate manual review.
    
    Tracking ID: ${trackingId}
    Document Type: ${validationResults.documentType}
    Risk Score: ${calculateRiskScore(validationResults)}/100


    VALIDATION ISSUES:
    ${validationResults.issues.map((issue) => `• ${issue}`).join("\n")}


    SCORES:
    • Completeness: ${validationResults.completenessScore}%
    • Compliance: ${validationResults.complianceScore}%
    • Format: ${validationResults.formatScore}%


    Please review this document immediately. 
    
    View Dashboard: ${getWebAppUrl()}?page=dashboard
    `;

  // send to users/supervisors
  const recipients = [userEmail, getProperty("SUPERVISOR_EMAIL")].filter(
    (email) => email
  );

  recipients.forEach((email) => {
    try {
      GmailApp.sendEmail(email, subject, body);
    } catch (error) {
      console.error(`Failed to send meail to ${email}:`, error);
    }
  });
}

// SEND STANDARD NOTIF

function sendStandardNotification(trackingId, validationResults, userEmail) {
  const subject = `Document Review Required - ${trackingId}`;

  const body = `
    A document has been processed and requires review. 


    Tracking ID: ${trackingId}
    Document Type: ${validationResults.documentType}
    Risk Score: ${calculateRiskScore(validationResults)}/100


    Issues Found: ${validationResults.issues.length}
    ${
      validationResults.issues.length > 0
        ? validationResults.issues.map((issues) => `• ${issue}`).join("\n")
        : "No major issues detected."
    }


    View Dashboard: ${getWebAppUrl()}?page=dashboard
    `;

  GmailApp.sendEmail(userEmail, subject, body);
}

// ===== GET WEB APP

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
