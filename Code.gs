/**
 * Automated Job Application & Follow-Up Tracker
 *
 * Public demonstration version.
 * Replace placeholder configuration values before running.
 */

function sendJobEmails() {

  // ===== SETTINGS =====
  const DAILY_LIMIT = 50;

  // Replace with your own Google Drive PDF resume file ID.
  const RESUME_FILE_ID = "YOUR_RESUME_FILE_ID";

  // Open current Google Sheet
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data =
    sheet.getDataRange().getValues();

  // ===== GET RESUME =====
  const resumeFile =
    DriveApp.getFileById(RESUME_FILE_ID);

  const resumePDF =
    resumeFile
      .getBlob()
      .getAs(MimeType.PDF)
      .setName("Resume.pdf");

  let sentCount = 0;

  // ===== PROCESS APPLICATIONS =====
  for (let i = 1; i < data.length; i++) {

    // Stop after daily limit
    if (sentCount >= DAILY_LIMIT) {
      break;
    }

    // ===== SHEET DATA =====
    const name = data[i][1];       // B = Name
    const email = data[i][2];      // C = Email
    const company = data[i][4];    // E = Company
    const result = data[i][7];     // H = Result

    // Skip empty email
    if (!email) {
      continue;
    }

    // ===== DUPLICATE EMAIL CHECK =====
    const currentEmail =
      String(email).trim().toLowerCase();

    let duplicate = false;

    for (let d = 1; d < i; d++) {

      const previousEmail =
        String(data[d][2]).trim().toLowerCase();

      if (previousEmail === currentEmail) {
        duplicate = true;
        break;
      }
    }

    // Skip duplicate email
    if (duplicate) {
      continue;
    }

    // Skip already sent applications
    if (
      String(result).trim().toLowerCase() === "sent"
    ) {
      continue;
    }

    // ===== PERSONALIZATION =====
    const recipientName =
      name || "Hiring Manager";

    const companyName =
      company || "your organization";

    // ===== EMAIL SUBJECT =====
    const subject =
      "Application for Opportunities at " +
      companyName;

    // ===== EMAIL BODY =====
    const body =
      "Dear " + recipientName + ",\n\n" +

      "I hope you are doing well.\n\n" +

      "I am writing to express my sincere interest in " +
      "exploring suitable career opportunities at " +
      companyName + ".\n\n" +

      "I am an enthusiastic MCA graduate with a strong " +
      "interest in Python development, Data Analytics, SQL, " +
      "REST APIs, Full Stack Development, AI technologies, " +
      "and workflow automation.\n\n" +

      "As a fresher, I am looking for an opportunity where " +
      "I can apply my technical knowledge, gain practical " +
      "industry experience, and contribute meaningfully " +
      "to the organization.\n\n" +

      "I would be grateful if you could consider my profile " +
      "for any suitable entry-level, fresher, internship, " +
      "or relevant opportunity available at " +
      companyName + ".\n\n" +

      "I have attached my resume in PDF format for your " +
      "kind consideration.\n\n" +

      "Thank you for your time and consideration. " +
      "I look forward to hearing from you.\n\n" +

      "Warm regards,\n" +
      "Your Name\n" +
      "Your Location\n" +
      "your.email@example.com";

    try {

      // ===== SEND EMAIL =====
      GmailApp.sendEmail(
        email,
        subject,
        body,
        {
          attachments: [resumePDF],
          name: "Your Name"
        }
      );

      // ===== SUCCESS =====

      // F = Sent Date
      sheet
        .getRange(i + 1, 6)
        .setValue(new Date());

      // G = Reply
      sheet
        .getRange(i + 1, 7)
        .setValue("No Reply");

      // H = Result
      sheet
        .getRange(i + 1, 8)
        .setValue("Sent");

      // Clear previous error information
      // M = Error Type
      // N = Error Details
      // O = Error Date
      sheet
        .getRange(i + 1, 13, 1, 3)
        .clearContent();

      sentCount++;

    } catch (error) {

      // ===== ERROR HANDLING =====

      const errorMessage =
        String(error.message || "");

      const lowerError =
        errorMessage.toLowerCase();

      let errorType =
        "Unknown Error";

      // ===== ERROR CLASSIFICATION =====

      if (
        lowerError.includes("quota") ||
        lowerError.includes("limit") ||
        lowerError.includes("too many")
      ) {

        errorType =
          "Quota Exceeded";

      } else if (
        lowerError.includes("permission") ||
        lowerError.includes("authorization") ||
        lowerError.includes("access denied")
      ) {

        errorType =
          "Permission Error";

      } else if (
        lowerError.includes("invalid") ||
        lowerError.includes("recipient") ||
        lowerError.includes("address")
      ) {

        errorType =
          "Invalid Email";
      }

      // ===== UPDATE RESULT =====

      // H = Result
      sheet
        .getRange(i + 1, 8)
        .setValue("Error");

      // M = Error Type
      sheet
        .getRange(i + 1, 13)
        .setValue(errorType);

      // N = Error Details
      sheet
        .getRange(i + 1, 14)
        .setValue(errorMessage);

      // O = Error Date
      sheet
        .getRange(i + 1, 15)
        .setValue(new Date());

      // Stop if quota is exceeded
      if (errorType === "Quota Exceeded") {
        break;
      }
    }
  }

  // ===== COMPLETION MESSAGE =====
  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      sentCount +
      " email(s) sent successfully.",
      "Job Email Automation",
      5
    );
}
