/**
 * Automated Job Application & Follow-Up Tracker
 *
 * Public demonstration version.
 * Replace placeholder configuration values before running.
 *
 * Features:
 * - Personalized job application emails
 * - Duplicate email protection
 * - Already sent application protection
 * - Error logging
 * - Invalid email / recipient rejection detection
 * - Permission error detection
 * - Gmail quota handling
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

  // ===== START FROM ROW 2 =====
  for (let i = 1; i < data.length; i++) {

    // Stop after daily limit
    if (sentCount >= DAILY_LIMIT) {
      break;
    }

    // ===== SHEET DATA =====
    const name = data[i][1];       // B = Name
    const email = data[i][2];      // C = Email
    const title = data[i][3];      // D = Title
    const company = data[i][4];    // E = Company
    const result = data[i][7];     // H = Result

    // ===== SKIP EMPTY EMAIL =====
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

    // ===== SKIP ALREADY PROCESSED RECORDS =====
    const currentResult =
      String(result).trim().toLowerCase();

    // Sent = already sent successfully
    // Error = previously failed, do not retry automatically
    if (
      currentResult === "sent" ||
      currentResult === "error"
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
      companyName + ". I am a motivated and enthusiastic " +
      "MCA graduate eager to begin the next stage of my " +
      "career and contribute to a dynamic organization.\n\n" +

      "I have a strong interest in Python development, " +
      "Data Analytics, SQL, REST APIs, Full Stack Development, " +
      "AI technologies, and workflow automation. During my " +
      "academic and project work, I have worked on practical " +
      "projects involving Python, Django, SQL, data analysis, " +
      "REST APIs, and automation.\n\n" +

      "As a fresher, I am looking for an opportunity where I " +
      "can apply my technical knowledge, gain practical " +
      "industry experience, and contribute meaningfully to " +
      "the organization.\n\n" +

      "I would be grateful if you could consider my profile " +
      "for any suitable entry-level, fresher, internship, " +
      "or relevant opportunity available at " +
      companyName + ".\n\n" +

      "I have attached my resume in PDF format for your kind " +
      "consideration.\n\n" +

      "Thank you for your time and consideration. " +
      "I look forward to hearing from you.\n\n" +

      "Warm regards,\n" +
      "Your Name\n" +
      "Your Location\n" +
      "your.email@example.com";

    // ===== SEND EMAIL =====
    try {

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

      // ===== CLASSIFY ERROR =====

      if (
        lowerError.includes("quota") ||
        lowerError.includes("limit") ||
        lowerError.includes("too many")
      ) {

        errorType =
          "Quota Exceeded";

      } else if (
        lowerError.includes("550 5.4.1") ||
        lowerError.includes("recipient address rejected") ||
        lowerError.includes("recipient rejected")
      ) {

        errorType =
          "Recipient Rejected";

      } else if (
        lowerError.includes("invalid") ||
        lowerError.includes("address") ||
        lowerError.includes("malformed")
      ) {

        errorType =
          "Invalid Email";

      } else if (
        lowerError.includes("permission") ||
        lowerError.includes("authorization") ||
        lowerError.includes("access denied")
      ) {

        errorType =
          "Permission Error";
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

      // ===== STOP ONLY IF QUOTA IS EXCEEDED =====
      //
      // Other email errors will not stop the script.
      // The script will continue with the next row.

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
