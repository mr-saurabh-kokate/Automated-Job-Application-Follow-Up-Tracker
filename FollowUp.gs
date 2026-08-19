/**
 * Weekly Follow-Up Automation
 *
 * Sends a follow-up email for applications that:
 * - Were successfully sent
 * - Have not received a reply
 * - Are at least 7 days old
 * - Have not received a follow-up within the last 7 days
 *
 * Sheet columns:
 * F = Sent Date
 * G = Reply
 * H = Result
 * K = Follow-up Count
 * L = Last Follow-up
 */

function sendWeeklyFollowUps() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data = sheet.getDataRange().getValues();

  const today = new Date();

  let followUpCount = 0;

  // Start from row 2
  for (let i = 1; i < data.length; i++) {

    // ===== SHEET DATA =====
    const name = data[i][1];          // B = Name
    const email = data[i][2];         // C = Email
    const company = data[i][4];       // E = Company
    const sentDate = data[i][5];      // F = Sent Date
    const reply = data[i][6];         // G = Reply
    const result = data[i][7];        // H = Result
    const followUpCount = data[i][10]; // K = Follow-up Count
    const lastFollowUp = data[i][11];  // L = Last Follow-up

    // Skip incomplete records
    if (!email || !sentDate) {
      continue;
    }

    // Only follow up on successfully sent applications
    if (String(result).trim() !== "Sent") {
      continue;
    }

    // Do not follow up if recruiter already replied
    if (
      String(reply).trim().toLowerCase() === "replied"
    ) {
      continue;
    }

    // ===== CHECK ORIGINAL EMAIL AGE =====
    const sent = new Date(sentDate);

    const daysSinceSent =
      Math.floor(
        (today - sent) /
        (1000 * 60 * 60 * 24)
      );

    // First follow-up only after 7 days
    if (daysSinceSent < 7) {
      continue;
    }

    // ===== CHECK LAST FOLLOW-UP =====
    if (lastFollowUp) {

      const previousFollowUp =
        new Date(lastFollowUp);

      const daysSinceFollowUp =
        Math.floor(
          (today - previousFollowUp) /
          (1000 * 60 * 60 * 24)
        );

      // Wait another 7 days before next follow-up
      if (daysSinceFollowUp < 7) {
        continue;
      }
    }

    // ===== PERSONALIZATION =====
    const recipientName =
      name || "Hiring Manager";

    const companyName =
      company || "your organization";

    // ===== EMAIL SUBJECT =====
    const subject =
      "Follow-Up: Application for Opportunities at " +
      companyName;

    // ===== EMAIL BODY =====
    const body =
      "Dear " + recipientName + ",\n\n" +

      "I hope you are doing well.\n\n" +

      "I am writing to follow up on my previous email regarding "
      + "potential career opportunities at " +
      companyName + ".\n\n" +

      "I remain very interested in exploring suitable "
      + "entry-level or fresher opportunities with your "
      + "organization. I would be grateful if you could "
      + "consider my profile for any relevant openings.\n\n" +

      "Please let me know if any additional information "
      + "is required from my side.\n\n" +

      "Thank you for your time and consideration. "
      + "I look forward to hearing from you.\n\n" +

      "Warm regards,\n" +
      "Your Name\n" +
      "Your Location\n" +
      "your.email@example.com";

    try {

      // ===== SEND FOLLOW-UP =====
      GmailApp.sendEmail(
        email,
        subject,
        body,
        {
          name: "Your Name"
        }
      );

      // ===== UPDATE FOLLOW-UP COUNT =====
      const currentCount =
        Number(followUpCount) || 0;

      sheet
        .getRange(i + 1, 11)
        .setValue(currentCount + 1);

      // ===== UPDATE LAST FOLLOW-UP =====
      sheet
        .getRange(i + 1, 12)
        .setValue(new Date());

      followUpCount++;

    } catch (error) {

      console.error(
        "Follow-up failed for " +
        email +
        ": " +
        error.message
      );
    }
  }

  // ===== COMPLETION MESSAGE =====
  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      followUpCount +
      " follow-up email(s) sent.",
      "Weekly Follow-Up",
      5
    );
}
