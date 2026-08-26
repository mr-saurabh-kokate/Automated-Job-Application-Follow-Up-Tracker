/**
 * Reply Tracker
 *
 * Checks Gmail for replies from contacts listed in the sheet
 * and updates reply status, date, details, and result.
 */

function checkReplies() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data =
    sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const email = data[i][2];       // C = Email
    const result = data[i][7];      // H = Result
    const reply = data[i][6];       // G = Reply

    const normalizedResult =
      String(result).trim().toLowerCase();

    const normalizedReply =
      String(reply).trim().toLowerCase();

    // Only check successfully sent applications
    if (
      !email ||
      normalizedResult !== "sent" ||
      normalizedReply === "replied"
    ) {
      continue;
    }

    // Search Gmail for messages from the contact
    const query =
      "from:" + String(email).trim();

    const threads =
      GmailApp.search(query);

    if (threads.length === 0) {
      continue;
    }

    let replyFound = false;

    for (let j = 0; j < threads.length; j++) {

      if (replyFound) {
        break;
      }

      const messages =
        threads[j].getMessages();

      for (let k = 0; k < messages.length; k++) {

        const message =
          messages[k];

        const sender =
          message.getFrom().toLowerCase();

        const trackedEmail =
          String(email).trim().toLowerCase();

        // Confirm sender matches tracked email
        if (
          sender.indexOf(trackedEmail) !== -1
        ) {

          // ===== REPLY STATUS =====
          sheet
            .getRange(i + 1, 7)
            .setValue("Replied");

          // ===== REPLY DATE =====
          sheet
            .getRange(i + 1, 9)
            .setValue(message.getDate());

          // ===== REPLY DETAILS =====
          let replyText =
            message.getPlainBody();

          // Keep sheet content manageable
          if (replyText.length > 500) {
            replyText =
              replyText.substring(0, 500) + "...";
          }

          sheet
            .getRange(i + 1, 10)
            .setValue(replyText);

          // ===== RESULT =====
          sheet
            .getRange(i + 1, 8)
            .setValue("Reply Received");

          replyFound = true;
          break;
        }
      }
    }
  }

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      "Reply check completed.",
      "Job Application Tracker",
      5
    );
}
