function onOpen(e) {
  /*
   * Add a menu whenever the Spreadsheet is opened
   */
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("AppScript")
    .addItem("✉️ Send BizCard Sample", "sendBizcards")
    .addToUi();
}

function sendBizcards() {
  /*
   * Usage:
   * Gets the spreadsheet with the given id and sheet name. The sheet should have the following columns: Email, Subject, Message.
   * The sheet gets appended with a 'Status' column to avoid duplication of sending emails. The function loops through all the filled rows and sends email if the Status columns is blank.
   * After sending the email the corresponding Status column is filled with 'Email Sent'.
   */
  let { SHEET_ID } = initialize();
  clearLogs();
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);

  const responsesSheet = spreadsheet.getSheetByName("Form Responses 1");
  const responseValues = responsesSheet.getDataRange().getDisplayValues();
  const lastRow = responsesSheet.getLastRow();
  const columnHeaders = responsesSheet.getRange("A1:F1").getDisplayValues()[0];

  /*
   * Add a 'Status' column if it doesn't exist. This column will indicate if the email has been sent
   */

  if (!columnHeaders.includes("Status")) {
    var lastBlankColumnHeader = responsesSheet.getLastColumn() + 1;
    Logger.log("Adding Status column to avoid duplicates");
    responsesSheet.getRange(1, lastBlankColumnHeader).setValue("Status");
  }

  /*
   * Loop through the rows and get corersponding values
   */
  for (var i = 1; i < lastRow; i++) {
    var dataRow = responseValues[i];
    var name = dataRow[columnHeaders.indexOf("Name")];
    var email = dataRow[columnHeaders.indexOf("Email")];
    var mobileNumber = dataRow[columnHeaders.indexOf("Mobile Number")];
    var position = dataRow[columnHeaders.indexOf("Position")];
    var status = dataRow[columnHeaders.indexOf("Status")];

    if (status !== "Email sent") {
      /*
       * Create a copy of the slides template
       */
      let { SLIDES_ID } = initialize();
      Logger.log(`Generating bizcard for ${name}`);
      const bizcardTemplate = DriveApp.getFileById(SLIDES_ID);

      var fileName = `Bizcard for ${name}`;
      var bizcard = bizcardTemplate.makeCopy(fileName);
      var bizcardId = bizcard.getId();
      var slide = SlidesApp.openById(bizcardId).getSlides()[0];

      /*
       * Replace content of duplicated file with the values in the current row
       */
      slide.replaceAllText("<<NAME>>", name);
      slide.replaceAllText("<<EMAIL>>", email);
      slide.replaceAllText("<<MOBILE NUMBER>>", mobileNumber);
      slide.replaceAllText("<<POSITION>>", position);

      /*
       * Build email content and send email
       */
      var subject = "Your bizcard!";
      var message =
        "Hi " +
        name +
        "! \n Welcome to the company!\n\n" +
        "Please double check your bizcard before we print it out. Thank you!";
      var attachment = DriveApp.getFileById(bizcardId);
      SlidesApp.openById(bizcardId).saveAndClose();

      GmailApp.sendEmail(email, subject, message, {
        attachments: [attachment.getAs(MimeType.PDF)],
      });
      Logger.log(`Bizcard for ${name} is available at ${bizcardId}`);

      /*
       * Update status column with "Email sent"
       */
      responsesSheet
        .getRange(i + 1, responsesSheet.getLastColumn())
        .setValue("Email sent");

      Logger.log(`Sent ${name}'s bizcard.`);
    }
  }
}
