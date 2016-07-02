// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT";

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 7)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var name = row[0];          // 1st column
    var emailAddress = row[1];  // 2nd column
    var message = row[5];       // 6th column

    var html =
    '<body>' +
      '<p>YOUR HTML WITH INLINE CSS STYLES HERE</p>'
    '</body>'

    var emailSent = row[6];     // 7th column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Subject";
        MailApp.sendEmail(
          emailAddress,         // recipient
          subject,              // subject
          'body', {             // body
            htmlBody: html,     // advanced options
            // bcc: "",        // blind carbon copy
            //replyTo: "" // reply to
          }
        );
      sheet.getRange(startRow + i, 7).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
