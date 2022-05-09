//author: aidan chang
//interpreted from https://developers.google.com/apps-script/articles/sending_emails
//indexing starts at 1
//sends emails with data from the current spreadsheet.


// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 24; // Number of rows to process  (to row 25)
  var startColumn = 1; // First column of data to process
  var numColumns = 15; // Number of columns to process  (cols A-M)
  // Fetch the range of cells A2:O25
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var recNum = parseInt(row[0]); //col A

    var recRow = data[recNum - 2];

    var emailAddress = row[5]; // Col F
    var emailSent = row[15]; // Col O
    var subject = 'Secret Santa!';

    //constructing message
    //developed from: https://stackoverflow.com/questions/10720832/line-break-in-a-message
    var name = row[3]; // col D
    var message = [];

    message.push("Hey " + name + ",___LINE_BREAK___");
    message.push("Glad you are a part of the 2021 secret santa!___LINE_BREAK___A few reminders:___LINE_BREAK___1. We are going to be exchanging gifts the morning of Saturday, December 4th before we go to Six Flags___LINE_BREAK___2. Keep the total gift cost < $20___LINE_BREAK___3. This isn't a white elephant gift exchange, everyone is assigned a specific person to give a gift to___LINE_BREAK______LINE_BREAK___<u>Here is to whom you are giving:</u>");
    message.push(row[2] + "___LINE_BREAK___");

    message.push("<u>Their favorites</u>");
    if (recRow[6] != "") {message.push(recRow[6]);}
    if (recRow[7] != "") {message.push("book / movie / tv show: " + recRow[7]);}
    if (recRow[8] != "") {message.push("music artist: " + recRow[8]);}
    if (recRow[9] != "") {message.push("type of food: " + recRow[9]);}
    if (recRow[10] != "") {message.push("restaurant / place: " + recRow[10]);}
    if (recRow[11] != "") {message.push("clothing brand / store: " + recRow[11]);}
    if (recRow[12] != "") {message.push("snack / drink: " + recRow[12]);}

    message.push("___LINE_BREAK___--___LINE_BREAK___Thank You,___LINE_BREAK___Aidan Chang");

    // Combine content into a single string
    var preFormatContent = message.join('___LINE_BREAK___');

    // Replace text with \n for plain text
    var plainTextContent = preFormatContent.replace(/___LINE_BREAK___/g, '\n');
    // Replace text with <br /> for HTML
    var htmlContent = preFormatContent.replace(/___LINE_BREAK___/g, '<br />');


    //var emailSent = row[14]; // col O
    if (emailSent !== EMAIL_SENT && name == "Kevin") { // Prevents sending duplicates

      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: htmlContent
      });
      
      sheet.getRange((startRow + i), 15).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
    else {
      sheet.getRange((startRow + i), 15).setValue("0"); //email failed
    }
  }
}

//limit is 500 emails per day? assuming one recipient per email
function getEmailQuota() {
  console.log(MailApp.getRemainingDailyQuota());
}
