function sendDueSoonAlert() {

  // Step 1: Set the recipient (stored as a constant — not buried in the code)
  var RECIPIENT = "tanya@goalkeep.net";

  // Step 2: Get the projects sheet and read data
  var projectsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Projects");
  var emailLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Log");
  var projData      = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 8).getValues();

  // Step 3: Build a list of Due Soon projects
  var dueSoonProjects = [];
  for (var i = 0; i < projData.length; i++) {
    if (projData[i][5] === "Due Soon") {
      dueSoonProjects.push(projData[i]);
    }
  }

  // Step 4: Get today's date
  var today      = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM dd, yyyy");

  // Step 5: Build the email subject and body
  var subject = "[Pinnacle Analytics] Weekly Due Soon Alert — " + dateString;
  var status  = "";

  // Step 6: If NO projects are due soon — skip the email, just log it
  if (dueSoonProjects.length === 0) {
    logEmail(emailLogSheet, today, RECIPIENT, subject, 0, "Skipped — 0 projects due soon");
    return;  // Stop the function here
  }

  // Step 7: Build the email body
  var body = "Hello,\n\nThe following projects are Due Soon as of " + dateString + ":\n\n";
  for (var j = 0; j < dueSoonProjects.length; j++) {
    var p        = dueSoonProjects[j];
    var dueDate  = Utilities.formatDate(new Date(p[4]), Session.getScriptTimeZone(), "MMM dd, yyyy");
    body        += "• " + p[0] + " | " + p[1] + " | Due: " + dueDate + " | Hours Logged: " + p[7] + "\n";
  }
  body += "\nPlease review and take action as needed.\n\nPinnacle Analytics — Automated Alert";

  // Step 8: Try to send the email (wrapped in try/catch for safety)
  try {
    GmailApp.sendEmail(RECIPIENT, subject, body);
    status = "Sent";
  } catch (e) {
    status = "Error: " + e.message;
  }

  // Step 9: Log it to the Email Log tab
  logEmail(emailLogSheet, today, RECIPIENT, subject, dueSoonProjects.length, status);
}

// Helper function to write a row to the Email Log tab
function logEmail(sheet, timestamp, recipient, subject, numProjects, status) {
  var lastRow    = sheet.getLastRow();
  var timeString = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // If headers aren't there yet, add them
  if (lastRow < 3) {
    sheet.getRange(3, 1, 1, 6).setValues([["Timestamp", "Recipient", "Subject", "Projects Flagged", "Status", "Error (if any)"]]);
    sheet.getRange(3, 1, 1, 6).setFontWeight("bold");
    lastRow = 3;
  }

  sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[timeString, recipient, subject, numProjects, status]]);
}