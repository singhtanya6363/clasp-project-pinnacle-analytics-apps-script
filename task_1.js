// Practice CI/CD Test - 14 July 2026
function flagProjectStatuses() {

  // Step 1: Open the spreadsheet and go to the Projects tab and start doing work
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Projects");

  // Step 2: Find the last row that has data (so we don't hardcode 17)
  var lastRow = sheet.getLastRow();

  // Step 3: Read ALL the data at once (faster than reading cell by cell)
  // getRange(startRow, startCol, numRows, numCols)
  // Row 2 = skip header, Col 1 = Project ID, we need 6 columns (A to F)
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // Step 4: Get today's date
  var today = new Date();

  // Step 5: Go through every row one by one
  for (var i = 0; i < data.length; i++) {

    var dueDate  = data[i][4];  // Column E = index 4 = Due Date
    var hours    = data[i][7];  // Column H = index 7 = Hours Logged
    var status   = "";          // We'll fill this in

    // Step 6: Check if anything is blank
    if (dueDate === "" || hours === "") {
      status = "Incomplete Data";

    // Step 7: Due date is in the PAST
    } else if (dueDate < today) {
      if (hours == 0) {
        status = "Cancelled";
      } else {
        status = "Complete";
      }

    // Step 8: Due date is within 14 days from today
    } else if ((dueDate - today) / (1000 * 60 * 60 * 24) <= 14) {
      status = "Due Soon";

    // Step 9: Due date is more than 14 days away
    } else {
      status = "On Track";
    }

    // Step 10: Write the status into column F (column index 6, which is col 6)
    sheet.getRange(i + 2, 6).setValue(status);
  }
}