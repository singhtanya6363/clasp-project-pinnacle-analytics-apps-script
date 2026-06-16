function generateWeeklyReport() {

  // Step 0: Run Task 1 first so Status column is filled
  flagProjectStatuses();

  // Step 1: Get all sheets needed
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var projectsSheet = ss.getSheetByName("Projects");
  var hoursSheet    = ss.getSheetByName("Team Hours");
  var reportSheet   = ss.getSheetByName("Weekly Report");

  // Step 2: Clear everything in Weekly Report before writing fresh
  reportSheet.clearContents();
  reportSheet.clearFormats();

  // Step 3: Read Projects data
  // Projects tab has title in row 1, headers in row 2, data from row 3
  var projLastRow = projectsSheet.getLastRow();
  var projData    = projectsSheet.getRange(3, 1, projLastRow - 2, 8).getValues();
  // [0]=Project ID, [1]=Client, [2]=Project Name, [3]=Start, [4]=Due Date, [5]=Status, [6]=Budget, [7]=Hours

  // Step 4: Read Team Hours data
  // Team Hours tab has title in row 1, headers in row 2, data from row 3
  var hoursLastRow = hoursSheet.getLastRow();
  var hoursData    = hoursSheet.getRange(3, 1, hoursLastRow - 2, 4).getValues();
  // [0]=Analyst, [1]=Project ID, [2]=Week Ending, [3]=Hours Worked

  // Step 5: Get today's date formatted nicely
  var today      = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM dd, yyyy");

  // ═══════════════════════════════════════
  // ROW 1 — Title (bold, merged)
  // ═══════════════════════════════════════
  var titleRange = reportSheet.getRange(1, 1, 1, 5);  // merge across 5 columns
  titleRange.merge();
  titleRange.setValue("Weekly Operations Report — " + dateString);
  titleRange.setFontWeight("bold");
  titleRange.setFontSize(14);
  titleRange.setBackground("#3d3c63");
  titleRange.setFontColor("#ffffff");
  titleRange.setHorizontalAlignment("center");

  // ROW 2 — intentionally empty (gap)

  // ═══════════════════════════════════════
  // ROW 3 — Summary (total + count by status)
  // ═══════════════════════════════════════

  // Count how many projects are in each status
  var counts = { "Complete": 0, "Cancelled": 0, "Due Soon": 0, "On Track": 0, "Incomplete Data": 0 };
  for (var i = 0; i < projData.length; i++) {
    var s = projData[i][5];  // Column F = Status
    if (counts[s] !== undefined) counts[s]++;
  }

  // Write summary as one line in row 3
  var summaryText = "Total Projects: " + projData.length +
                    "   |   Complete: "        + counts["Complete"] +
                    "   |   Due Soon: "         + counts["Due Soon"] +
                    "   |   On Track: "         + counts["On Track"] +
                    "   |   Cancelled: "        + counts["Cancelled"] +
                    "   |   Incomplete Data: "  + counts["Incomplete Data"];

  var summaryRange = reportSheet.getRange(3, 1, 1, 5);
  summaryRange.merge();
  summaryRange.setValue(summaryText);
  summaryRange.setFontWeight("bold");
  summaryRange.setBackground("#e8f0fe");

  // ROW 4 — intentionally empty (gap)

  // ═══════════════════════════════════════
  // ROW 5 — Total hours logged
  // ═══════════════════════════════════════

  // Add up all hours from Team Hours tab
  var totalHours = 0;
  for (var j = 0; j < hoursData.length; j++) {
    totalHours += hoursData[j][3];  // Column D = Hours Worked
  }

  var hoursRange = reportSheet.getRange(5, 1, 1, 5);
  hoursRange.merge();
  hoursRange.setValue("Total Hours Logged Across All Analysts: " + totalHours);
  hoursRange.setFontWeight("bold");
  hoursRange.setBackground("#e8f0fe");

  // ROW 6 — intentionally empty (gap)

  // ═══════════════════════════════════════
  // ROW 7+ — Due Soon projects table
  // ═══════════════════════════════════════

  // Section heading at row 7
  var dueSoonTitle = reportSheet.getRange(7, 1, 1, 4);
  dueSoonTitle.merge();
  dueSoonTitle.setValue("DUE SOON PROJECTS");
  dueSoonTitle.setFontWeight("bold");
  dueSoonTitle.setBackground("#fff3cd");

  // Table header at row 8
  var tableHeader = reportSheet.getRange(8, 1, 1, 4);
  tableHeader.setValues([["Project ID", "Client", "Due Date", "Hours Logged"]]);
  tableHeader.setFontWeight("bold");
  tableHeader.setBackground("#d9d9d9");

  // Write each Due Soon project starting at row 9
  var currentRow = 9;
  for (var k = 0; k < projData.length; k++) {
    if (projData[k][5] === "Due Soon") {
      var dueDateFormatted = Utilities.formatDate(
        new Date(projData[k][4]),
        Session.getScriptTimeZone(),
        "MMM dd, yyyy"
      );
      reportSheet.getRange(currentRow, 1, 1, 4).setValues([
        [projData[k][0], projData[k][1], dueDateFormatted, projData[k][7]]
      ]);
      currentRow++;
    }
  }

  // ═══════════════════════════════════════
  // AFTER Due Soon table — Cancelled projects
  // ═══════════════════════════════════════

  // One empty gap row after the Due Soon table
  currentRow++;

  // Cancelled section heading
  var cancelledTitle = reportSheet.getRange(currentRow, 1, 1, 4);
  cancelledTitle.merge();
  cancelledTitle.setValue("CANCELLED PROJECTS — Flagged for Follow-Up");
  cancelledTitle.setFontWeight("bold");
  cancelledTitle.setBackground("#f8d7da");
  currentRow++;

  // Table header for cancelled
  var cancelledHeader = reportSheet.getRange(currentRow, 1, 1, 3);
  cancelledHeader.setValues([["Project ID", "Client", "Project Name"]]);
  cancelledHeader.setFontWeight("bold");
  cancelledHeader.setBackground("#d9d9d9");
  currentRow++;

  // Write each Cancelled project
  for (var m = 0; m < projData.length; m++) {
    if (projData[m][5] === "Cancelled") {
      reportSheet.getRange(currentRow, 1, 1, 3).setValues([
        [projData[m][0], projData[m][1], projData[m][2]]
      ]);
      currentRow++;
    }
  }
}