function aggregateAnalystHours() {

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hoursSheet  = ss.getSheetByName("Team Hours");
  var reportSheet = ss.getSheetByName("Weekly Report Analyst");

  // Read all Team Hours data (skip header)
  var lastRow = hoursSheet.getLastRow();
  var data    = hoursSheet.getRange(3, 1, lastRow - 1, 4).getValues();

  // Object to accumulate totals per analyst
  var totals = {};

  for (var i = 0; i < data.length; i++) {
    var analyst = data[i][0];  // Column A
    var project = data[i][1];  // Column B
    var hours   = data[i][3];  // Column D

    if (analyst === "" || analyst === null) continue;

    // Create entry if analyst seen for first time
    if (!totals[analyst]) {
      totals[analyst] = { totalHours: 0, projects: [] };
    }

    // Add hours
    totals[analyst].totalHours += hours;

    // Add project only if not already in the list
    if (project !== "" && totals[analyst].projects.indexOf(project) === -1) {
      totals[analyst].projects.push(project);
    }
  }

  // Build sortable array
  var summary = [];
  for (var name in totals) {
    var numProjects = totals[name].projects.length;
    var avgHours    = numProjects > 0 ? Math.round((totals[name].totalHours / numProjects) * 100) / 100 : 0;
    summary.push([name, totals[name].totalHours, numProjects, avgHours]);
  }

  // Sort by Total Hours — highest first
  summary.sort(function(a, b) { return b[1] - a[1]; });

  // Clear Weekly Report tab
  reportSheet.clearContents();
  reportSheet.clearFormats();

  // Write bold header at row 1
  var headerRange = reportSheet.getRange(1, 1, 1, 4);
  headerRange.setValues([["Analyst Name", "Total Hours", "Number of Projects", "Avg Hours/Project"]]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#d9d9d9");

  // Write data from row 2
  if (summary.length > 0) {
    reportSheet.getRange(2, 1, summary.length, 4).setValues(summary);
  }

  // BONUS: Colour Total Hours cell by volume
  for (var r = 0; r < summary.length; r++) {
    var hoursVal = summary[r][1];
    var cell     = reportSheet.getRange(r + 2, 2);
    if (hoursVal >= 100)     { cell.setBackground("#ff33ec"); }  // Red
    else if (hoursVal >= 70) { cell.setBackground("#ffff00"); }  // Yellow
    else                     { cell.setBackground("#00cc44"); }  // Green
  }
}