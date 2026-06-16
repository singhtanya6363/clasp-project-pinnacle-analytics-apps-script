// ── PART A: Custom Menu ──
// This runs AUTOMATICALLY whenever anyone opens the spreadsheet
function onOpen() {
  var ui   = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Pinnacle Ops");  // Creates the menu

  // Add each button with a label and the function it calls
  menu.addItem("🔄 Update Project Statuses", "flagProjectStatuses");
  menu.addItem("📊 Aggregate Analyst Hours",  "aggregateAnalystHours");
  menu.addItem("📄 Generate Weekly Report",   "generateWeeklyReport");
  menu.addItem("📧 Send Due Soon Alert",      "sendDueSoonAlert");
  menu.addSeparator();  // Draws a line ——
  menu.addItem("ℹ️ About This Tool",          "showAbout");

  menu.addToUi();  // Actually adds the menu to the screen
}

// ── About popup ──
function showAbout() {
  SpreadsheetApp.getUi().alert(
    "Pinnacle Analytics Automation Tool\n" +
    "Built by: Tanya\n" +
    "Date Built: March 2026\n\n" +
    "This tool automates project status flagging, hours aggregation, weekly reporting, and email alerts."
  );
}

