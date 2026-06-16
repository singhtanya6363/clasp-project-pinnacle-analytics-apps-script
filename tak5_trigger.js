// ── PART B: Weekly Trigger ──
// Run this function ONCE manually — it sets up the alarm clock
// Run this function only once to register the trigger.
function setupWeeklyTrigger() {

  // Step 1: Check if a trigger for generateWeeklyReport already exists
  var existingTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === "generateWeeklyReport") {
      Logger.log("Trigger already exists — skipping creation.");
      return;  // Stop! Don't create a duplicate
    }
  }

  // Step 2: No existing trigger found — create a new one
  ScriptApp.newTrigger("generateWeeklyReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)   // 8 AM
    .create();

  Logger.log("Weekly trigger created: generateWeeklyReport runs every Monday at 8 AM.");
}