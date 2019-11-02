// If any trigger exists, cancel it.
function cancelAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  // Loop over all triggers and delete them.
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

// Add an hourly trigger to run the merge.
function setHourlyTrigger() {
  // First cancel any existing triggers.
  cancelAllTriggers();
  var trigger = ScriptApp.newTrigger('populateMergedSheet')
    .timeBased()
    .everyHours(1)
    .create();
  // Inform the user.
  SpreadsheetApp.getUi().alert('Automatic updates set to once an hour.');
}

function cancelTriggersAndAlert() {
  cancelAllTriggers();
  // Inform the user.
  SpreadsheetApp.getUi().alert('Automatic updates canceled.');
}