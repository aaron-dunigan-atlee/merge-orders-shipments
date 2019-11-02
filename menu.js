// Set up a menu to allow us to run the merge command at any time.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Merge Orders and Shipments')
      .addItem('Update now', 'mergeMenuItem')
      .addItem('Set schedule to hourly', 'setHourlyTrigger')
      .addItem('Turn off scheduled update', 'cancelTriggersAndAlert')
      .addToUi();
}

function mergeMenuItem() {
  populateMergedSheet();
}