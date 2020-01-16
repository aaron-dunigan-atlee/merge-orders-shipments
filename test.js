function testTime() {
  var start = new Date();
  populateMergedSheet();
  var end = new Date();
  Logger.log(end-start);
}

function testHeaders() {
  var headersArray = MERGED_SHEET.getDataRange().getValues().slice(0,2);
  Logger.log(headersArray);
  var headerProps = headersArray[0].filter(function(item, index) {
      return headersArray[1][index] == true;
    });
  Logger.log(headerProps);
}

function testPivot() {
  var pivotTable = SpreadsheetApp.getActive().getSheetByName('Report7').getPivotTables()[0];
  Logger.log("PivotValues");
  var pivotValues = pivotTable.getPivotValues();
  for (var i=0; i<pivotValues.length; i++) {
    Logger.log(pivotValues[i].getFormula());
  }
  Logger.log("RowGroups");
  var pivotGroups = pivotTable.getRowGroups();
  for (var i=0; i<pivotGroups.length; i++) {
    Logger.log(pivotGroups[i].getIndex());
  }
  var pivotFilters = pivotTable.getFilters();
  for (var i=0; i<pivotFilters.length; i++) {
    Logger.log(pivotFilters[i].getFilterCriteria().getVisibleValues());
  }
}

function testChart() {
  var pivotTableSheet = SpreadsheetApp.getActive().getSheetByName("Report15");
  var labelRange = pivotTableSheet.getRange("A2:A13");
  var valueRange = pivotTableSheet.getRange("C2:C13");
  var chartBuilder = pivotTableSheet.newChart();
  chartBuilder.addRange(labelRange)
      .addRange(valueRange)
      .setChartType(Charts.ChartType.PIE)
      .setOption('useFirstColumnAsDomain', true);
  pivotTableSheet.insertChart(chartBuilder.build());
}