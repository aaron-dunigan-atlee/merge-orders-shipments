/**
 * Create pivot table for reports.
 * Adapted from https://sites.google.com/site/scriptsexamples/learn-by-example/google-sheets-api/pivot
 */
function createPivotTable() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  MERGED_SHEET_TITLES = ss.getRangeByName('OrdersAndShipmentsHeaders').getValues()[0];
  MERGED_SHEET_HEADERS = ss.getDataRange().getValues()[0];
  var reportValueKey = ss.getRangeByName('ReportColumnKey').getValue();
  var reportValueColumn = getColumnByHeader(reportValueKey);
  var reportGroupKey = ss.getRangeByName('ReportGroupByKey').getValue();
  var reportGroupColumn = getColumnByHeader(reportGroupKey);
  var reportStartDate = ss.getRangeByName('ReportStartDate').getValue();
  var reportEndDate = ss.getRangeByName('ReportEndDate').getValue();
  var reportFilterKey1 = ss.getRangeByName('FilterKey1').getValue();
  var reportFilterKey2 = ss.getRangeByName('FilterKey2').getValue();
  var reportFilterKey3 = ss.getRangeByName('FilterKey3').getValue();
  var reportFilterValue1 = ss.getRangeByName('FilterValue1').getValue();
  var reportFilterValue2 = ss.getRangeByName('FilterValue2').getValue();
  var reportFilterValue3 = ss.getRangeByName('FilterValue3').getValue();
  var includeCount = ss.getRangeByName('IncludeCount').getValue();
  var includeSum = ss.getRangeByName('IncludeSum').getValue();
  var includePercent = ss.getRangeByName('IncludePercent').getValue();
  var includePieChart = ss.getRangeByName('IncludePieChart').getValue();
  var includeBarChart = ss.getRangeByName('IncludeBarChart').getValue();
  var includeQuarterlyChart = ss.getRangeByName('IncludeQuarterlyChart').getValue();
  
  // The name of the sheet containing the data you want to put in a table.
  var sheetName = "Orders and Shipments";
  var sourceSheet = ss.getSheetByName(sheetName);
  // Use display values to match the pivot table.
  var sourceData = sourceSheet.getDataRange().getDisplayValues();
  MERGED_SHEET_HEADERS = sourceData.shift();
  for (var i=1; i<MERGED_SHEET_HEADER_ROW_COUNT; i++) {
    sourceData.shift();
  }
  var pivotTableParams = {};
  
  // The source indicates the range of data you want to put in the table.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  pivotTableParams.source = {
    sheetId: sourceSheet.getSheetId(),
    startRowIndex: MERGED_SHEET_HEADER_ROW_COUNT - 1,
    endRowIndex: sourceSheet.getLastRow()
  };
  
  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = [{
    sourceColumnOffset: reportGroupColumn, 
    sortOrder: "ASCENDING",
    showTotals: true
  }];

  // Defines how a value in a pivot table should be calculated.
  pivotTableParams.values = [];
  if (includeCount) {
    pivotTableParams.values.push({
      summarizeFunction: "COUNTA",
      sourceColumnOffset: reportValueColumn,
      name: "Count"
    });
  }
  if (includeSum) {
    pivotTableParams.values.push({
      summarizeFunction: "SUM",
      sourceColumnOffset: reportValueColumn,
      name: "Total"
    });
  }
  if (includePercent) {
    pivotTableParams.values.push({
      summarizeFunction: "SUM",
      sourceColumnOffset: reportValueColumn,
      calculatedDisplayType: "PERCENT_OF_COLUMN_TOTAL",
      name: "Percent"
    });
  }
  
  // Set filters
  // No zeroes or blanks.
  var reportGroupUniqueValues = getUniqueNonemptyValues(sourceData, reportGroupColumn);
  var reportValueUniqueValues = getUniqueNonemptyValues(sourceData, reportValueColumn);
  pivotTableParams.criteria = {};
  pivotTableParams.criteria[reportGroupColumn] = {'visibleValues': reportGroupUniqueValues};
  pivotTableParams.criteria[reportValueColumn] = {'visibleValues': reportValueUniqueValues};
  
  // Date filter
  var startDate = new Date(reportStartDate);
  startDate.setHours(0,0,0,0);
  var endDate = new Date(reportEndDate);
  endDate.setHours(0,0,0,0);
  var dateColumn = getColumnIndex('orders_orderDate');
  var dateColumnFilterValues = getUniqueNonemptyValues(sourceData, dateColumn).filter(function(value){
    var dateValue = new Date(value);
    dateValue.setHours(0,0,0,0);
    return (startDate <= dateValue && dateValue <= endDate);
  });
  pivotTableParams.criteria[dateColumn] = {'visibleValues': dateColumnFilterValues};

  // Custom filters from sheet
  var filterHeaders = []
  var filterKeys = [reportFilterKey1, reportFilterKey2, reportFilterKey3];
  var filterValues = [reportFilterValue1, reportFilterValue2, reportFilterValue3];
  for (var i=0; i<3; i++) {
    if (filterKeys[i] != '' && filterKeys[i] != undefined) {
      var filterColumn = getColumnByHeader(filterKeys[i]);
      pivotTableParams.criteria[filterColumn] = {'visibleValues': [filterValues[i]]};
      filterHeaders.push(filterKeys[i] + " is " + filterValues[i]);
    }
  }

  // Create a new sheet which will contain our Pivot Table
  var documentProperties = PropertiesService.getDocumentProperties();
  var reportNumber = documentProperties.getProperty("REPORT_NUMBER");
  if (reportNumber == undefined) {
    reportNumber = '0';
  }
  reportNumber = (parseInt(reportNumber, 10) + 1) % 1000;
  documentProperties.setProperty("REPORT_NUMBER", reportNumber);
  var pivotTableSheetName = "Report" + reportNumber;
  var pivotTableSheet = ss.insertSheet(pivotTableSheetName);
  var pivotTableSheetId = pivotTableSheet.getSheetId();
  
  // Add header rows
  var headerStartDate = ss.getRangeByName('ReportStartDate').getDisplayValue();
  var headerEndDate = ss.getRangeByName('ReportEndDate').getDisplayValue();
  var pivotHeader = "Report for dates " + headerStartDate + " - " + headerEndDate;
  pivotTableSheet.appendRow([pivotHeader]);
  for (var i=0; i<filterHeaders.length; i++) {
    pivotTableSheet.appendRow([filterHeaders[i]]);
  }
  var pivotHeaderRowCount = 1 + filterHeaders.length;
  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": pivotTableSheetId,
        "rowIndex": pivotHeaderRowCount,
        "columnIndex": 0
      },
      "fields": "pivotTable"
    }
  };

  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, ss.getId());
  // Set cell wrapping
  pivotTableSheet.getDataRange().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  pivotTableSheet.getRange(1,1,pivotHeaderRowCount,1).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  


  var pivotTable = pivotTableSheet.getPivotTables()[0];
  var pivotDataStartRow = pivotTable.getAnchorCell().getRow() + 1;
  var pivotChartLabelColumn = pivotTable.getAnchorCell().getColumn();
  var pivotDataHeight = pivotTableSheet.getLastRow() - pivotDataStartRow;
  var pivotChartValueColumn = pivotChartLabelColumn + includeCount + includeSum;
  var pivotDataLastColumn = pivotChartLabelColumn + includeCount + includeSum + includePercent;
  // Add pie chart.
  if (includePieChart) {
    var labelRange = pivotTableSheet.getRange(pivotDataStartRow, pivotChartLabelColumn, pivotDataHeight, 1);
    var valueRange = pivotTableSheet.getRange(pivotDataStartRow, pivotChartValueColumn, pivotDataHeight, 1);
    var chartBuilder = pivotTableSheet.newChart();
    chartBuilder
      .setChartType(Charts.ChartType.PIE)
      .addRange(labelRange)
      .addRange(valueRange)
      .setPosition(1,pivotDataLastColumn+1,0,0)
      .setOption('useFirstColumnAsDomain', true);
    pivotTableSheet.insertChart(chartBuilder.build());
  }
  // Add pie chart.
  if (includeBarChart) {
    var labelRange = pivotTableSheet.getRange(pivotDataStartRow, pivotChartLabelColumn, pivotDataHeight, 1);
    var valueRange = pivotTableSheet.getRange(pivotDataStartRow, pivotChartValueColumn, pivotDataHeight, 1);
    var chartBuilder = pivotTableSheet.newChart();
    chartBuilder
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(labelRange)
      .addRange(valueRange)
      .setPosition(1,pivotDataLastColumn+1,0,0)
      .setOption('useFirstColumnAsDomain', true);
    pivotTableSheet.insertChart(chartBuilder.build());
  }
}
/**
 * Get unique values in column, as strings, that are not 0 or blank.
 * @param {Array[][]} matrix 
 * @param {number} column 
 */
function getUniqueNonemptyValues(matrix, column) {
  var uniqueHash = {};
  var uniqueValues = [];
  for (var row=0; row<matrix.length; row++) {
    var item = matrix[row][column];
    if (uniqueHash[item] == undefined) {
      uniqueHash[item] = true;
      if (item != 0 && item != '' && item != '0' && item != undefined) {
        uniqueValues.push(item.toString());
      }
    }
  }
  return uniqueValues;
}

/**
 * Get column index (0-based) from a header name 
 * (not property name, but natural-language header in last row of headers).
 */
function getColumnByHeader(header) {
  return MERGED_SHEET_TITLES.indexOf(header);
}

function clearFilter1() {
  SpreadsheetApp.getActive().getRangeByName('FilterKey1').setValue('');
  SpreadsheetApp.getActive().getRangeByName('FilterValue1').setValue('');
}

function clearFilter2() {
  SpreadsheetApp.getActive().getRangeByName('FilterKey2').setValue('');
  SpreadsheetApp.getActive().getRangeByName('FilterValue2').setValue('');
}

function clearFilter3() {
  SpreadsheetApp.getActive().getRangeByName('FilterKey3').setValue('');
  SpreadsheetApp.getActive().getRangeByName('FilterValue3').setValue('');
}