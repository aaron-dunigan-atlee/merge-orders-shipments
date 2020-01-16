/*
Database functions for the merged sheet.
Interacts with rows as objects using the headers in row 1.
They should be of the form orders_propertyName or shipments_propertyName,
or merged_propertyName.
*/

var MergeDb = (function () {

  function getRows(sheetObject, filterObject) {
    var objects = getObjects(sheetObject);
    return objects.filter(function(object) {
      return objectMatchesFilterObject(object, filterObject)
    });
  }
  
  function addRows(sheetName, hashedObjects) {
    // Add rows to the sheet.  Values will be taken from objects, 
    // where attribute names match column headers.  Attributes that are
    // not in the headers will be ignored.
    var sheet = getSheet(sheetName);
    var matrixToAppend = [];
    for (var orderKey in hashedObjects) {
      var orderItemArray = hashedObjects[orderKey];
      matrixToAppend.push(constructOrderHeaderRow(orderItemArray[0], orderItemArray.length));
      for (var i=0; i<orderItemArray.length; i++) {
        matrixToAppend.push(constructItemRow(orderItemArray[i]));
      }
    }
    var firstBlankRow = sheet.getLastRow() + 1;
    // Set formats and write to sheet.
    setMergedSheetFormats(firstBlankRow, matrixToAppend.length);
    appendMatrix(sheet, firstBlankRow, matrixToAppend);
  }

  // Append a matrix of data at the row indicated.
  function appendMatrix(sheetObject, row, matrix) {
    var height = matrix.length;
    var width = matrix[0].length;
    sheetObject.getRange(row, 1, height, width).setValues(matrix);
  }
  
  function fillRow(sheetName, filterObject, rowObject) {
    // Update a single row in the sheet.  Cells with prior values will not be replaced.
    // Values will be taken from rowObject, 
    // where attribute names match column headers.  Attributes that are
    // not in the headers will be ignored.  Rows to be changed will be 
    // determined by filterObject, which should be one of the headers/attributes
    // If a matching row is not found, a message will be logged.
    
    var sheet = getSheet(sheetName);
    var dbObjects = getObjects(sheet);
    var offsetHeaderRow = 1;
    var offsetZeroBasedIndex = 1;

    var foundMatchingRow = false;
    // Iterate from end for more efficiency.
    for (var i = dbObjects.length - 1; i >= 0; i--) {
      if (objectMatchesFilterObject(dbObjects[i], filterObject)) { 
        foundMatchingRow = true; 
        var rowToUpdate = i + MERGED_SHEET_HEADER_ROW_COUNT + ROW_INDEX_OFFSET;
        var rowRange = sheet.getRange(rowToUpdate, 1, 1, sheet.getLastColumn());
        var rowValues = rowRange.getValues()[0];
        var rowFormulas = rowRange.getFormulas()[0];
        for (var column=1; column<=rowValues.length; column++) {
          var columnIndex = column - COLUMN_INDEX_OFFSET;
          if (rowValues[columnIndex] == '' && rowFormulas[columnIndex] == '') {
            var property = MERGED_SHEET_HEADERS[columnIndex];
            if (rowObject.hasOwnProperty(property)) {
              // Set one cell at a time.  Inefficient, but it can't be avoided.
              sheet.getRange(rowToUpdate, column).setValue(rowObject[property]);
            }
          }
        }
      }
    }
    if (!foundMatchingRow) {
      var message = "Could not find row in merged sheet where " + JSON.stringify(filterObject);
      Logger.log(message);
    } 
  }
  
  function delRows() {  }
  
  
  
  function getObjects(sheetObject) {
    var values = getValues(sheetObject);
    // Get headers for property names.
    var headers = values.shift();
    // Remove additional header rows.
    for (var i=1; i<MERGED_SHEET_HEADER_ROW_COUNT; i++) {values.shift()}
    // Create an object for each row.
    var objects = [];
    for (var row = 0; row < values.length; row++) {
      var object = {};
      for (var col = 0; col < headers.length; col++) {
        var propertyName = headers[col];
        object[propertyName] = values[row][col];
      }
      objects.push(object);
    }
    return objects;
  }
  
  function objectMatchesFilterObject(object, filterObject) {
    for (var property in filterObject) {
      if (filterObject[property].indexOf(object[property]) == -1) { return false; }
    }
    return true;
  }
  
  function getSheet(sheetName) {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }

  function getValues(sheetObject) {
    return sheetObject.getDataRange().getValues();
  }
 
  function getHeaders(sheetName) {
    // Return a string array of the headers as they appear on the sheet.
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    return sheet.getDataRange().getValues()[0];
  }

  function getMainEntryProperties(sheetObject) {
    var headersArray = sheetObject.getDataRange().getValues().slice(0,2);
    return headersArray[0].filter(function(item, index) {
      return headersArray[1][index] == true;
    });
  }

  /* Get data from merged sheet and organize it. */
  function getJson(sheetObject) {
    // Returns the data from sheetObject, hashed into the json structure 
    // specified in the specs.
    var array = MergeDb.getRows(sheetObject, {});
    var hashedData = array.reduce(function(accumulator, object, index) {
      var orderKey = object['orders_orderKey'];
      if (accumulator[orderKey] == undefined) {
        // Template for desired structure.
        accumulator[orderKey] = [object];
      } else {
        accumulator[orderKey].push(object);
      }
      return accumulator;
    }, {});
    return hashedData;
  }

  return {
    getRows: getRows,
    addRows: addRows,
    fillRow: fillRow,
    getHeaders: getHeaders,
    getJson: getJson,
    getMainEntryProperties: getMainEntryProperties
  }
})();

