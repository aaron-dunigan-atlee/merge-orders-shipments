// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
/*
 * @param {sheet} sheet with data to be pulled from.
 * @param {range} range where the data is in the sheet, headers are above;  
 * @param {row} 
 */


function getRowsData(sheet, range, columnHeadersRowIndex, displayValues, getBlanks) {
  displayValues = displayValues || false;
  getBlanks = getBlanks || false;
  if (sheet.getLastRow() < 2){
    return [];
  }
  var headersIndex = columnHeadersRowIndex || (range ? range.getRowIndex() - 1 : 1);
  var dataRange = range ||
    sheet.getRange(headersIndex+1, 1, sheet.getLastRow() - headersIndex, sheet.getLastColumn());
  var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  // NOTE THAT WE ARE NOT NORMALIZING HEADERS HERE.  HEADERS ARE ALREADY FORMATTED IN THE SHEET.
  if (displayValues == true){
    return getObjects_(dataRange.getDisplayValues(), headers, getBlanks);
  } else {
    return getObjects_(dataRange.getValues(), headers, getBlanks);
  }
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys, getBlanks) {
  var objects = [];
  var timeZone = Session.getScriptTimeZone();

  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        if (getBlanks){
          object[keys[j]] = '';
        }
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}


// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  // NOTE THAT WE ARE NOT NORMALIZING HEADERS HERE.  HEADERS ARE ALREADY FORMATTED IN THE SHEET.
  var headers = headersRange.getValues()[0];

  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
    }
    data.push(values);
  }

  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(), 
                                        objects.length, headers.length);
  destinationRange.setValues(data);
  return destinationRange
}

/**
 * Get the range object corresponding to the data range minus a header row.
 * @param {Sheet} sheet 
 */
function getDataRangeMinusHeaders(sheet, optHeaderRowCount) {
  var headerRowCount = optHeaderRowCount || 1;
  var dataRange = sheet.getDataRange();
  var height = dataRange.getHeight();
  if (header <= headerRowCount) {
    return null;
  }
  var width = dataRange.getWidth();
  return sheet.getRange(1, 1, height-headerRowCount, width);
}

/**
 * Fill one row of data per object defined in the objects Array,
 * starting after any existing data.
 */
function appendRowsData(sheet, objects, optHeadersRange) {
  var firstDataRowIndex = sheet.getDataRange().getLastRow() + 1;
  setRowsData(sheet, objects, optHeadersRange, firstDataRowIndex);
}

/**
 * Clear sheet and then write data.
 * @param {sheet} sheet 
 * @param {rows data} data 
 */
function writeNewData(sheet, data) {
  var lastRow = sheet.getLastRow()
  if (lastRow > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear()
  }
  setRowsData(sheet, data)
}


/**
 * Set rows data, but preserve any formulas in the sheet, rather than 
 * overwriting with values.
 */
function setRowsDataKeepFormulas(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  // NOTE THAT WE ARE NOT NORMALIZING HEADERS HERE.  HEADERS ARE ALREADY FORMATTED IN THE SHEET.
  var headers = headersRange.getValues()[0];

  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(), 
                                        objects.length, headers.length);
  var formulas = destinationRange.getFormulas();
  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      if (formulas[i][j]) {
        values.push(formulas[i][j])
      } else {
        var header = headers[j];
        values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
      }
    }
    data.push(values);
  }

  
  destinationRange.setValues(data);
  return destinationRange
}