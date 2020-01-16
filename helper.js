/* Helper functions */

function compileHashedGsdbData(sheetObject, key, prefix) {
  // Returns the data from sheetObject, hashed by key.
  // Key may not be unique per row; therefore each item is an array of objects that correspond to that value of key.
  var array = GsDb.getRows(sheetObject, {}, prefix);
  var hashedData = array.reduce(function(accumulator, object, index) {
    var propertyName = object[key];
    if (accumulator[propertyName] == undefined) {
      accumulator[propertyName] = [object];
    } else {
      accumulator[propertyName] = accumulator[propertyName].concat(object);
    }
    return accumulator;
  }, {});
  return hashedData;
}

/**
 * Hash an array by the given unique key.
 * @param {Array} array 
 * @param {string} key 
 */
function hashArray(array, key) {
  var hashedData = array.reduce(function(accumulator, object, index) {
    var propertyName = object[key];
      accumulator[propertyName] = object;
    return accumulator;
  }, {});
  return hashedData;
}

// Create an array of specified length, filled with fillItem.
function filledArray(arrayLength, fillItem) {
  var arr = [];
    for(var i=0; i<arrayLength; i++){
        arr.push(fillItem);
    }
    return arr;
} 

// Get the *Spreadsheet* row number of the next row to be appended, 
// based on the current length of MERGED_DATA
function getWorkingRowNumber() {
  // The next row number (in the spreadsheet) is the length (number of rows) of the current array
  // plus the number of header rows, plus one more.
  return MERGED_DATA.length + MERGED_SHEET_HEADER_ROW_COUNT + 1;
}

// Get all of the orders and shipments as objects, referenced by their orderKey. 
// Note that throughout the code, the variable 'orderNumber' is really an orderKey, because
// we changed keys part way through the project.
function getOrdersData() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  return compileHashedGsdbData(sheet, 'orders_orderKey', 'orders_');
}

function getShipmentsData() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Shipments');
  return compileHashedGsdbData(sheet, 'shipments_orderKey', 'shipments_');
}

// Check whether an order and a shipment correspond.
function orderMatchesShipment(orderItem, shipmentItem) {
  return (orderItem.items_name == shipmentItem.shipmentItems_name 
          && orderItem.items_quantity == shipmentItem.shipmentItems_quantity
          && orderItem.orderNumber == shipmentItem.orderNumber);
}

// In a list of objects, determine if the key: value pair exists in at least one.
function valueFoundInObjectList(key, value, array) {
  for (var i=0; i<array.length; i++) {
    if (array[i][key] === value) {
      return true;
    }
  }
  return false;
}

function setTimeStamp() {
  var now = new Date();
  var timeStamp = [["Last updated:", now]];
  SpreadsheetApp.getActive().getRangeByName('Timestamp')
    .setValues(timeStamp)
    .setNumberFormat('m/d/yy h:mm');
}

function getOrderHeaderFields() {}

/* Get the (0-based) column index for a given property
 * in the merged sheet.
 */
function getColumnIndex(property) {
  return MERGED_SHEET_HEADERS.indexOf(property);
}

// Construct an array whose values come from rowObject, 
// and correspond with the property names in headers.
function constructArrayFromObject(headers, rowObject) {
  var arrayData = [];
  for (var i=0; i<headers.length; i++) {
    if (rowObject.hasOwnProperty(headers[i])) {
      arrayData.push(rowObject[headers[i]]);
    } else {
      arrayData.push(""); 
    }
  }
  return arrayData;
}