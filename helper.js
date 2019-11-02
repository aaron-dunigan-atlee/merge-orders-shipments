/* Helper functions */

function compileHashedData(sheetObject, key) {
  // Returns the data from sheetName, hashed by key.
  // Key may not be unique per row; therefore each item is an array of objects that correspond to that value of key.
  var array = GsDb.getRows(sheetObject, {});
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
  var sheet = SpreadsheetApp.openById(ORDERS_SHEET_ID).getSheets()[0];
  return compileHashedData(sheet, 'orderKey');
}
function getShipmentsData() {
  var sheet = SpreadsheetApp.openById(SHIPMENTS_SHEET_ID).getSheets()[0];
  return compileHashedData(sheet, 'orderKey');
}

// Check whether an order and a shipment correspond.
function orderMatchesShipment(orderItem, shipmentItem) {
  return (orderItem.items_name == shipmentItem.shipmentItems_name 
          && orderItem.items_quantity == shipmentItem.shipmentItems_quantity
          && orderItem.orderNumber == shipmentItem.orderNumber);
}