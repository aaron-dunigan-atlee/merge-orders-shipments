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
  console.time('getOrdersData')
  var sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  var hash = compileHashedGsdbData(sheet, 'orders_orderKey', 'orders_');
  console.timeEnd('getOrdersData')
  return hash
}

function getShipmentsData() {
  console.time('getShipmentsData')
  var sheet = SpreadsheetApp.getActive().getSheetByName('Shipments');
  var hash = compileHashedGsdbData(sheet, 'shipments_orderKey', 'shipments_');
  console.timeEnd('getShipmentsData')
  return hash
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

function leftJoin(leftObjects, rightObjects, leftKey, rightKey) {
  var hashRightObjects = hashArr(rightObjects, rightKey);
  return leftObjects.map(function(lObj) {
    var rObj = hashRightObjects[lObj[leftKey]];
    for (var property in rObj) {
      lObj[property] = rObj[property]
    }
    return lObj;
  });
}

function fillRow(dbObjects, filterObject, rowObject) {
  // Update a single row in the data.  Cells with prior values will not be replaced.
  // Values will be taken from rowObject, 
  // where attribute names match column headers.  Attributes that are
  // not in the headers will be ignored.  Rows to be changed will be 
  // determined by filterObject, which should be one of the headers/attributes
  // If a matching row is not found, a message will be logged.

  var offsetHeaderRow = 1;
  var offsetZeroBasedIndex = 1;
  
  var foundMatchingRow = false;
  // Iterate from end for more efficiency.
  for (var i = dbObjects.length - 1; i >= 0; i--) {
    if (objectMatchesFilterObject(dbObjects[i], filterObject)) { 
      foundMatchingRow = true; 
      for (var property in rowObject) {
        if (dbObjects[i][property] == undefined || dbObjects[i][property] == '') {
          dbObjects[i][property] = rowObject[property]
        }
      }
      break;
    }
  }
  if (!foundMatchingRow) {
    var message = "Could not find row in merged sheet where " + JSON.stringify(filterObject);
    console.log(message);
  } 
}

function objectMatchesFilterObject(object, filterObject) {
  for (var property in filterObject) {
    if (filterObject[property].indexOf(object[property]) == -1) { return false; }
  }
  return true;
}