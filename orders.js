/*
 * Functions related to formatting data from the orders sheet. 
 */

/****************************
* Main-entry formulas       *
* (first row of each entry) *
*****************************/

// Construct an array for the first row in an order.
function constructOrderHeaderRow(rowObject, orderSize) {
  // Create a blank array where we can insert data.
  var rowDataToAppend = EMPTY_ROW.slice();
  // Fill in only the properties that are marked as header properties.
  for (var i=0; i<MAIN_ENTRY_PROPERTIES.length; i++) {
    var property = MAIN_ENTRY_PROPERTIES[i];
    columnNumber = getColumnIndex(property);
    if (columnNumber != -1 && rowObject.hasOwnProperty(property)) {
      rowDataToAppend[columnNumber] = rowObject[property];
    }
  }
  // Mark as header row
  rowDataToAppend[IS_HEADER_COLUMN_INDEX] = "TRUE";
  // Add the formulas for columns that don't contain static data.
  rowDataToAppend[ITEM_SHIPPED_COLUMN_INDEX] = " ";
  rowDataToAppend = addStoreNameFormula(rowDataToAppend);
  rowDataToAppend = addFulfilledFormula(rowDataToAppend, orderSize);
  rowDataToAppend = addOrderTotalFormula(rowDataToAppend, orderSize);
  rowDataToAppend = addOrderDate(rowDataToAppend, rowObject);
  // All set.
  return rowDataToAppend;
}

// Add a formula for the store name to a row of data.
function addStoreNameFormula(rowDataArray) {
  // Get the cell reference for the store ID.
  var columnOffset = STORE_ID_COLUMN_INDEX - STORE_NAME_COLUMN_INDEX;
  var storeIdCellReference = 'INDIRECT("R[0]C[' + columnOffset + ']", FALSE)'; 
  // Create a formula for looking up the corresponding store name.
  rowDataArray[STORE_NAME_COLUMN_INDEX] = '=VLOOKUP(' + storeIdCellReference + ',StoreNameLookup,2)';
  return rowDataArray;
}

function addFulfilledFormula(rowDataArray, orderSize) {
  // Set the Fulfilled checkbox to be TRUE if ALL of the items have shipped for this order.
  var columnOffset = ITEM_SHIPPED_COLUMN_INDEX - ORDER_FULFILLED_COLUMN_INDEX;
  var shippedCellsReference = 'INDIRECT("R[1]C[' + columnOffset + ']:R[' + orderSize + ']C[' + columnOffset + ']", FALSE)'; 
  var formula = '=COUNTIF(' + shippedCellsReference +',"FALSE") = 0';
  rowDataArray[ORDER_FULFILLED_COLUMN_INDEX] = formula;
  return rowDataArray;
}

function addOrderTotalFormula(rowDataArray, orderSize) {
  // Find sum of all item totals for this order.
  var columnOffset = ITEM_TOTAL_COLUMN_INDEX - ORDER_TOTAL_COLUMN_INDEX;
  var itemTotalsReference = 'INDIRECT("R[1]C[' + columnOffset + ']:R[' + orderSize + ']C[' + columnOffset + ']", FALSE)'; 
  var formula = '=SUM(' + itemTotalsReference +')';
  rowDataArray[ORDER_TOTAL_COLUMN_INDEX] = formula;
  return rowDataArray;
}

function addOrderDate(rowDataArray, orderObject) {
  // Remove the time data from the date field and add to the row data.
  var orderDate = orderObject.orders_orderDate;
  if (orderDate != undefined) {
    rowDataArray[ORDER_DATE_COLUMN_INDEX] = orderDate.slice(0,10); // Just get first 10 characters, e.g. '2019-04-10';
  }
  return rowDataArray;
}

/***************************
* Individual-item formulas *
****************************/

// Construct an array for an item row in an order.
function constructItemRow(rowObject) {
  // Convert to array.
  var rowDataToAppend = constructArrayFromObject(MERGED_SHEET_HEADERS, rowObject); 
  // Add the formulas for columns that don't contain static data.
  rowDataToAppend[ORDER_FULFILLED_COLUMN_INDEX] = " ";
  rowDataToAppend = addOrderDate(rowDataToAppend, rowObject); 
  rowDataToAppend = addItemShippedFormula(rowDataToAppend, rowObject); 
  rowDataToAppend = addItemTotalFormula(rowDataToAppend);
  rowDataToAppend = addStoreNameFormula(rowDataToAppend);
  // If there is a shipment present, add formulas.
  var orderItemId = rowObject.shipments_orderId
  if (orderItemId != undefined && orderItemId != '') {
    rowDataToAppend = addShipmentFormulas(rowDataToAppend, rowObject);
  }
  // Row data is ready: write it to the array.
  return rowDataToAppend;
}

function addItemTotalFormula(rowDataArray) {
  // Create formula to multiply item price by item quantity
  var columnOffset = ITEM_QTY_COLUMN_INDEX - ITEM_TOTAL_COLUMN_INDEX;
  var itemQtyCellReference = 'INDIRECT("R[0]C[' + columnOffset + ']", FALSE)'; 
  columnOffset = ITEM_PRICE_COLUMN_INDEX - ITEM_TOTAL_COLUMN_INDEX;
  var itemPriceCellReference = 'INDIRECT("R[0]C[' + columnOffset + ']", FALSE)';
  var formula = '=' + itemQtyCellReference + '*' + itemPriceCellReference;
  rowDataArray[ITEM_TOTAL_COLUMN_INDEX] = formula;
  return rowDataArray;
}

function addItemShippedFormula(rowDataArray, orderObject) {
  if (orderObject.orders_orderStatus == 'shipped'){
    var shipped = 'TRUE';
  } else {
    var shipped = 'FALSE';
  }
  rowDataArray[ITEM_SHIPPED_COLUMN_INDEX] = shipped;
  return rowDataArray;
}

