/*
 * Functions related to formatting data from the orders sheet. 
 */

/****************************
* Main-entry formulas       *
* (first row of each entry) *
*****************************/

// Add a formula for the store name to a row of data.
function addStoreNameFormula(rowNumber, rowDataArray) {
  // Get the cell reference, e.g. H3, for the store ID.
  var storeIdCellReference = STORE_ID_COLUMN_LETTER + rowNumber;
  // Create a formula for looking up the corresponding store name.
  rowDataArray[STORE_NAME_COLUMN_INDEX] = '=VLOOKUP(' + storeIdCellReference + ',StoreNameLookup,2)';
  return rowDataArray;
}

function addFulfilledFormula(rowNumber, rowDataArray, orderSize) {
  // Set the Fulfilled checkbox to be TRUE if ALL of the items have shipped for this order.
  var shippedCellsReference = MERGED_SHEET.getRange(rowNumber + 1, ITEM_SHIPPED_COLUMN_INDEX + COLUMN_INDEX_OFFSET, orderSize, 1).getA1Notation();
  var formula = '=COUNTIF(' + shippedCellsReference +',"FALSE") = 0';
  rowDataArray[ORDER_FULFILLED_COLUMN_INDEX] = formula;
  return rowDataArray;
}

function addOrderTotalFormula(rowNumber, rowDataArray, orderSize) {
  // Find sum of all item totals for this order.
  var itemTotalsReference = MERGED_SHEET.getRange(rowNumber + 1, ITEM_TOTAL_COLUMN_INDEX + COLUMN_INDEX_OFFSET, orderSize, 1).getA1Notation();
  var formula = '=SUM(' + itemTotalsReference +')';
  rowDataArray[ORDER_TOTAL_COLUMN_INDEX] = formula;
  return rowDataArray;
}

function addOrderDate(rowDataArray, orderObject) {
  // Remove the time data from the date field and add to the row data.
  var orderDate = orderObject.orderDate.slice(0,10); // Just get first 10 characters, e.g. '2019-04-10'
  rowDataArray[ORDER_DATE_COLUMN_INDEX] = orderDate;
  return rowDataArray;
}

/***************************
* Individual-item formulas *
****************************/

function addItemTotalFormula(rowNumber, rowDataArray) {
  // Create formula to multiply item price by item quantity
  var itemQtyCellReference = ITEM_QTY_COLUMN_LETTER + rowNumber;
  var itemPriceCellReference = ITEM_PRICE_COLUMN_LETTER + rowNumber;
  var formula = '=' + itemQtyCellReference + '*' + itemPriceCellReference;
  rowDataArray[ITEM_TOTAL_COLUMN_INDEX] = formula;
  return rowDataArray;
}

function addItemShippedFormula(rowDataArray, orderObject) {
  if (orderObject.orderStatus == 'shipped'){
    var shipped = 'TRUE';
  } else {
    var shipped = 'FALSE';
  }
  rowDataArray[ITEM_SHIPPED_COLUMN_INDEX] = shipped;
  return rowDataArray;
}

