/* The main function: get all the data and merge it into one sheet */
function populateMergedSheet() {
  MERGED_DATA = [];
  ORDER_ROWS = [];
  var ordersData = getOrdersData();
  var shipmentsData = getShipmentsData();
  // Go through orders and process each orderNumber
  for (var orderNumber in ordersData) {
    // Procede if orderNumber is not blank.  We were picking up noise from empty rows.
    if (orderNumber) {
      var thisOrder = ordersData[orderNumber];
      // thisOrder is an array of order items.  
      // The first row in the spreadsheet will be order-level data: order number, name, address, etc.
      // The remaining rows will contain the data for each item in the order.
      for (var i=0; i<thisOrder.length; i++) {
        // Insert the order-level data just once.
        if (i==0) {
          // Find out how many items are in this order.
          var orderSize = thisOrder.length;
          insertMainEntry(thisOrder[0], orderSize);
        }
        // Insert each order item as a separate line
        // Create an array (filled with empty strings) where we can insert data,
        // by copying the EMPTY_ROW variable.
        var rowDataToAppend = EMPTY_ROW.slice();
        var rowNumber = getWorkingRowNumber();
        // Insert all the static values from the orders sheet.
        for (var columnNumber in mergedSheetItemOrderColumns) {
          columnNumber = parseInt(columnNumber, 10);
          var propertyName = mergedSheetItemOrderColumns[columnNumber];
          rowDataToAppend[columnNumber] = thisOrder[i][propertyName];
        }
        // Add the formulas for columns that don't contain static data.
        rowDataToAppend[ORDER_FULFILLED_COLUMN_INDEX] = " ";
        rowDataToAppend = addOrderDate(rowDataToAppend, thisOrder[i]); 
        rowDataToAppend = addItemShippedFormula(rowDataToAppend, thisOrder[i]); 
        rowDataToAppend = addItemTotalFormula(rowNumber, rowDataToAppend);
        rowDataToAppend = addShipmentData(rowDataToAppend, orderNumber, thisOrder[i], rowNumber, shipmentsData);
        // Row data is ready: write it to the array.
        MERGED_DATA.push(rowDataToAppend);
      }
      appendRowsForUnmatchedShipments(orderNumber, shipmentsData, thisOrder[0].orderDate.slice(0,10));
    }
  }
  // Clear existing data
  clearMergedDataFromSheet();
  // Set cell formats, then write to sheet.
  // MUST be done in this order: inserting checkboxes overwrites data.
  setMergedSheetFormats();
  writeMergedDataToSheet();
  setTimeStamp();
}

// Insert data for the main entry (date, customer name, address, etc.)
function insertMainEntry(orderObject, orderSize) {
  // Create a blank array where we can insert data.
  var rowDataToAppend = EMPTY_ROW.slice();
  var rowNumber = getWorkingRowNumber();
  // Insert all the static values from the orders sheet.
  for (var columnNumber in mergedSheetMainEntryOrderColumns) {
    columnNumber = parseInt(columnNumber, 10);
    var propertyName = mergedSheetMainEntryOrderColumns[columnNumber];
    rowDataToAppend[columnNumber] = orderObject[propertyName];
  }
  // Add the formulas for columns that don't contain static data.
  rowDataToAppend[ITEM_SHIPPED_COLUMN_INDEX] = " ";
  rowDataToAppend = addStoreNameFormula(rowNumber, rowDataToAppend);
  rowDataToAppend = addFulfilledFormula(rowNumber, rowDataToAppend, orderSize);
  rowDataToAppend = addOrderTotalFormula(rowNumber, rowDataToAppend, orderSize);
  rowDataToAppend = addOrderDate(rowDataToAppend, orderObject);
  // Row data is ready: write it to the array.
  MERGED_DATA.push(rowDataToAppend);
  // Add the range for this row to the list of main entry rows, for formatting purposes.
  var rowReference = rowNumber + ':' + rowNumber;
  MAIN_ENTRY_ROWS.push(rowReference);
}

// Clear all data from the merged sheet.
function clearMergedDataFromSheet() {
  // Get the range that encompasses current data.
  var dataRange = MERGED_SHEET.getDataRange();
  // Subtract header rows because they won't be cleared.
  var height = dataRange.getHeight() - MERGED_SHEET_HEADER_ROW_COUNT;
  // If there are no data rows, don't try to clear them.
  if (height > 0) {
    var width = dataRange.getWidth();
    // Remove checkboxes and clear all other data from the sheet.
    MERGED_SHEET.getRange(MERGED_SHEET_HEADER_ROW_COUNT + 1, 1, height, width).removeCheckboxes().clear();
  }
}

// Write the MERGED_DATA matrix to the sheet.
function writeMergedDataToSheet() {
  // Get height and width of MERGED_DATA array.
  var height = MERGED_DATA.length;
  var width = MERGED_SHEET_WIDTH;
  // Write data to sheet.
  MERGED_SHEET.getRange(MERGED_SHEET_HEADER_ROW_COUNT + 1, 1, height, width).setValues(MERGED_DATA);
}

// Set the formats for cells in the sheet.
// These formats must be set BEFORE the data is written,
// because .insertCheckboxes() sets the cell value to false.
function setMergedSheetFormats() {
  // Apply date format to whole date column.
  MERGED_SHEET.getRange(MERGED_SHEET_HEADER_ROW_COUNT + 1, ORDER_DATE_COLUMN_INDEX + COLUMN_INDEX_OFFSET, MERGED_DATA.length, 1).setNumberFormat('m/d/yyy');
  // Apply checkboxes to Fulfilled and Shipped columns.
  MERGED_SHEET.getRange(MERGED_SHEET_HEADER_ROW_COUNT + 1, ORDER_FULFILLED_COLUMN_INDEX + COLUMN_INDEX_OFFSET, MERGED_DATA.length, 1).insertCheckboxes();
  MERGED_SHEET.getRange(MERGED_SHEET_HEADER_ROW_COUNT + 1, ITEM_SHIPPED_COLUMN_INDEX + COLUMN_INDEX_OFFSET, MERGED_DATA.length, 1).insertCheckboxes();
  // Apply Shading to main entry rows.
  MERGED_SHEET.getRangeList(MAIN_ENTRY_ROWS).setBackground(SHADING_COLOR);
  // Don't extend order keys outside the cell.
  MERGED_SHEET.getRange(MERGED_SHEET_HEADER_ROW_COUNT + 1, ORDER_KEY_COLUMN_INDEX + COLUMN_INDEX_OFFSET, MERGED_DATA.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function setTimeStamp() {
  var now = new Date();
  var timeStamp = [["Last updated:", now]];
  MERGED_SHEET.getRange(1,1,1,2).setValues(timeStamp);
  MERGED_SHEET.getRange(1,2,1,1).setNumberFormat('m/d/yy h:mm');
}