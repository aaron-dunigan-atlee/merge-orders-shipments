/* The main function: get all the data and merge it into one sheet */
function populateMergedSheet() {
  initializeGlobals();
  // Go through shipments and add to existing orders where needed.
  for (var orderNumber in shipmentsData) {
    // Proceed if orderNumber is not blank.  We were picking up noise from empty rows.
    if (orderNumber) {
      if (existingMergedData[orderNumber] != undefined) {
        updateShipmentsInMerged(orderNumber);
      }
    }
  }
  // Go through orders and add new ones to the pending additions.
  for (var orderNumber in ordersData) {
    // Proceed if orderNumber is not blank.  We were picking up noise from empty rows.
    if (orderNumber) {
      // If this order is not already in the merged sheet, 
      // add it to pending changes.
      if (existingMergedData[orderNumber] == undefined) {
        addOrderToMergedUpdate(orderNumber); 
      } 
    }
  }
  // Write the new orders to the sheet.
  MergeDb.addRows(MERGED_SHEET, mergedDataToAdd);
}

// Add data if it's not part of the existing spreadsheet.
function addOrderToMergedUpdate(orderNumber) {
  // Join orders and shipments.
  var joinedOrder = MergeDb.leftJoin(ordersData[orderNumber],shipmentsData[orderNumber],'orders_items_orderItemId','shipments_shipmentItems_orderItemId');
  mergedDataToAdd[orderNumber] = joinedOrder;
}

function updateShipmentsInMerged(orderNumber) {
  // existingMergedObject will be an array of row objects.
  var existingMergedObject = existingMergedData[orderNumber];
  // Iterate through shipments with this orderKey.
  for (var i=0; i<shipmentsData[orderNumber].length; i++) {
    var shipmentObject = shipmentsData[orderNumber][i];
    // If this shipment is not yet in the spreadsheet, add it.
    // Keying by shipmentItems_orderItemId.  Very rarely, there are duplicates.  
    var shipmentItemId = shipmentObject['shipments_shipmentItems_orderItemId']; 
    if (!hasShipmentData(existingMergedObject, shipmentItemId)) {
      // Add this shipment to the spreadsheet under the correct order. 
      var filter = {'orders_items_orderItemId': shipmentObject['shipments_shipmentItems_orderItemId']};
      MergeDb.fillRow(MERGED_SHEET_NAME, filter, shipmentObject);
    }
  }
}

function initializeGlobals() {
  // Initialize global variables.
  mergedDataToUpdate = {};
  mergedDataToAdd = {};
  ordersData = getOrdersData();
  shipmentsData = getShipmentsData();
  existingMergedData = MergeDb.getJson(MERGED_SHEET);
}

function writeUpdatesToMergedSheet() {
  // MergeDb.setRows(mergedDataToUpdate);
  MergeDb.addRows(mergedDataToAdd);
}

function hasShipmentData(orderArray, orderItemId) {
  // 1. Find item in orderArray that has items_orderItemId equal to orderItemId
  // 2. Check if this item already has a value for shipmentItems_orderItemId.
  for (var i=0; i<orderArray.length; i++) {
    if (orderArray[i].orders_items_orderItemId == orderItemId) {
      if (orderArray[i].shipments_shipmentItems_orderItemId == orderItemId) {
        return true;
      } else {
        return false;
      }
    }
  }
  return false;
}

/* For now, we're not using this.
function updateOrdersInMerged(orderNumber) {
  var existingMergedObject = existingMergedData[orderNumber];
  for (var i=0; i<ordersData[orderNumber].length; i++) {
    var orderObject = ordersData[orderNumber][i];
    // If this specific line of the order is not found, 
    // add it to the order and then add that order to list
    // of orders to update.
    var key = 'items_orderItemId';
    if (!valueFoundInObjectList(key, orderObject[key], existingMergedObject.orders)) {
      existingMergedObject.orders.push(orderObject);
      mergedDataToUpdate[orderNumber] = existingMergedObject;
    }
  }
}
*/

// Set the formats for cells in the sheet, starting at a given row and for a given height.
// These formats must be set BEFORE the data is written,
// because .insertCheckboxes() sets the cell value to false.
function setMergedSheetFormats(row, height) {
  // Apply date format to whole date column.
  MERGED_SHEET.getRange(row + 1, ORDER_DATE_COLUMN_INDEX + COLUMN_INDEX_OFFSET, height, 1).setNumberFormat('m/d/yyy');
  // Apply checkboxes to Fulfilled and Shipped columns.
  MERGED_SHEET.getRange(row + 1, ORDER_FULFILLED_COLUMN_INDEX + COLUMN_INDEX_OFFSET, height, 1).insertCheckboxes();
  MERGED_SHEET.getRange(row + 1, ITEM_SHIPPED_COLUMN_INDEX + COLUMN_INDEX_OFFSET, height, 1).insertCheckboxes();
  // Apply Shading to main entry rows.
  // ????????????
  MERGED_SHEET.getRangeList(MAIN_ENTRY_ROWS).setBackground(SHADING_COLOR);
  // Don't extend order keys outside the cell.
  MERGED_SHEET.getRange(row + 1, ORDER_KEY_COLUMN_INDEX + COLUMN_INDEX_OFFSET, height, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}