/*
 * Functions related to inserting and formatting data from the shipments sheet. 
 */

// For a given order number and item in that order, find the corresponding data from the 
// Shipments tab and add it to the row data.
function addShipmentData(rowDataToAppend, orderNumber, orderItem, rowNumber, shipmentsData) {
  var thisShipment = shipmentsData[orderNumber];
  // If there is no shipment with this order number, skip the next part.
  if (thisShipment) {
    // Loop through each of the shipment items under this order number
    for (var i=0; i<thisShipment.length; i++) {    
      var shipmentItem = thisShipment[i];
      // If this shipment item has already been matched with an order item, don't duplicate it.
      // Otherwise, find the item it matches with.  
      if (!shipmentItem.matched && orderMatchesShipment(orderItem, shipmentItem)) {
        // This shipment item that matches the order item, so mark it as matched.
        shipmentItem.matched = true;
        // Insert all the static values from the shipments sheet.
        for (var columnNumber in mergedSheetItemShipmentColumns) {
          columnNumber = parseInt(columnNumber, 10);
          var propertyName = mergedSheetItemShipmentColumns[columnNumber];
          rowDataToAppend[columnNumber] = thisShipment[i][propertyName];
        }
        // Add the formulas for columns that don't contain static data.
        rowDataToAppend = addDimensions(rowDataToAppend, thisShipment[i]);
        rowDataToAppend = addWeight(rowDataToAppend, thisShipment[i]);
        rowDataToAppend = addCarrierService(rowDataToAppend, thisShipment[i]);   
        rowDataToAppend = addQuarterFormula(rowNumber, rowDataToAppend); 
        // Break the for-loop, because we only want one matched shipment per row.
        break;
      }
    }
  }
  return rowDataToAppend;
}

// If there are any shipments for this order number that are not marked as .matched
// (i.e. they have not been matched up with an order item),
// append them as separate rows.  This should really only happen if an individual item has been 
// broken into multiple shipment installments.  (e.g. order for 8 widgets was shipped as 4 and 4)
function appendRowsForUnmatchedShipments(orderNumber, shipmentsData, orderDate) {
  var thisShipment = shipmentsData[orderNumber];
  // If there is no shipment with this order number, skip the next part.
  if (thisShipment) {
    appendRowsForShipmentObject(thisShipment, orderDate);
  }
  /* This part no longer needed since we are now keying by orderKey
  // Sometimes (always?) installments are given a related order number.
  // e.g. order xxxxx becomes xxxxx-1, xxxxx-2, etc.
  // This next section checks for these and appends them as well.
  var suffixCount = 1;
  var suffixedOrderNumber = orderNumber + '-' + suffixCount;
  while (shipmentsData[suffixedOrderNumber] != undefined) {
    // The suffix number existed, so append a row:
    appendRowsForShipmentObject(shipmentsData[suffixedOrderNumber]);
    // Increment the suffix number and check again
    suffixCount += 1;
    suffixedOrderNumber = orderNumber + '-' + suffixCount;
  }
  */
}

// Add a row for a shipment object (when it hasn't been matched to an order).
function appendRowsForShipmentObject(shipmentObject, orderDate) {
  for (var i=0; i<shipmentObject.length; i++) {
    // Check if this shipment item is not matched.  If it is, just skip this.
    if (!shipmentObject[i].matched) {
      // Insert each item as a separate line
      // Create a blank array where we can insert data.
      var rowDataToAppend = EMPTY_ROW.slice();
      var rowNumber = getWorkingRowNumber();
      // Insert all the static values from the orders sheet.
      for (var columnNumber in mergedSheetItemShipmentColumns) {
        columnNumber = parseInt(columnNumber, 10);
        var propertyName = mergedSheetItemShipmentColumns[columnNumber];
        rowDataToAppend[columnNumber] = shipmentObject[i][propertyName];
      }
      // Add the formulas for columns that don't contain static data.
      rowDataToAppend = addDimensions(rowDataToAppend, shipmentObject[i]);
      rowDataToAppend = addWeight(rowDataToAppend, shipmentObject[i]);
      rowDataToAppend = addCarrierService(rowDataToAppend, shipmentObject[i]);  
      rowDataToAppend = addQuarterFormula(rowNumber, rowDataToAppend);
      rowDataToAppend = setOrderValuesForUnmatchedShipment(rowDataToAppend, orderDate);
      // Row data is ready: write it to the array.
      MERGED_DATA.push(rowDataToAppend);
    }
  }
}

function addDimensions(rowDataArray, shipmentObject) {
  // Add dimensions compiled into one entry.
  rowDataArray[DIMENSIONS_COLUMN_INDEX] = shipmentObject.dimensions_length + 'X' 
    + shipmentObject.dimensions_width + 'X' 
    + shipmentObject.dimensions_height;
  return rowDataArray;
}

function addWeight(rowDataArray, shipmentObject) {
  // Weights are given in ounces.  Convert to pounds and ounces.
  var weightPounds = Math.floor(shipmentObject.weight_value / 16);
  var weightOunces = shipmentObject.weight_value % 16;
  rowDataArray[WEIGHT_COLUMN_INDEX] = weightPounds + ' lb ' + weightOunces + ' oz';
  return rowDataArray;
}

function addCarrierService(rowDataArray, shipmentObject) {
  // Split serviceCode into carrier and service, and insert into row.
  var splitServiceCode = shipmentObject.serviceCode.split('_',2);
  rowDataArray[CARRIER_USED_COLUMN_INDEX] = splitServiceCode[0];
  rowDataArray[CARRIER_COLUMN_INDEX] = splitServiceCode[0];
  rowDataArray[SERVICE_USED_COLUMN_INDEX] = splitServiceCode[1];
  rowDataArray[SERVICE_COLUMN_INDEX] = splitServiceCode[1];
  return rowDataArray;
}

function addQuarterFormula(rowNumber, rowDataArray) {
  // Get the cell reference, e.g. H3, for the shipment date.
  var shipDateCellReference = SHIP_DATE_COLUMN_LETTER + rowNumber;
  // Create a formula for looking up the corresponding store name.
  rowDataArray[QUARTER_COLUMN_INDEX] = '="Q"&roundup(month(' + shipDateCellReference + ')/3)&" "&year(' + shipDateCellReference + ')';
  return rowDataArray;
}

function applyShipmentItemFormats(rowNumber) {
  // No formats to apply in shipment items.
}

// Set values for an unmatched shipment, in the orders section of the sheet.
function setOrderValuesForUnmatchedShipment(rowDataToAppend, orderDate) {
  rowDataToAppend[ITEM_SHIPPED_COLUMN_INDEX] = 'TRUE';
  rowDataToAppend[ORDER_FULFILLED_COLUMN_INDEX] = ' ';
  rowDataToAppend[ORDER_DATE_COLUMN_INDEX] = orderDate;
  rowDataToAppend[ITEM_NAME_COLUMN_INDEX] = "!extra shipment data: couldn't match with an order";
  return rowDataToAppend;
}