/**
* One-time patch to update shipped column.
* Applied January 29 2020
*/
function updateShippedBlankToFalse() {
  var range = SpreadsheetApp.getActive().getSheetByName('Orders and Shipments').getRange("I6:I1028");
  var values = range.getValues()
  var backgrounds = range.getBackgrounds()
  for (var i=0; i<values.length; i++) {
    if (values[i][0] == "" && backgrounds[i][0] != "#b7e1cd") {
      values[i][0] = "FALSE"
    }
  }
  range.setValues(values)
}

/**
* One-time patch to update tax and shipping to the proper fields.
* Applied January 2020
*/
function updateTaxAndShipping() {
  initializeGlobals();
  var range = MERGED_SHEET.getRange(6, 1, 1010, MERGED_SHEET_WIDTH);
  var mergedData = getRowsData(MERGED_SHEET, range, 1, false, true);
  var ordersData = getRowsData(SpreadsheetApp.getActive().getSheetByName('Orders'));
  var data = [];
  for (var row=0; row<mergedData.length; row++) {
    var taxAndShipping = [mergedData[row].orderstaxamount,mergedData[row].ordersshippingamount]
    var orderItemId = mergedData[row].orderId
    if (orderItemId) {
      var order = ordersData.filter(function(element) {
        return element.orderId == orderItemId;
      })[0];
      if (order) {
        var tax = order.taxamount;
        var shipping = order.shippingamount;
        if (tax) taxAndShipping[0] = tax;
        if (shipping) taxAndShipping[1] = shipping;
      }
    }
    data.push(taxAndShipping);
  }
  var targetRange = MERGED_SHEET.getRange("R6:S1015");
  targetRange.setValues(data);
  //setRowsData(MERGED_SHEET, mergedData, null, 6);
}

/**
* One-time patch to add orderId field.
* Applied January 19, 2020
*/
function patchOrderId() {
  initializeGlobals();
  var range = MERGED_SHEET.getRange(6, 1, 1131-5, MERGED_SHEET_WIDTH);
  var mergedData = getRowsData(MERGED_SHEET, range, 1, false, true);
  var ordersData = getRowsData(SpreadsheetApp.getActive().getSheetByName('Orders'));
  var data = [];
  for (var row=0; row<mergedData.length; row++) {
    var uniqueKey = mergedData[row].orders_items_orderItemId;
    var rowDataToAdd = ['']
    if (uniqueKey) {
      var order = ordersData.filter(function(element) {
        return element.items_1_orderItemId == uniqueKey;
      })[0];
      if (order) {
        var orderId = order.orderId;
        if (orderId) {
          var rowDataToAdd = [orderId];
        }
      }
    }
    data.push(rowDataToAdd);
  }
  var targetRange = MERGED_SHEET.getRange("AK6:AK1131");
  targetRange.setValues(data);
}

/**
* One-time patch to add orders_advancedOptions_billToAccount field.
* Applied January 19, 2020
*/
function patchBillToAcct() {
  initializeGlobals();
  var range = MERGED_SHEET.getRange(6, 1, 1194-5, MERGED_SHEET_WIDTH);
  var mergedData = getRowsData(MERGED_SHEET, range, 1, false, true);
  var ordersData = getRowsData(SpreadsheetApp.getActive().getSheetByName('Orders'));
  var data = [];
  for (var row=0; row<mergedData.length; row++) {
    var uniqueKey = mergedData[row].orders_items_orderItemId;
    var rowDataToAdd = [mergedData[row].orders_advancedOptions_billToAccount];
    var found = false;
    if (uniqueKey) {
      var order = ordersData.filter(function(element) {
        return element.items_1_orderItemId == uniqueKey;
      })[0];
      if (order) {
        var orderId = order.advancedOptions_billToAccount;
        if (orderId) {
          found = true;
          rowDataToAdd = [orderId];
        }
      }
    } 
    if (!found) {
      var uniqueKey = mergedData[row].orders_orderId;
      if (uniqueKey) {
        var order = ordersData.filter(function(element) {
          return element.orderId == uniqueKey;
        })[0];
        if (order) {
          var orderId = order.advancedOptions_billToAccount;
          if (orderId) {
            rowDataToAdd = [orderId];
          }
        }
      }
    }
    data.push(rowDataToAdd);
  }
  var targetRange = MERGED_SHEET.getRange("AE6:AE1194");
  targetRange.setValues(data);
}

/**
* One-time patch to add orders_advancedOptions_billToAccount field.
* Applied January 19, 2020
*/
function patchBillToAcct2() {
  initializeGlobals();
  var range = MERGED_SHEET.getRange(6, 1, 1194-5, MERGED_SHEET_WIDTH);
  var mergedData = getRowsData(MERGED_SHEET, range, 1, false, true);
  var ordersData = getRowsData(SpreadsheetApp.getActive().getSheetByName('Orders'));
  var data = [];
  for (var row=0; row<mergedData.length; row++) {
    var uniqueKey = mergedData[row].orders_orderKey;
    var rowDataToAdd = [mergedData[row].orders_advancedOptions_billToAccount];
    var found = false;
    if (uniqueKey) {
      var order = ordersData.filter(function(element) {
        return element.orderKey == uniqueKey;
      })[0];
      if (order) {
        var orderId = order.advancedOptions_billToAccount;
        if (orderId) {
          found = true;
          rowDataToAdd = [orderId];
        }
      }
    } 

    data.push(rowDataToAdd);
  }
  var targetRange = MERGED_SHEET.getRange("AE6:AE1194");
  targetRange.setValues(data);
}


