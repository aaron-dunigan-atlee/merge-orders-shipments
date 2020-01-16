/**
* Patch to update tax and shipping to the proper fields.
*/

function updateTaxAndShipping() {
  initializeGlobals();
  var range = MERGED_SHEET.getRange(6, 1, 1010, MERGED_SHEET_WIDTH);
  var mergedData = getRowsData(MERGED_SHEET, range, 1, false, true);
  var ordersData = getRowsData(SpreadsheetApp.getActive().getSheetByName('Orders'));
  var data = [];
  for (var row=0; row<mergedData.length; row++) {
    var taxAndShipping = [mergedData[row].orderstaxamount,mergedData[row].ordersshippingamount]
    var orderItemId = mergedData[row].ordersitemsorderitemid
    if (orderItemId) {
      var order = ordersData.filter(function(element) {
        return element.itemsorderitemid == orderItemId;
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
