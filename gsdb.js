/*
This entire page sets up an abstract structure that allows us to get data from a sheet
and reference it by the header name given at the top.
So, for example, on the main.gs page, I use the command
orderItem.items_name
and it knows which column to find "items_name" in on the order sheet
because I've run this code first.
*/

var GsDb = (function () {

  function getRows(sheetObject, filterObject) {
    var objects = getObjects(sheetObject);
    return objects.filter(function(object) {
      return objectMatchesFilterObject(object, filterObject)
    });
  }
  
  function addRows() {  }
  
  function setRow() {  }
  
  function delRows() {  }
  
  function getObjects(sheetObject) {
    var values = getValues(sheetObject);
    // Attribute names will be of the form sheetname_headerName
    var prefix = sheetObject.getName().toLowerCase() + '_';
    var headers = values.shift();
    // Create an object for each row.
    var objects = [];
    for (var i = 0; i < values.length; i++) {
      var object = {};
      for (var j = 0; j < headers.length; j++) {
        var propertyName = prefix + headers[j];
        object[propertyName] = values[i][j];
      }
      objects.push(object);
    }
    return objects;
  }
  
  function objectMatchesFilterObject(object, filterObject) {
    for (var property in filterObject) {
      if (filterObject[property].indexOf(object[property]) == -1) { return false; }
    }
    return true;
  }
  
  function getSheet(sheetName) {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }

  function getValues(sheetObject) {
    return sheetObject.getDataRange().getValues();
  }
 
  function getHeaders(sheetName) {
    // Return a string array of the headers as they appear on the sheet.
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    return sheet.getDataRange().getValues()[0];
  } 
  
  return {
    getRows: getRows
  }
})();

