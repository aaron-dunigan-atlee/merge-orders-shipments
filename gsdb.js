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
  
  function addRows(sheetName, objects) {
    // Add rows to the sheet.  Values will be taken from objects, 
    // where attribute names match column headers.  Attributes that are
    // not in the headers will be ignored.
    
    var sheet = getSheet(sheetName);
    var headers = getHeader(sheetName);
    var ranges = [];
    for (var i=0; i<objects.length; i++) {
      var rowData = constructArrayFromObject(headers, objects[i]);  
      sheet.appendRow(rowData);
    }
  }
  
  function constructArrayFromObject(headers, rowObject) {
    var arrayData = [];
    // Construct an array whose values come from rowObject, 
    // and correspond with the property names in headers.
    for (var i=0; i<headers.length; i++) {
      if (rowObject.hasOwnProperty(headers[i])) {
        arrayData.push(rowObject[headers[i]]);
      } else {
        arrayData.push(""); 
      }
    }
    return arrayData;
  }
  
  function setRows(sheetName, keyName, rowObjects) {
    // Update some rows in the sheet.  Values will be taken from objects, 
    // where attribute names match column headers.  Attributes that are
    // not in the headers will be ignored.  Rows to be changed will be 
    // determined by filterObject, which should be one of the headers/attributes
    // If the value at keyName in each object does not uniquely determine a row,
    // all matching rows will be updated.
    // If a matching row is not found, a new row will be created.
    
    // Lock the script so we don't get multiple invocations trying to modify the sheet simulteaneously.
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);  
    var sheet = getSheet(sheetName);
    var headers = getHeadersAsPropertyNames(sheetName);
    var dbObjects = getObjects(sheetName);
    var offsetHeaderRow = 1;
    var offsetZeroBasedIndex = 1;
    for (var j=0; j<rowObjects.length; j++) {
      if (rowObjects[j].hasOwnProperty(keyName)) {
        var filterObject = {};
        filterObject[keyName] = [rowObjects[j][keyName]];
        var foundMatchingRow = false;
        var rowData = constructArrayFromObject(headers, rowObjects[j]); 
        for (var i = dbObjects.length - 1; i >= 0; i--) {
          if (objectMatchesFilterObject(dbObjects[i], filterObject)) { 
            foundMatchingRow = true;
            var rowToUpdate = i + offsetHeaderRow + offsetZeroBasedIndex;
            sheet.getRange(rowToUpdate, 1, 1, rowData.length).setValues([rowData]);
          }
        }
        if (!foundMatchingRow) {
          sheet.appendRow(rowData);
        }
      }
    }
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
  
  function delRows(sheetName, filterObject) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);  
    var sheet = getSheet(sheetName);
    var objects = getObjects(sheetName);
    var offsetHeaderRow = 1;
    var offsetZeroBasedIndex = 1;
    for (var i = objects.length - 1; i >= 0; i--) {
      if (objectMatchesFilterObject(objects[i], filterObject)) { 
        var rowToDelete = i + offsetHeaderRow + offsetZeroBasedIndex;
        sheet.deleteRow(rowToDelete);
      }
    }
    // Flush sheet before releasing lock.  See https://developers.google.com/apps-script/reference/lock/lock#releaseLock()
    SpreadsheetApp.flush();  
    lock.releaseLock();
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
  
  function hashArr(array, key) {
    return array.reduce(function(accumulator, object, index) {
      var propertyName = object[key];
      accumulator[propertyName] = object;
      return accumulator;
    }, {});
  }
  
  function getObjects(sheetObject) {
    var values = getValues(sheetObject);
    var headers = values.shift();
    var objects = [];
    for (var i = 0; i < values.length; i++) {
      var object = {};
      for (var j = 0; j < headers.length; j++) {
        var propertyName = headers[j];
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
  
  function getObjectsWithFilter(sheetName, filter, desiredProperties) {
    // A more efficient version of getRows():
    // Create an array of objects, one for each row in sheet sheetName.
    // 'desiredProperties' is an array of desired property names, e.g. studentsFirstName.
    // Return only rows that match the filter, 
    // and only properties from columns that match the properties parameter.
    // filter has same format as other filters: {property: [values]}

    var values = getValues(sheetName);
    var allPropertyNames = values.shift();
    if (desiredProperties == undefined) {
        desiredProperties = allPropertyNames;
    }
    // Get column indices corresponding to the properties parameter.
    var columnIndicesToReturn = [];
    for (var i=0; i<desiredProperties.length; i++) {
      var index = allPropertyNames.indexOf(desiredProperties[i]);
      if (index != -1) {
        columnIndicesToReturn.push(index);
      }
    }
    // Construct object filter columns whose items are of form
    // columnNumber: [values]
    var filterColumns = {};
    for (var filterProperty in filter) {
      var columnIndex = allPropertyNames.indexOf(filterProperty);
      if (columnIndex != -1) {
        filterColumns[columnIndex] = filter[filterProperty];
      }
    }
    // Now construct the array of objects.  
    var objects = [];
    for (var i = 0; i < values.length; i++) {
      if (rowMatchesIndexedFilterObject(values[i], filterColumns)) {
        var object = {};
        for (var j=0; j < columnIndicesToReturn.length; j++) {  
          var column = columnIndicesToReturn[j];
          object[allPropertyNames[column]] = values[i][column];
        }
        objects.push(object);
      }
    }
    return objects;
  }

  function rowMatchesIndexedFilterObject(rowArray, filterObject) {
    // Check whether the row of values matches the filter.
    // They match if for each column index in the filter, the value in
    // the row at that column matches one of the values in the filter 
    for (var columnIndex in filterObject) {
      if (filterObject[columnIndex].indexOf(rowArray[columnIndex]) == -1) { return false; }
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
    getRows: getRows,
    getObjectsWithFilter: getObjectsWithFilter,
    addRows: addRows,
    setRows: setRows,
    delRows: delRows,
    leftJoin: leftJoin,
    hashArr: hashArr
  }
})();

