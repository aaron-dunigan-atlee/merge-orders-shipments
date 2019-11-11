function testTime() {
  var start = new Date();
  populateMergedSheet();
  var end = new Date();
  Logger.log(end-start);
}

function testHeaders() {
  var headersArray = MERGED_SHEET.getDataRange().getValues().slice(0,2);
  Logger.log(headersArray);
  var headerProps = headersArray[0].filter(function(item, index) {
      return headersArray[1][index] == true;
    });
  Logger.log(headerProps);
}