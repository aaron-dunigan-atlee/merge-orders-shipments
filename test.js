function testTime() {
  var start = new Date();
  populateMergedSheet();
  var end = new Date();
  Logger.log(end-start);
}