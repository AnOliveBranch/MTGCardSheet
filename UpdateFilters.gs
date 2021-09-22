var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');
var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview');
var filteredSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Filtered Cards');

function updateFilters() {
  var optionsRange = overviewSheet.getRange(2, 3, 6).getValues();
  var colorOption = optionsRange[0][0];
  var setOption = optionsRange[1][0];
  var rarityOption = optionsRange[2][0];
  var locationOption = optionsRange[3][0];
  var quantityOption = optionsRange[4][0];
  var typeOption = optionsRange[5][0];
  var dataRange = cardSheet.getDataRange();
  var data = dataRange.getValues();

  var filteredData = [data[0]];
  filteredSheet.clearContents();

  for (var i = 2; i <= dataRange.getNumRows(); i++) {
    var vals = data[i-1];
    if (vals[9] != colorOption && colorOption != 'Any') {
      continue;
    }
    if (vals[4] != setOption && setOption != 'Any') {
      continue;
    }
    if (vals[5] != rarityOption && rarityOption != 'Any') {
      continue;
    }
    if (vals[8] != locationOption && locationOption != 'Any') {
      continue;
    }
    if (vals[7] != quantityOption && quantityOption != 'Any') {
      continue;
    }
    if (vals[10] != typeOption && typeOption != 'Any') {
      continue;
    }
    filteredData.push(vals);
  }
  var range = filteredSheet.getRange(1, 1, filteredData.length, filteredData[0].length);
  range.setValues(filteredData);
}