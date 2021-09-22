var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');

function sort() {
  cardSheet.getDataRange().offset(1, 0).sort([{ column : 11, ascending : true }, { column : 10, ascending : true }, { column : 3, ascending : true }, { column : 5, ascending : true }, { column : 2, ascending : true }, { column : 1, ascending : true }, { column : 9, ascending : true }]);
}