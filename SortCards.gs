var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');

// Define the sort order 
// Card List column header first, then ascending true or false
var sortOrder = ['Type', true, 'Color Identity', true, 'Name', true, 'Set', true, 'Foil?', true, 'Product ID', true, 'Location', true];

function convertColumn(name) {
  // Get the list of the sheet headers
  var headers = cardSheet.getRange(1, 1, 1, cardSheet.getDataRange().getNumColumns()).getValues()[0];

  // Loop through the headers
  // We need the column number, and column and row numbers in sheets start at 1
  for (var i = 1; i <= headers.length; i++) {
    // Check for match, return the column number
    if (name == headers[i-1]) {
      return i;
    }
  }
  // Invalid name, return null
  return null;
}

function getFinalSortOrder() {
  var columnOrder = [];
  // Loop through headers in sort order
  for (var i = 0; i < sortOrder.length; i+=2) {
    // Get column number and ascending or not for each entry
    var columnNum = convertColumn(sortOrder[i]);
    var asc = sortOrder[i+1];

    // Ensure headers are correct, script will throw error if not
    if (columnNum == null) {
      throw new Error('Failed: Could not get column number for ' + sortOrder[i]);
    }
    
    // Create the pair and add it to the list
    var pair = { column : columnNum, ascending: asc };
    columnOrder.push(pair);
  }
  return columnOrder;
}

function sort() {
  // Offset by 1 to not sort the header
  cardSheet.getDataRange().offset(1, 0).sort(getFinalSortOrder());
}