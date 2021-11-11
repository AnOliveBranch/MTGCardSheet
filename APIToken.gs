// The URL of this sheet
var sheetUrl = "[url]"
var spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);

// Get sheet with name 'Tokens', create it if it doesn't exist
var tokenSheet = spreadsheet.getSheetByName('Tokens');
if (tokenSheet == null) {
  spreadsheet.insertSheet('Tokens');
  tokenSheet = spreadsheet.getSheetByName('Tokens');
}

function getNewAccessToken() {
  var url = "https://api.tcgplayer.com/token";
  var accessOptions = { method: 'post', headers: { Accept: 'application/x-www-form-urlencoded'}, payload: 'grant_type=client_credentials&client_id=[public_key]&client_secret=[private_key]'}
  var response = UrlFetchApp.fetch(url, accessOptions);
  return JSON.parse(response.getContentText());
}

function genToken() {
  // Get the new bearer token 
  var tokenInfo = getNewAccessToken();

  // Extract the token itself and its expiration
  var token = tokenInfo.access_token;
  var expiration = tokenInfo['.expires'];

  // Turn the string expiration into millis
  var date = Date.parse(expiration);
  var range = tokenSheet.getDataRange();

  // Set the new token to the bottom row of the sheet (this way we'll keep a history of past tokens as well)
  var expirationRange = tokenSheet.getRange(range.getNumColumns() + 1, 1);
  var tokenRange = tokenSheet.getRange(range.getNumColumns() + 1, 2);
  expirationRange.setValue(date);
  tokenRange.setValue(token);
}

function getToken() {
  if (tokenSheet.getRange(1, 1).getValue() == '') {
    tokenSheet.deleteRow(1);
  }
  // Get the bottom entries of the sheet
  var dataRange = tokenSheet.getDataRange();
  var expiration = tokenSheet.getRange(dataRange.getNumColumns(), 1).getValue();
  var token = tokenSheet.getRange(dataRange.getNumColumns(), 2).getValue();

  // Check token expiration against current date
  var d = Date.parse(new Date());
  // Make new token if current will expire in less than an hour (or has expired)
  if (expiration == null || token == null || d-3600000 > expiration) {
    // Token is expired, gen a new one and re-set the values
    genToken();
    dataRange = tokenSheet.getDataRange();
    token = tokenSheet.getRange(dataRange.getNumColumns(), 2).getValue();
  }
  return token;
}