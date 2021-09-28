var bearerToken = "[token]"

function getNewAccessToken() {
  var url = "https://api.tcgplayer.com/token";
  var accessOptions = { method: 'post', headers: { Accept: 'application/x-www-form-urlencoded'}, payload: 'grant_type=client_credentials&client_id=[public_key]&client_secret=[private_key]'}
  var response = UrlFetchApp.fetch(url, accessOptions);
  Logger.log(response);
}