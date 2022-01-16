/*
 * This script will attempt to update pricing data for cards
 * It uses the product ID to search TCGPlayer's API
 * It returns TCGPlayer market value
 *
 * Source code by Ryan Henderson, 2022
 * Licensed under GPLv3
*/

var access_token = CardLibrary.getToken();
var options = { method: 'get', headers: { Accept: 'application/json', Authorization: 'bearer ' + access_token } };
var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');

/*
 * Primary function to update card prices
*/
function updateAllPrices() {
  let range = cardSheet.getDataRange().offset(1, 0);
  let vals = range.getValues();
  let allProductIds = getProductIdList(vals);
  let prices = new Map();

  allProductIds.forEach(function (id) {
    prices.set(id, [-1, -1]);
  });

  let pricingData = getPricingData(allProductIds, 0, prices);
  let valueColumn = [];
  for (let i = 0; i < allProductIds.length; i++) {
    let productId = vals[i][0].toString();
    let foil = vals[i][1] === 'Y';
    let pricing = pricingData.get(productId);
    if (pricing === null) {
      throw new Error(`Could not find pricing data for product ID ${productId}`);
    }
    if (foil) {
      valueColumn.push([pricing[1]]);
    } else {
      valueColumn.push([pricing[0]]);
    }
  }
  
  let valueRange = cardSheet.getRange(2, 7, allProductIds.length);
  valueRange.setValues(valueColumn);

  let dateRange = overviewSheet.getRange(8, 7);
  let date = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy");
  dateRange.setValue('Last updated ' + date);
}

/*
 * Returns a list of prices for a list of product IDs
 * 
 * @param {array[string]} productIds List of product IDs to pull pricing data for
 * @param {number} counter Number of products already searched for recursion
 * @param {Map<string, array[number]>>} cards List of cards for recursion
 * @return {Map<string, array[number]>} List of cards and their prices number within cardSet
*/
function getPricingData(productIds, counter, prices) {
  if (counter >= productIds.length) {
    return prices;
  }

  let productIdString = '';
  for (let i = counter; i < (counter+100 < productIds.length ? counter+100 : productIds.length); i++) {
    productIdString += productIds[i] + ',';
  }
  productIdString = productIdString.slice(0, -1);
  let url = `https://api.tcgplayer.com/pricing/product/${productIdString}`;
  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (err) {
    if (err.toString().includes('403')) {
      throw new Error('403. TCGPlayer API does this, retry');
    }
  }
  
  let json = JSON.parse(response.getContentText());
  json.results.forEach(function (pricingData) {
    let productId = pricingData.productId.toString();
    let currentlyStoredPrices = prices.get(productId);
    if (pricingData.subTypeName === 'Normal') {
      currentlyStoredPrices[0] = pricingData.marketPrice;
    } else if (pricingData.subTypeName === 'Foil') {
      currentlyStoredPrices[1] = pricingData.marketPrice;
    } else {
      throw new Error(`Unknown pricing subtype: ${pricingData.subTypeName} for response ${pricingData}`);
    }
  });

  return getPricingData(productIds, counter+100, prices);
}

/*
 * Returns the list of items currently existing in the sheet for a given column
 * 
 * @param {string[][]} range Values of DataRange of cards in the sheet
 * @return {string[]} List of values
*/
function getProductIdList(range) {
  let numRows = range.length;
  let values = [];
  for (let i = 0; i < numRows; i++) {
    if (range[i][0].toString() !== '') {
      values.push(range[i][0].toString());
    }
  }
  return values;
}