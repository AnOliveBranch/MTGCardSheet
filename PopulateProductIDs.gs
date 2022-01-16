/*
 * This script will attempt to fill a TCGPlayer product ID for each item in the 'Card List' sheet
 * It uses the card name and set to search TCGPlayer's API
 * These product IDs will later be used to fetch pricing information in a different script
 *
 * Source code by Ryan Henderson, 2022
 * Licensed under GPLv3
*/

var access_token = CardLibrary.getToken();
var options = { method: 'get', headers: { Accept: 'application/json', Authorization: 'bearer ' + access_token } };
var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');

/*
 * Primary function to populate product IDs in cardSheet
 * Looks through sheet for items with no product ID and attempts to find and fill that value
*/
function populateProductIds() {
  let range = cardSheet.getDataRange().offset(1, 0);
  let vals = range.getValues();
  let currentProductIds = getItemListFromColumn(vals, 0);
  let sets = getSetListNoProductId(vals);
  let cardSetInfo = new Map();

  sets.forEach(function (cardSet) {
    cardSetInfo.set(cardSet, getCardProductIDsFromSet(cardSet));
  });

  for (const [setName, cardList] of cardSetInfo.entries()) {
    let cardsInSet = getCardsFromSetNoProductId(vals, setName);
    for (const [rowNum, cardName] of cardsInSet.entries()) {
      let productId = getProductId(cardName, cardList);
      if (productId === null) {
        throw new Error(`Could not find product ID for ${cardName} in set ${setName} on row ${rowNum}`);
      }
      currentProductIds[rowNum-2] = [productId];
    }
  }

  let productIdRange = cardSheet.getRange(2, 1, vals.length-1);
  productIdRange.setValues(currentProductIds);
}

/*
 * Returns a list of cards and product IDs from a given card set
 * 
 * @param {string} cardSet Set which the card is from (referred to as groupName in the API)
 * @param {number} counter Number of times this function has been called for this set
 * @param {Map<string, number>} cards List of cards for recursion
 * @return {Map<string, number>} List of cards and their product IDs within cardSet
*/
function getCardProductIDsFromSet(cardSet, counter=0, cards=new Map()) {
  let url = `https://api.tcgplayer.com/catalog/products?groupName=${convertEscapeCharacters(cardSet)}&productTypes=Cards&limit=100&offset=${counter*100}`;
  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (err) {
    if (err.toString().includes('403')) {
      throw new Error('403. TCGPlayer API does this, retry');
    } else if (cards.size === 0) {
      throw new Error(`404. Could not find set ${cardSet}`);
    } else {
      return cards;
    }
  }
  
  let json = JSON.parse(response.getContentText());

  json.results.forEach(function (cardInfo) {
    cards.set(cardInfo.name, cardInfo.productId);
  });
  return getCardProductIDsFromSet(cardSet, counter+1, cards);
}

/*
 * Returns the product ID for a given card from a set map
 *
 * @param {string} cardName Name of the card
 * @param {Map<string, number>} Map of cards for the set
 * @return {string} Product ID of the card, or null if not found
*/
function getProductId(cardName, cardSetMap) {
  let id = cardSetMap.get(cardName);
  return id === undefined ? null : id.toString();
}

/*
 * Returns a list of card sets that have a card with no product ID
 * 
 * @param {string[][]} range Values of DataRange of cards in the sheet
 * @return {Array[string]} List of card sets in range with a card with no product ID
*/
function getSetListNoProductId(range) {
  let numRows = range.length;
  let sets = [];
  for (let i = 0; i < numRows; i++) {
    if (range[i][0] === '') {
      let setName = range[i][4];
      if (!sets.includes(setName) && setName !== '') {
        sets.push(setName);
      }
    }
  }
  return sets;
}

/*
 * Returns a list of cards in a given set with no product ID
 * 
 * @param {string[][]} range Values of DataRange of cards in the sheet
 * @param {string} setName Name of set to search for
 * @return {Map<number, string>} List of cards in setName with no product ID and their row
*/
function getCardsFromSetNoProductId(range, setName) {
  let numRows = range.length;
  let cards = new Map();
  for (let i = 0; i < numRows; i++) {
    if (range[i][0] === '') {
      let rangeSetName = range[i][4];
      if (setName === rangeSetName) {
        let cardName = range[i][2];
        if (cardName !== '') {
          cards.set(i+2, cardName);
        }
      }
    }
  }
  return cards;
}

/*
 * Returns the list of items currently existing in the sheet for a given column
 * 
 * @param {string[][]} range Values of DataRange of cards in the sheet
 * @return {string[][]} List of values in the format [[%value1%],[%value2%],...]
*/
function getItemListFromColumn(range, column) {
  let numRows = range.length;
  let values = [];
  for (let i = 0; i < numRows; i++) {
    values.push([range[i][column].toString()]);
  }
  values.pop();
  return values;
}

/*
 * Returns the setName with necessary escape characters for API calls to function
 * 
 * @param {string} setName
 * @return {string} The setName with necessary escape characters replaced 
*/
function convertEscapeCharacters(setName) {
  setName = setName.replaceAll('&', '%26');

  return setName;
}