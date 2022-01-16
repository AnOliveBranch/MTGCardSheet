/*
 * This script will attempt to fill collector number and rarity for cards
 * It uses the card name and set to search TCGPlayer's API
 *
 * Source code by Ryan Henderson, 2022
 * Licensed under GPLv3
*/

var access_token = CardLibrary.getToken();
var options = { method: 'get', headers: { Accept: 'application/json', Authorization: 'bearer ' + access_token } };
var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');
var dropdownSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dropdowns'); 

/*
 * Primary function to fill card info in cardSheet
 * Looks through sheet for cards with missing rarity and fills rarity and collector number
 * 
 * Then updates the dropdowns sheet with the most recent information
*/
function populateCardInfo() {
  let range = cardSheet.getDataRange().offset(1, 0);
  let vals = range.getValues();
  let currentCollectorNumbers = getItemListFromColumn(vals, 3);
  let currentRarities = getItemListFromColumn(vals, 5);
  let sets = getSetListNoRarity(vals);
  let cardSetInfo = new Map();

  sets.forEach(function (cardSet) {
    cardSetInfo.set(cardSet, getCardExtendedDataFromSet(cardSet));
  });

  for (const [setName, cardList] of cardSetInfo.entries()) {
    let cardsInSet = getCardsFromSetNoRarity(vals, setName);
    for (const [rowNum, cardName] of cardsInSet.entries()) {
      let extendedData = getExtendedData(cardName, cardList);
      if (extendedData === null) {
        throw new Error(`Could not find card info for ${cardName} in set ${setName} on row ${rowNum}`);
      }
      currentCollectorNumbers[rowNum-2] = [extendedData[0]];
      currentRarities[rowNum-2] = [extendedData[1]];
    }
  }

  let collectorNumberRange = cardSheet.getRange(2, 4, vals.length-1);
  collectorNumberRange.setValues(currentCollectorNumbers);
  let rarityRange = cardSheet.getRange(2, 6, vals.length-1);
  rarityRange.setValues(currentRarities);

  dropdownSheet.getDataRange().offset(1, 0).clear();
  let headerRange = cardSheet.getRange(1, 1, 1, 11);
  headerRange.getValues()[0].forEach(function (value) {
    if (getDropdownColumn(value) != -1) {
      generateDropdowns(value);
    }
  });
}

/*
 * Returns a list of cards and rarities and collector numbers from a given card set
 * 
 * @param {string} cardSet Set which the card is from (referred to as groupName in the API)
 * @param {number} counter Number of times this function has been called for this set
 * @param {Map<string, array[string]>>} cards List of cards for recursion
 * @return {Map<string, array[string]>} List of cards and their rarity and collector number within cardSet
*/
function getCardExtendedDataFromSet(cardSet, counter=0, cards=new Map()) {
  let url = `https://api.tcgplayer.com/catalog/products?groupName=${convertEscapeCharacters(cardSet)}&productTypes=Cards&limit=100&offset=${counter*100}&getExtendedFields=true`;
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
    let extendedData = cardInfo.extendedData;
    let rarity;
    let collectorNumber;
    extendedData.forEach(function (data) {
      if (data.name === 'Rarity') {
        rarity = data.value;
      } else if (data.name === 'Number') {
        collectorNumber = data.value;
      }
    });
    if (collectorNumber === undefined) {
      collectorNumber = '';
    }
    cards.set(cardInfo.name, [collectorNumber.toString(), convertRarity(rarity)]);
  });
  return getCardExtendedDataFromSet(cardSet, counter+1, cards);
}

/*
 * Returns the extended data for a given card from a set map
 *
 * @param {string} cardName Name of the card
 * @param {Map<string, array[string]>} Map of cards for the set
 * @return {array[string]} Extended data of the card, or null if not found
*/
function getExtendedData(cardName, cardSetMap) {
  let extendedData = cardSetMap.get(cardName);
  return extendedData === undefined ? null : extendedData;
}

/*
 * Returns the full name rarity given a TCGPlayer API rarity response
 * 
 * @param {string} rarity Single letter rarity of the card
 * @return {string} Complete word rarity of the card
*/ 
function convertRarity(rarity) {
  switch (rarity) {
    case 'C':
    return 'Common';
    case 'U':
    return 'Uncommon';
    case 'R':
    return 'Rare';
    case 'M':
    return 'Mythic';
    case 'T':
    return 'Token';
    case 'S':
    return 'Special';
    case 'L':
    return 'Land';
    case 'P':
    return 'Promo';
    default:
    return 'Unknown';
  }
}

/*
 * Returns a list of card sets that have a card with no rarity
 * 
 * @param {string[][]} range Values of DataRange of cards in the sheet
 * @return {Array[string]} List of card sets in range with a card with no rarity
*/
function getSetListNoRarity(range) {
  let numRows = range.length;
  let sets = [];
  for (let i = 0; i < numRows; i++) {
    if (range[i][5] === '') {
      let setName = range[i][4];
      if (!sets.includes(setName) && setName !== '') {
        sets.push(setName);
      }
    }
  }
  return sets;
}

/*
 * Returns a list of cards in a given set with no rarity
 * 
 * @param {string[][]} range Values of DataRange of cards in the sheet
 * @param {string} setName Name of set to search for
 * @return {Map<number, string>} List of cards in setName with no rarity and their row
*/
function getCardsFromSetNoRarity(range, setName) {
  let numRows = range.length;
  let cards = new Map();
  for (let i = 0; i < numRows; i++) {
    if (range[i][5] === '') {
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
 * Returns the column number with title of the given attritube in the Dropdowns sheet
 * 
 * @param {string} attribute Name of the attribute
 * @return {number} Index of the column where attribute is the header
*/
function getDropdownColumn(attribute) {
  let data = dropdownSheet.getRange('A1:Z1').getValues();
  return data[0].indexOf(attribute);
}

/*
 * Fills the column in the Dropdowns sheet with the values from cardSheet
 * 
 * @param {string} attribute Column to fill
*/
function generateDropdowns(attribute) {
  let entries = getDropdownEntries(attribute);
  let range = dropdownSheet.getRange(2, getDropdownColumn(attribute) + 1, entries.length);
  range.setValues(entries);
}

/*
 * Returns the list of items to fill in the Dropdowns sheet for an attribute
 * 
 * @param {string} attribute Column to get values for
 * @return {array[string]} Values to fill in the column
*/
function getDropdownEntries(attribute) {
  let headers = cardSheet.getRange('A1:Z1').getValues();
  let column = headers[0].indexOf(attribute);
  let range = cardSheet.getRange(2, column + 1, cardSheet.getLastRow());
  let rangeVals = range.getValues();
  let entries = [];
  for (let i = 0; i < rangeVals.length; i++) {
    let val = rangeVals[i][0];
    if (!entries.includes(val) && val !== '') {
      entries.push(val);
    }
  }
  entries.sort();
  entries.unshift('Any');
  let nestedEntries = [];
  for (let i = 0; i < entries.length; i++) {
    nestedEntries.push([entries[i]]);
  }
  return nestedEntries;
}