var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');
var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview');

function updateStatistics() {
  var allCards = cardSheet.getDataRange().offset(1, 0).getValues();
  overviewSheet.getRange(2, 7).setValue(getTotalValue(allCards));
  overviewSheet.getRange(3, 7).setValue(getTotalCount(allCards));
  overviewSheet.getRange(4, 7).setValue(getAverageValue(allCards));
  var allHighest = getHighestValue(allCards);
  overviewSheet.getRange(5, 7).setValue(allHighest[2]);
  overviewSheet.getRange(5, 8).setValue(allHighest[6]);

  var moneyThreshold = overviewSheet.getRange(1, 12).getValue();
  var moneyCards = getCardsAbove(moneyThreshold);
  overviewSheet.getRange(2, 11).setValue(getTotalValue(moneyCards));
  overviewSheet.getRange(3, 11).setValue(getTotalCount(moneyCards));
  overviewSheet.getRange(4, 11).setValue(getAverageValue(moneyCards));
  var moneyHighest = getHighestValue(moneyCards);
  overviewSheet.getRange(5, 11).setValue(moneyHighest[2]);
  overviewSheet.getRange(5, 12).setValue(moneyHighest[6]);

  var filteredCards = getFilteredCards();
  overviewSheet.getRange(2, 15).setValue(getTotalValue(filteredCards));
  overviewSheet.getRange(3, 15).setValue(getTotalCount(filteredCards));
  overviewSheet.getRange(4, 15).setValue(getAverageValue(filteredCards));
  var filteredHighest = getHighestValue(filteredCards);
  overviewSheet.getRange(5, 15).setValue(filteredHighest[2]);
  overviewSheet.getRange(5, 16).setValue(filteredHighest[6]);

  var labelRange = overviewSheet.getRange(12, 2, 4);
  var header = overviewSheet.getRange(11, 3);
  var labels = [['Total Value'], ['Total Card Count'], ['Average Card Value'], ['Highest Card Value']];
  getLocations().forEach(function (location) {
    labelRange.setValues(labels);
    header.setValue(location);

    var cards = getCards(location);
    header.offset(1, 0).setValue(getTotalValue(cards));
    header.offset(1, 0).setNumberFormat('$0.00');
    header.offset(2, 0).setValue(getTotalCount(cards));
    header.offset(3, 0).setValue(getAverageValue(cards));
    header.offset(3, 0).setNumberFormat('$0.00');
    var highest = getHighestValue(cards);
    header.offset(4, 0).setValue(highest[2]);
    header.offset(4, 1).setValue(highest[6]);
    header.offset(4, 1).setNumberFormat('$0.00');

    if (header.getColumn() == 15) {
      header = header.offset(6, -12);
      labelRange = labelRange.offset(6, -12);
    } else {
      header = header.offset(0, 4);
      labelRange = labelRange.offset(0, 4);
    }
  });
}

function getLocations() {
  var locations = [];
  var entries = getDropdownEntries('Location');
  entries.forEach(function (value) {
    locations.push(value[0]);
  });
  locations.shift();
  return locations;
}

function getCards(location) {
  var locationCards = [];
  var allCards = cardSheet.getDataRange().offset(1, 0).getValues();
  allCards.forEach(function (card) {
    if (card[8] == location) {
      locationCards.push(card);
    }
  });
  return locationCards;
}

function getCardsAbove(amount) {
  var moneyCards = [];
  var allCards = cardSheet.getDataRange().offset(1, 0).getValues();
  allCards.forEach(function (card) {
    if (card[6] >= amount) {
      moneyCards.push(card);
    }
  });
  return moneyCards;
}

function getTotalValue(rangeValues) {
  var total = 0.0;
  rangeValues.forEach(function (value) {
    var cardTotal = value[6] * value[7];
    total += cardTotal;
  });
  return total;
}

function getTotalCount(rangeValues) {
  var total = 0;
  rangeValues.forEach(function (value) {
    var cardCount = Number(value[7]);
    total += cardCount;
  });
  return total;
}

function getAverageValue(rangeValues) {
  return getTotalValue(rangeValues) / getTotalCount(rangeValues);
}

function getHighestValue(rangeValues) {
  var highest = rangeValues[0];
  rangeValues.forEach(function (value) {
    if (value[6] > highest[6]) {
      highest = value;
    }
  });
  return highest;
}

function getDropdownEntries(attribute) {
  var headers = cardSheet.getRange('A1:Z1').getValues();
  var col = headers[0].indexOf(attribute);
  var range = cardSheet.getRange(2, col + 1, cardSheet.getLastRow());
  var rangeVals = range.getValues();
  var entries = [];
  for (var i = 0; i < rangeVals.length; i++) {
    var val = rangeVals[i][0];
    if (entries.indexOf(val) == -1 && val != '') {
      entries.push(val);
    }
  }
  entries.sort();
  entries.unshift('Any');
  var nestedEntries = [];
  for (var i = 0; i < entries.length; i++) {
    var tempArray = [entries[i]];
    nestedEntries.push(tempArray);
  }
  return nestedEntries;
}

function getFilteredCards() {
  var optionsRange = overviewSheet.getRange(2, 3, 6).getValues();
  var colorOption = optionsRange[0][0];
  var setOption = optionsRange[1][0];
  var rarityOption = optionsRange[2][0];
  var locationOption = optionsRange[3][0];
  var quantityOption = optionsRange[4][0];
  var typeOption = optionsRange[5][0];
  var dataRange = cardSheet.getDataRange();
  var data = dataRange.getValues();

  var filteredData = [];

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
  return filteredData;
}