var access_token = "[redacted]"
var options = { method: 'get', headers: { Accept: 'application/json', Authorization: 'bearer ' + access_token } };
var cardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Card List');
var dropdownSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dropdowns'); 
var infoLineCounter = 2;
var valueLineCounter = 2;

function test() {
  var url = "https://api.tcgplayer.com/pricing/product/240595,165571";
  //url = "https://api.tcgplayer.com/catalog/products/203538?getExtendedFields=true"
  url = "https://api.tcgplayer.com/catalog/products?productName=Shapeshifter Token (001) &groupName=Modern Horizons"
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);
}

function getNewAccessToken() {
  var url = "https://api.tcgplayer.com/token";
  var accessOptions = { method: 'post', headers: { Accept: 'application/x-www-form-urlencoded'}, payload: 'grant_type=client_credentials&client_id=[redacted]&client_secret=[redacted]'}
  var response = UrlFetchApp.fetch(url, accessOptions);
  Logger.log(response);
}

function updateAll() {
  var range = cardSheet.getDataRange().offset(1, 0);
  var vals = range.getValues();
  var idList;
  var numRows = range.getNumRows();

  var headerRange = cardSheet.getRange(1, 1, 1, 11);
  headerRange.getValues()[0].forEach(function (value) {
    if (getDropdownColumn(value) != -1) {
      generateDropdowns(value);
    }
  });
  infoLineCounter = 2;
  valueLineCounter = 2;
  for (var i = 0; i < numRows; i+= 100) {
    idList = "";
    for (var j = i; j < i+100; j++) {
      if (j >= numRows) {
        break;
      }
      var val = vals[j][0];
      if (val == '') {
        continue;
      }
      idList += val + ',';
    }
    if (idList == '') {
      continue;
    }
    idList = idList.substr(0, idList.length - 1);
    setInfo(parseInfo(fetchInfo(idList)));
    setValues(parseValues(fetchValues(idList)));
  }
  var dateRange = overviewSheet.getRange(8, 7);
  var date = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy");
  dateRange.setValue('Last updated ' + date);
}

function setInfo(info) {
  var range = cardSheet.getDataRange();
  var data = unwrapValues(range.offset(1, 0, range.getNumRows(), 1).getValues());
  var dataToSet = [];
  var limits = infoLineCounter;
  for (var i = limits-2; i < limits+100; i++) {
    info.forEach(function (card) {
      if (card[0] == data[i]) {
        infoLineCounter++;
        var cardDataRange = cardSheet.getRange(i+2, 1, 1, 6);
        var cardCopy = [...card];
        cardCopy.splice(1, 0, cardDataRange.getValues()[0][1]);
        dataToSet.push(cardCopy);
        return;
      }
    });
  }
  var setRange = cardSheet.getRange((infoLineCounter-dataToSet.length), 1, dataToSet.length, 6);
  setRange.setValues(dataToSet);
}

function setValues(values) {
  var range = cardSheet.getDataRange();
  var data = unwrapValues(range.offset(1, 0, range.getNumRows(), 2).getValues());
  var dataToSet = [];
  var limits = valueLineCounter;
  for (var i = (limits-2)*2; i < (limits+100)*2; i+=2) {
    for (var j = 0; j < values.length; j++) {
      if (values[j][0] == data[i]) {
        if ((data[i+1] == 'Y' && values[j][2] == 'Foil') || (data[i+1] == 'N' && values[j][2] == 'Normal')) {
          valueLineCounter++;
          dataToSet.push([values[j][1]]);
          break;
        }
      }
    }
  }
  var setRange = cardSheet.getRange((valueLineCounter-dataToSet.length), 7, dataToSet.length, 1);
  setRange.setValues(dataToSet);
}

function fetchInfo(idList) {
  var infoUrl = "https://api.tcgplayer.com/catalog/products/" + idList + "?getExtendedFields=true";
  var response = UrlFetchApp.fetch(infoUrl, options);
  return JSON.parse(response.getContentText());
}

function fetchValues(idList) {
  var valueUrl = "https://api.tcgplayer.com/pricing/product/" + idList;
  var response = UrlFetchApp.fetch(valueUrl, options);
  return JSON.parse(response.getContentText());
}

function parseInfo(response) {
  if (!response.success) {
    Logger.log(response.errors);
    throw new Error('Failed: ' + response.errors);
  }
  var info = [];
  response.results.forEach(function (cardData) {
    var cardInfo = [cardData.productId, cardData.name];
    cardData.extendedData.forEach(function (extended) {
      if (extended.name == 'Number') {
        cardInfo[2] = extended.value;
      } else if (extended.name == 'Rarity') {
        cardInfo[4] = convertRarity(extended.value);
      }
    });
    cardInfo[3] = convertSet(cardData.groupId);
    info.push(cardInfo);
  });
  return info;
}

function parseValues(response) {
  if (!response.success) {
    Logger.log(response.errors);
    throw new Error('Failed: ' + response.errors);
  }
  var info = [];
  response.results.forEach(function (valueData) {
    var cardValue = [valueData.productId, valueData.marketPrice, valueData.subTypeName];
    info.push(cardValue);
  });
  return info;
}

function unwrapValues(values) {
  var vals = [];
  values.forEach(function (value) {
    value.forEach(function (temp) {
      vals.push(temp);
    });
  });
  return vals;
}

function wrapValues(values) {
  var wrapped = [];
  wrapped.push(values);
  return wrapped;
}

function getDropdownColumn(attribute) {
  var data = dropdownSheet.getRange('A1:Z1').getValues();
  return data[0].indexOf(attribute);
}

function generateDropdowns(attribute) {
  var entries = getDropdownEntries(attribute);
  var range = dropdownSheet.getRange(2, getDropdownColumn(attribute) + 1, entries.length);
  range.setValues(entries);
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

function convertSet(setId) {
  switch (setId) {
    case 7:
    return "Limited Edition Alpha";
    case 17:
    return "Limited Edition Beta";
    case 115:
    return "Unlimited Edition";
    case 1526:
    return "Collectors’ Edition";
    case 1527:
    return "Intl. Collectors’ Edition";
    case 11:
    return "Arabian Nights";
    case 8:
    return "Antiquities";
    case 97:
    return "Revised Edition";
    case 1689:
    return "Foreign Black Border";
    case 65:
    return "Legends";
    case 2642:
    return "Summer Magic / Edgar";
    case 109:
    return "The Dark";
    case 41:
    return "Fallen Empires";
    case 72:
    return "Media Inserts";
    case 46:
    return "Fourth Edition";
    case 58:
    return "Ice Age";
    case 22:
    return "Chronicles";
    case 2379:
    return "Renaissance";
    case 57:
    return "Homelands";
    case 6:
    return "Alliances";
    case 2778:
    return "Rivals Quick Start Set";
    case 74:
    return "Mirage";
    case 2742:
    return "Introductory Two-Player Set";
    case 120:
    return "Visions";
    case 44:
    return "Fifth Edition";
    case 86:
    return "Portal";
    case 119:
    return "Vanguard Series";
    case 121:
    return "Weatherlight";
    case 2198:
    return "World Championship Decks 1997";
    case 108:
    return "Tempest";
    case 62:
    return "Judge Gift Cards 1998";
    case 107:
    return "Stronghold";
    case 40:
    return "Exodus";
    case 87:
    return "Portal Second Age";
    case 113:
    return "Unglued";
    case 9:
    return "Asia Pacific Land Program";
    case 118:
    return "Urza's Saga";
    case 1275:
    return "Anthologies";
    case 62:
    return "Judge Gift Cards 1999";
    case 117:
    return "Urza's Legacy";
    case 23:
    return "Classic Sixth Edition";
    case 116:
    return "Urza's Destiny";
    case 105:
    return "Starter 1999";
    case 88:
    return "Portal Three Kingdoms";
    case 56:
    return "Guru";
    case 73:
    return "Mercadian Masques";
    case 15:
    return "Battle Royale Box Set";
    case 61:
    return "Junior Super Series";
    case 45:
    return "Friday Night Magic 2000";
    case 62:
    return "Judge Gift Cards 2000";
    case 38:
    return "European Land Program";
    case 78:
    return "Nemesis";
    case 106:
    return "Starter 2000";
    case 94:
    return "Prophecy";
    case 16:
    return "Beatdown Box Set";
    case 60:
    return "Invasion";
    case 62:
    return "Judge Gift Cards 2001";
    case 85:
    return "Planeshift";
    case 2:
    return "Seventh Edition";
    case 10:
    return "Apocalypse";
    case 80:
    return "Odyssey";
    case 1311:
    return "Deckmasters";
    case 62:
    return "Judge Gift Cards 2002";
    case 112:
    return "Torment";
    case 63:
    return "Judgment";
    case 81:
    return "Onslaught";
    case 62:
    return "Judge Gift Cards 2003";
    case 66:
    return "Legions";
    case 101:
    return "Scourge";
    case 3:
    return "Eighth Edition";
    case 75:
    return "Mirrodin";
    case 62:
    return "Judge Gift Cards 2004";
    case 27:
    return "Darksteel";
    case 43:
    return "Fifth Dawn";
    case 20:
    return "Champions of Kamigawa";
    case 114:
    return "Unhinged";
    case 62:
    return "Judge Gift Cards 2005";
    case 18:
    return "Betrayers of Kamigawa";
    case 99:
    return "Saviors of Kamigawa";
    case 4:
    return "Ninth Edition";
    case 95:
    return "Ravnica: City of Guilds";
    case 62:
    return "Judge Gift Cards 2006";
    case 55:
    return "Guildpact";
    case 21:
    return "Champs and States";
    case 28:
    return "Dissension";
    case 24:
    return "Coldsnap";
    case 1348:
    return "Coldsnap Theme Decks";
    case 111:
    return "Time Spiral Timeshifted";
    case 110:
    return "Time Spiral";
    case 62:
    return "Judge Gift Cards 2007";
    case 83:
    return "Planar Chaos";
    case 93:
    return "Pro Tour Promos";
    case 54:
    return "Grand Prix Promos";
    case 51:
    return "Future Sight";
    case 1:
    return "Tenth Edition";
    case 67:
    return "Lorwyn";
    case 33:
    return "Duel Decks: Elves vs. Goblins";
    case 62:
    return "Judge Gift Cards 2008";
    case 77:
    return "Morningtide";
    case 102:
    return "Shadowmoor";
    case 39:
    return "Eventide";
    case 47:
    return "From the Vault: Dragons";
    case 103:
    return "Shards of Alara";
    case 35:
    return "Duel Decks: Jace vs. Chandra";
    case 62:
    return "Judge Gift Cards 2009";
    case 26:
    return "Conflux";
    case 31:
    return "Duel Decks: Divine vs. Demonic";
    case 5:
    return "Alara Reborn";
    case 68:
    return "Magic 2010";
    case 48:
    return "From the Vault: Exiled";
    case 84:
    return "Planechase";
    case 124:
    return "Zendikar";
    case 34:
    return "Duel Decks: Garruk vs. Liliana";
    case 91:
    return "Premium Deck Series: Slivers";
    case 62:
    return "Judge Gift Cards 2010";
    case 122:
    return "Worldwake";
    case 37:
    return "Duel Decks: Phyrexia vs. the Coalition";
    case 98:
    return "Rise of the Eldrazi";
    case 1274:
    return "Duels of the Planeswalkers";
    case 12:
    return "Archenemy";
    case 69:
    return "Magic 2011";
    case 50:
    return "From the Vault: Relics";
    case 32:
    return "Duel Decks: Elspeth vs. Tezzeret";
    case 100:
    return "Scars of Mirrodin";
    case 89:
    return "Premium Deck Series: Fire and Lightning";
    case 62:
    return "Judge Gift Cards 2011";
    case 76:
    return "Mirrodin Besieged";
    case 36:
    return "Duel Decks: Knights vs. Dragons";
    case 79:
    return "New Phyrexia";
    case 25:
    return "Commander 2011";
    case 70:
    return "Magic 2012";
    case 49:
    return "From the Vault: Legends";
    case 30:
    return "Duel Decks: Ajani vs. Nicol Bolas";
    case 59:
    return "Innistrad";
    case 90:
    return "Premium Deck Series: Graveborn";
    case 62:
    return "Judge Gift Cards 2012";
    case 125:
    return "Dark Ascension";
    case 367:
    return "Duel Decks: Venser vs. Koth";
    case 362:
    return "Avacyn Restored";
    case 363:
    return "Planechase 2012";
    case 364:
    return "Magic 2013";
    case 368:
    return "From the Vault: Realms";
    case 365:
    return "Duel Decks: Izzet vs. Golgari";
    case 370:
    return "Return to Ravnica";
    case 568:
    return "Commander's Arsenal";
    case 62:
    return "Judge Gift Cards 2013";
    case 569:
    return "Gatecrash";
    case 601:
    return "Duel Decks: Sorin vs. Tibalt";
    case 570:
    return "Dragon's Maze";
    case 1111:
    return "Modern Masters";
    case 1113:
    return "Magic 2014";
    case 1141:
    return "From the Vault: Twenty";
    case 1145:
    return "Duel Decks: Heroes vs. Monsters";
    case 1144:
    return "Theros";
    case 1164:
    return "Commander 2013";
    case 62:
    return "Judge Gift Cards 2014";
    case 1276:
    return "Born of the Gods";
    case 1166:
    return "Duel Decks: Jace vs. Vraska";
    case 1277:
    return "Journey into Nyx";
    case 1346:
    return "Modern Event Deck 2014";
    case 1312:
    return "Conspiracy";
    case 1293:
    return "Magic 2015";
    case 1475:
    return "From the Vault: Annihilation";
    case 1477:
    return "Duel Decks: Speed vs. Cunning";
    case 1356:
    return "Khans of Tarkir";
    case 1476:
    return "Commander 2014";
    case 33:
    return "Duel Decks Anthology: Elves vs. Goblins";
    case 62:
    return "Judge Gift Cards 2015";
    case 1507:
    return "Ugin's Fate";
    case 1497:
    return "Fate Reforged";
    case 1511:
    return "Duel Decks: Elspeth vs. Kiora";
    case 1515:
    return "Dragons of Tarkir";
    case 1520:
    return "Tarkir Dragonfury";
    case 1503:
    return "Modern Masters 2015";
    case 1512:
    return "Magic Origins";
    case 1577:
    return "From the Vault: Angels";
    case 1641:
    return "Duel Decks: Zendikar vs. Eldrazi";
    case 1645:
    return "Battle for Zendikar";
    case 1649:
    return "Zendikar Expeditions";
    case 1673:
    return "Commander 2015";
    case 62:
    return "Judge Gift Cards 2016";
    case 1693:
    return "Oath of the Gatewatch";
    case 1726:
    return "Duel Decks: Blessed vs. Cursed";
    case 1708:
    return "Shadows over Innistrad";
    case 1765:
    return "Welcome Deck 2016";
    case 1740:
    return "Eternal Masters";
    case 1790:
    return "Eldritch Moon";
    case 1821:
    return "From the Vault: Lore";
    case 1794:
    return "Conspiracy: Take the Crown";
    case 1835:
    return "Duel Decks: Nissa vs. Ob Nixilis";
    case 1791:
    return "Kaladesh";
    case 1837:
    return "Kaladesh Inventions";
    case 1792:
    return "Commander 2016";
    case 1793:
    return "Planechase Anthology";
    case 62:
    return "Judge Gift Cards 2017";
    case 1857:
    return "Aether Revolt";
    case 1879:
    return "Modern Masters 2017";
    case 1905:
    return "Duel Decks: Mind vs. Might";
    case 1930:
    return "Welcome Deck 2017";
    case 1882:
    return "Amonkhet";
    case 1909:
    return "Amonkhet Invocations";
    case 1933:
    return "Commander Anthology";
    case 1904:
    return "Archenemy: Nicol Bolas";
    case 1934:
    return "Hour of Devastation";
    case 2009:
    return "Commander 2017";
    case 2043:
    return "Ixalan";
    case 2076:
    return "Duel Decks: Merfolk vs. Goblins";
    case 2050:
    return "Iconic Masters";
    case 2077:
    return "Explorers of Ixalan";
    case 2078:
    return "From the Vault: Transform";
    case 2092:
    return "Unstable";
    case 62:
    return "Judge Gift Cards 2018";
    case 2098:
    return "Rivals of Ixalan";
    case 2189:
    return "Masters 25";
    case 2207:
    return "Duel Decks: Elves vs. Inventors";
    case 2199:
    return "Dominaria";
    case 2245:
    return "Battlebond";
    case 2246:
    return "Commander Anthology Volume II";
    case 2220:
    return "Signature Spellbook: Jace";
    case 2247:
    return "Global Series Jiang Yanggu & Mu Yanling";
    case 2250:
    return "Core Set 2019";
    case 2251:
    return "Commander 2018";
    case 2290:
    return "Guilds of Ravnica";
    case 2326:
    return "Mythic Edition";
    case 2334:
    return "GRN Guild Kit";
    case 2359:
    return "Game Night";
    case 2361:
    return "Ultimate Box Topper";
    case 2360:
    return "Ultimate Masters";
    case 2366:
    return "Ravnica Allegiance";
    case 2388:
    return "RNA Guild Kit";
    case 2418:
    return "War of the Spark";
    case 2422:
    return "Modern Horizons";
    case 2447:
    return "Signature Spellbook: Gideon";
    case 2441:
    return "Core Set 2020";
    case 92:
    return "Core Set 2020 Promos";
    case 2481:
    return "Commander 2019";
    case 2494:
    return "Throne of Eldraine";
    case 2552:
    return "Ponies: The Galloping";
    case 2570:
    return "Mystery Booster Playtest Cards 2019";
    case 2572:
    return "Mystery Booster";
    case 2359:
    return "Game Night 2019";
    case 2576:
    return "Secret Lair Drop";
    case 2568:
    return "Theros Beyond Death";
    case 2598:
    return "Unsanctioned";
    case 2571:
    return "Mystery Booster Retail Edition Foils";
    case 2634:
    return "Commander 2020";
    case 2603:
    return "Ikoria: Lair of Behemoths";
    case 2632:
    return "Secret Lair: Ultimate Edition";
    case 2633:
    return "Signature Spellbook: Chandra";
    case 2653:
    return "Core Set 2021";
    case 2654:
    return "Jumpstart";
    case 2655:
    return "Double Masters";
    case 2716:
    return "Zendikar Rising Commander";
    case 2713:
    return "Zendikar Rising Expeditions";
    case 2648:
    return "Zendikar Rising";
    case 2715:
    return "The List";
    case 2708:
    return "Commander Legends";
    case 2699:
    return "Commander Collection: Green";
    case 2766:
    return "Kaldheim Commander";
    case 2750:
    return "Kaldheim";
    case 2772:
    return "Time Spiral Remastered";
    case 2801:
    return "Commander 2021";
    case 2791:
    return "Strixhaven Mystical Archive";
    case 2773:
    return "Strixhaven: School of Mages";
    case 2422:
    return "Modern Horizons 1 Timeshifts";
    case 2809:
    return "Modern Horizons 2";
    case 2822:
    return "Art Series: Modern Horizons 2";
    case 2870:
    return "Forgotten Realms Commander";
    case 2823:
    return "Adventures in the Forgotten Realms";
    case 2570:
    return "Mystery Booster Playtest Cards 2021";
    case 2864:
    return "Innistrad: Midnight Hunt";
    case 2862:
    return "Innistrad: Crimson Vow";
    case 2905:
    return "Commander Collection: Black";
    case 2156:
    return "Buy-A-Box Promos";
    case 2543:
    return "Throne of Eldraine Promos";
    case 53:
    return "WPN & Gateway Promos";
    case 1163:
    return "Unique and Miscellaneous Promos";
    case 2376:
    return "MagicFest Cards";
    case 1874:
    return "Launch Party & Release Event Promos";
    case 52:
    return "Game Day & Store Championship Promos";
    case 1490:
    return "Duel Decks: Anthology";
    default:
    return 'Unknown';
  }
}