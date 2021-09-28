This project is a series of Google Apps Script files that I use for cataloging my own personal card collection found [here](https://docs.google.com/spreadsheets/d/1zdMBwkbjX_JPWCNWO0-qNcOaHB_9JA4JDOWHxhDbpd4/edit#gid=28962468)

It uses [TCGPlayer](https://www.tcgplayer.com/)'s API to fetch information on the cards and their values.

Sheet setup
-----
To use this, a couple of things would need to be done. First, you need to apply and be approved for TCGPlayer's API, which can be done [here](https://developer.tcgplayer.com/developer-application-form.html). You will also need to create a new Google Sheet. You could either make a copy of my existing sheet or make your own. If you make your own sheet, you'll need to go to `Tools`, then `Script Editor` and create the `UpdateFilters.gs`, `FetchValues.gs`, `UpdateStatistics.gs`, and `SortCards.gs` files, pasting in the code.

To copy the sheet, open the [sheet](https://docs.google.com/spreadsheets/d/1zdMBwkbjX_JPWCNWO0-qNcOaHB_9JA4JDOWHxhDbpd4/edit#gid=28962468), go to `File`, `Make a Copy` and choose a name and location.

Setting up the API Key Library
-----
You will need to create a second sheet for the purpose of holding your API keys. Keep this sheet private. Open the script editor, and paste in the contents of `APITokens.gs` into the script. Once you have API approval, deplace the public and private keys in `getNewAccessToken()`, then run that function directly from the Apps Script page. A bearer token will be pasted to the console, paste that into the `bearerToken` field. Click `Deploy` in the top right corner, click `New Deployment`, the type should be `Library`. Give it a description, then click `Deploy`. Go to the settings cog on the left side of the screen, and copy the `Script ID`. Go back to your main sheet with the other four scripts. Click the `+` on `Libraries` on the left side panel, and paste in the Script ID you just copied, and click `Look Up`. Keep the version as `HEAD (Development Mode)` and make the identifier `CardLibrary`, then click `Add`.

Modifying your sheet
-----
The scripts are setup very specifically to work with my sheet layout. You could change the sheet layout if desired, you'd just have to change some of the code as well to reflect that. TODO: Add comments to the code to show what values might need changing and to explain what the code does.

Using the sheet
-----
Once you have your sheet setup, here's how to actually use it. Go to the `Card List` sheet on the bottom. For each card, put the TCGPlayer product ID (I use a database from [MTGJSON](https://mtgjson.com/)), whether the card is foiled or not (`Y` or `N`), the quantity of that card you have, where it is, its color identity, and type. For card games other than Magic: The Gathering these fields may need to be changed, which would mean changing the code. Once you've put in your product IDs (you can do lots at once), go back to the `Overview` sheet and click the `Update Card Info & Pricing` button. Depending on how many cards you have, this process may take a while. Once it's finished, the `Update Statistics` button will update all of the statistics about the collection on the overview page. The `Sort Cards` button will sort the `Card List` sheet according to the order listed in `SortCards.gs`. Finally, `Update Filters` will copy cards that match the filter options into the `Filtered Cards` sheet.