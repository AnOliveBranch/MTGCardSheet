This project is a series of Google Apps Script files that I use for cataloging my own personal card collection found [here](https://docs.google.com/spreadsheets/d/1_zD-1VYKzvuM7W_ja9X6SWEYxPK_HNof1y0orEvWdWY/edit?usp=sharing) (this sheet used as a template & demonstration, actual collection held elsewhere)

It uses [TCGPlayer](https://www.tcgplayer.com/)'s API to fetch information on the cards and their values.

Sheet setup
-----
To use this, a couple of things would need to be done. First, you need to apply and be approved for TCGPlayer's API, which can be done [here](https://developer.tcgplayer.com/developer-application-form.html). You will also need to create a new Google Sheet. You could either make a copy of my existing sheet or make your own. If you make your own sheet, you'll need to go to `Tools`, then `Script Editor` and create the all of the `.gs` files (except `APIToken.gs`) and paste in the code.

To copy the sheet, open the [sheet](https://docs.google.com/spreadsheets/d/1_zD-1VYKzvuM7W_ja9X6SWEYxPK_HNof1y0orEvWdWY/edit?usp=sharing), go to `File`, `Make a Copy` and choose a name and location.

Go to the `Card List` sheet and delete all entries except the header row. Open the script editor and delete the `CardLibrary` Library entry (Libraries are private, the one being copied won't work for you).

Setting up the API Key Library
-----
You will need to create a second sheet for the purpose of holding your API keys. Keep this sheet private. Open the script editor, and paste in the contents of `APITokens.gs` into the script and save. Once you have API approval, replace the public and private keys in `getNewAccessToken()`. Set the `sheetUrl` to the URL of the private sheet. Click `Deploy` in the top right corner, click `New Deployment`, the type should be `Library`. Give it a description, then click `Deploy`. Go to the settings cog on the left side of the screen, and copy the `Script ID`. 

Go back to your main sheet with the other four scripts. Click the `+` on `Libraries` on the left side panel, and paste in the Script ID you just copied, and click `Look Up`. Keep the version as `HEAD (Development Mode)` and make the identifier `CardLibrary`, then click `Add`.

Modifying your sheet
-----
The scripts are setup very specifically to work with my sheet layout. You could change the sheet layout if desired, you'd just have to change some of the code as well to reflect that. 

Using the sheet
-----
Once you have your sheet setup, here's how to actually use it. Go to the `Card List` sheet on the bottom. For each card, enter the card name, set, quantity, location, color identity, and type. For card games other than Magic: The Gathering these fields may need to be changed, which would mean changing the code. Once you've put in the card names and sets, go back to the `Overview` sheet and click the `Populate Product IDs` button to fill out the product IDs for each card. Depending on how many cards you have, this process may take a while. Then click the `Populate Card Info` button to get the collector number and rarity of each card. Finally, click the `Update Pricing` button to fill the value column. Once everything is finished, the `Update Statistics` button will update all of the statistics about the collection on the overview page. The `Sort Cards` button will sort the `Card List` sheet according to the order listed in `SortCards.gs`. Finally, `Update Filters` will copy cards that match the filter options into the `Filtered Cards` sheet.
