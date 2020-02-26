# Terraforming-Mars-Production
Google Sheets macro to track player incomes and run production phase

## Background
[Terraforming Mars](https://www.fryxgames.se/games/terraforming-mars/) is a board game for 1 to 5 players. A key component of the game involves strategically acquiring and spending resources. Using a spreadsheet to record stored resource income and income rates offers the following advantages:
* Players can easily see opponents' incomes helping to make strategic decisions. It can be difficult to precisely determine this publicly available information with resource cubes, especially when seated around a large table.
* Income rates are less likely to change by accident. Personal experience has shown how easy it is for players to accidentally knock the cubes out of place.
* Makes production phase faster. Moving resource cubes by hand is a slow process. This macro can run production phase in approximately half a second.

## How to get started
1. Set up a Google Sheets spreadsheet to your aesthetic. An example spreadsheet layout can be found in example-spreadsheet-layout/Sheet1.html. Be sure to include cells for each player's name, corporation, terraform rating (TR), and stored income and income rate for: M€, steel, titanium, plants, energy, and heat.
2. Open the script editor (Tools > Script editor). Overwrite the default code in macros.gs or code.gs with the version from this repo and save the changes.
3. Depending on how you have organized the spreadsheet, you will likely need to update the cell location values in the script. These follow a zero-based \[row, column\] format. This information is stored near the beginning of the script in: PLAYER, CORPORATION, TR, MC_PROD, MC_SUPPLY, STEEL_PROD, STEEL_SUPPLY, TITANIUM_PROD, TITANIUM_SUPPLY, PLANT_PROD, PLANT_SUPPLY, ENERGY_PROD, ENERGY_SUPPLY, HEAT_PROD, HEAT_SUPPLY. For example, if your spreadsheet has player names recorded in cells A1:A5, then `var PLAYER = [[7, 2],[31, 2],[55, 2],[79, 2],[103, 2]];` would change to: `var PLAYER = [[0, 0],[1, 0],[2, 0],[3, 0],[4, 0]];`.
4. From the script editor update the appsscript.json file (View > Show manifest file). You will likely want to add menu items for production phase and reset as well as a keyboard shortcut for production phase. After your changes appsscript.json may look like:
```
{
  "timeZone": "America/New_York",
  "dependencies": {
  },
  "exceptionLogging": "STACKDRIVER",
  "sheets": {
    "macros": [{
      "menuName": "production phase",
      "functionName": "production",
      "defaultShortcut": "Ctrl+Alt+Shift+1"
    }, {
      "menuName": "reset",
      "functionName": "reset"
    }]
  }
}
```
5. When running the macro for the first time you will be asked to grant permission for the script to run.
