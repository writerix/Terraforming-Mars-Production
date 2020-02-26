/** @OnlyCurrentDoc */

//these values are constant however the macro doesn't support 'const' properly at this time
//https://stackoverflow.com/questions/48293261/google-apps-script-redeclaration-of-const-error

var MAX_PLAYERS = 5;

//specify the cells where the values are stored using zero-based [row, column] format
//for example: cell C8 would be [7, 2]

var PLAYER = [[7, 2],[31, 2],[55, 2],[79, 2],[103, 2]];//['C8', 'C32', 'C56', 'C80', 'C104']
var CORPORATION = [[8, 2],[32, 2],[56, 2],[80, 2],[104, 2]];

var TR = [[9, 2],[33, 2],[57, 2],[81, 2],[105, 2]];

var MC_PROD = [[11, 2],[35, 2],[59, 2],[83, 2],[107, 2]];
var MC_SUPPLY = [[12, 2],[36, 2],[60, 2],[84, 2],[108, 2]];

var STEEL_PROD = [[14, 2],[38, 2],[62, 2],[86, 2],[110, 2]];
var STEEL_SUPPLY = [[15, 2],[39, 2],[63, 2],[87, 2],[111, 2]];

var TITANIUM_PROD = [[17, 2],[41, 2],[65, 2],[89, 2],[113, 2]];
var TITANIUM_SUPPLY = [[18, 2],[42, 2],[66, 2],[90, 2],[114, 2]];

var PLANT_PROD = [[20, 2],[44, 2],[68, 2],[92, 2],[116, 2]];
var PLANT_SUPPLY = [[21, 2],[45, 2],[69, 2],[93, 2],[117, 2]];

var ENERGY_PROD = [[23, 2],[47, 2],[71, 2],[95, 2],[119, 2]];
var ENERGY_SUPPLY = [[24, 2],[48, 2],[72, 2],[96, 2],[120, 2]];

var HEAT_PROD = [[26, 2],[50, 2],[74, 2],[98, 2],[122, 2]];
var HEAT_SUPPLY = [[27, 2],[51, 2],[75, 2],[99, 2],[123, 2]];

function reset(){
  
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetValues = spreadsheet.getDataRange().getValues();
  
  for(var i = 0; i < MAX_PLAYERS; i++){
    sheetValues[PLAYER[i][0]][PLAYER[i][1]] = '';
    sheetValues[CORPORATION[i][0]][CORPORATION[i][1]] = '';
    
    sheetValues[TR[i][0]][TR[i][1]] = 0;
    sheetValues[MC_PROD[i][0]][MC_PROD[i][1]] = 0;
    sheetValues[MC_SUPPLY[i][0]][MC_SUPPLY[i][1]] = 0;
    sheetValues[STEEL_PROD[i][0]][STEEL_PROD[i][1]] = 0;
    sheetValues[STEEL_SUPPLY[i][0]][STEEL_SUPPLY[i][1]] = 0;
    sheetValues[TITANIUM_PROD[i][0]][TITANIUM_PROD[i][1]] = 0;
    sheetValues[TITANIUM_SUPPLY[i][0]][TITANIUM_SUPPLY[i][1]] = 0;
    sheetValues[PLANT_PROD[i][0]][PLANT_PROD[i][1]] = 0;
    sheetValues[PLANT_SUPPLY[i][0]][PLANT_SUPPLY[i][1]] = 0;
    sheetValues[ENERGY_PROD[i][0]][ENERGY_PROD[i][1]] = 0;
    sheetValues[ENERGY_SUPPLY[i][0]][ENERGY_SUPPLY[i][1]] = 0;
    sheetValues[HEAT_PROD[i][0]][HEAT_PROD[i][1]] = 0;
    sheetValues[HEAT_SUPPLY[i][0]][HEAT_SUPPLY[i][1]] = 0;
  }
  spreadsheet.getDataRange().setValues(sheetValues);
}

function produceHeat(sv, i){
  var energyTransfer = parseInt(sv[ENERGY_SUPPLY[i][0]][ENERGY_SUPPLY[i][1]]);
  var heatProd = parseInt(sv[HEAT_PROD[i][0]][HEAT_PROD[i][1]]);
  var heatSupply = parseInt(sv[HEAT_SUPPLY[i][0]][HEAT_SUPPLY[i][1]]);
  
  sv[ENERGY_SUPPLY[i][0]][ENERGY_SUPPLY[i][1]] = 0;
  sv[HEAT_SUPPLY[i][0]][HEAT_SUPPLY[i][1]] = energyTransfer + heatProd + heatSupply;
}

function produceEnergy(sv, i){
  sv[ENERGY_SUPPLY[i][0]][ENERGY_SUPPLY[i][1]] = parseInt(sv[ENERGY_PROD[i][0]][ENERGY_PROD[i][1]]);
}

function producePlants(sv, i){
  sv[PLANT_SUPPLY[i][0]][PLANT_SUPPLY[i][1]] = parseInt(sv[PLANT_PROD[i][0]][PLANT_PROD[i][1]]) + parseInt(sv[PLANT_SUPPLY[i][0]][PLANT_SUPPLY[i][1]]);
}

function produceTitanium(sv, i){
  sv[TITANIUM_SUPPLY[i][0]][TITANIUM_SUPPLY[i][1]] = parseInt(sv[TITANIUM_PROD[i][0]][TITANIUM_PROD[i][1]]) + parseInt(sv[TITANIUM_SUPPLY[i][0]][TITANIUM_SUPPLY[i][1]]);
}

function produceSteel(sv, i){
  sv[STEEL_SUPPLY[i][0]][STEEL_SUPPLY[i][1]] = parseInt(sv[STEEL_PROD[i][0]][STEEL_PROD[i][1]]) + parseInt(sv[STEEL_SUPPLY[i][0]][STEEL_SUPPLY[i][1]]);
}

function produceMC(sv, i){
  sv[MC_SUPPLY[i][0]][MC_SUPPLY[i][1]] = parseInt(sv[TR[i][0]][TR[i][1]]) + parseInt(sv[MC_PROD[i][0]][MC_PROD[i][1]]) + parseInt(sv[MC_SUPPLY[i][0]][MC_SUPPLY[i][1]]);
}
/* Production phase calculations and sheet updates
*/
function production() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetValues = spreadsheet.getDataRange().getValues();
  
  for(var i = 0; i < MAX_PLAYERS; i++){
    produceHeat(sheetValues, i);
    produceEnergy(sheetValues, i);
    producePlants(sheetValues, i);
    produceTitanium(sheetValues, i);
    produceSteel(sheetValues, i);
    produceMC(sheetValues, i);
  }
  spreadsheet.getDataRange().setValues(sheetValues);
  
}