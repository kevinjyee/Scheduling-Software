//converts all cells from relative to absolute ... will erase anything that is not a formula :/... must have '!' 
function switchRelAbs() {
  var range = SpreadsheetApp.getActiveRange();
  var values = range.getFormulas();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var formula = values[i][j];
      if (formula[0] == '=') {
        var result = "";
        var rowBool = 0;
        for (var k = 0; k < formula.length; k++) {
          if (rowBool == 1 && isNumeric(formula[k])) {
            rowBool = 0;
            result += '$';
          }
          result += formula[k];
          if (formula[k] == '!' || formula[k] == ':') {
            rowBool = 1;
            result += '$';
          } 
        }
        formula = result;
      }
      values[i][j] = formula
    }
  }
  range.setFormulas(values);
}

//checks to see if a string is a number
function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}
        
