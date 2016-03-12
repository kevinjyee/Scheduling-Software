function FTE() {
  
  var blue = "#0066cc";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var B6Data = ss.getSheetByName("FTE");
  var maxColOffset = 31; 
  var firstRowInSched = 25;
  var firstColInSched = 2;
  var maxRowOffset = 136-firstRowInSched;

  
 
    var B6DataRange = B6Data.getRange(firstRowInSched,firstColInSched,maxRowOffset,maxColOffset);

  //input "Autoclave" and "Batch" into blue cells 
  var values = new Array(maxRowOffset); //multidimensional array for batch calls 
  for (var i = 0; i < maxRowOffset; i++) {

    values[i] = new Array(maxColOffset);
    for (var j = 0; j < maxColOffset; j++) {
      var cell = B6DataRange.getCell(i+1,j+1);
      if (cell.getBackground() == blue && B6DataRange.getCell(i+1,j+2).getBackground() == blue) {             
        values[i][j] = "Autoclave";
      } else if (cell.getBackground() == blue) {
        values[i][j] = "Batch";
      } else {
        values[i][j] = cell.getValue();
      }
    }
  }
  B6DataRange.setValues(values);
  
  
  //To be added
  //copy schedule to dashboard with right dates and stuff (need some kind of user input and limit it to number of days required for B6 data) 
//  var schedule = ss.getSheetByName("Live Schedule");
//  var maxCols = schedule.getMaxColumns();
//  var maxRows = schedule.getMaxRows();
//  var scheduleRange = schedule.getRange(1,1,schedule.getMaxRows(),schedule.getMaxRows());
//  scheduleRange.copyTo(B6DataRange);
  
  
}
