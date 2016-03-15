/***Function:hideWeekends***
----------------------------
*Hides the weekend dates
*/


function hideWeekends() {
  
  var startcolumn = 3;
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var Ganttsheet = ss.getSheetByName("Gantt Chart");
  var compareString = "";
  var toFind = "Sat";
  var lastColumn = Ganttsheet.getLastRow();
  
  while(true){
   compareString = Ganttsheet.getRange(2,startcolumn).getValue(); 
  
    if(compareString.localeCompare(toFind) == 0){
      break;
    }
    startcolumn++;
  }
  
  Logger.log(startcolumn);
 
  
  
  for(var i =0; i <= lastColumn; i+=7){
 var toHide = Ganttsheet.getRange(2,startcolumn+i);
 Ganttsheet.hideColumn(toHide);
   
  }
  
  for(var i = 1; i<=lastColumn;i+=7){
    var toHide = Ganttsheet.getRange(2,startcolumn + i);
    Ganttsheet.hideColumn(toHide);
}
}


function ShowColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Gantt Chart");
  var range = sheet.getRange("1:1");
  sheet.unhideColumn(range);
}
