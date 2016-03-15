function GenerateganttS() {
  /*
*Ideas: Copy and paste just the first two rows
*Run a remove duplicate function
*All duplicates are moved and then copied and pasted into new sheet
*Print dates
*Index if match else false
*/ 
  
var ss = SpreadsheetApp.getActive();
var datasheet = ss.getSheetByName("Gantt Chart Producer S");
  try{var dupsheet = ss.insertSheet("Remove Duplicates");
     }
  catch (e){
   var dupsheet = ss.getSheetByName("Remove Duplicates"); 
  }
var lastRow = datasheet.getLastRow();
var source = datasheet.getRange(1,2,lastRow,1);
  
var contentsin = dupsheet.getRange("A1:B100");
contentsin.clearContent();
  
source.copyTo(ss.getRange('Remove Duplicates!A3'));
//var source = datasheet.getRange(1,4,lastRow,1);
//source.copyTo(ss.getRange('Remove Duplicates!B3'));
  

  var data = ss.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  contentsin.clearContent();
  dupsheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);

  
 
  var select= ss.getSheetByName("Remove Duplicates");
  var range = select.getRange("Remove Duplicates!A3:B100");
  range.sort(1);//sorting messes up formula. go ahead and do this in another sheet
  var header = [["Scaledown"]]
  var range = dupsheet.getRange(2,1,1,1);
  
  range.setValues(header);
  range.setFontWeight("bold");  
  range.setFontSize(16);
  range.merge();
  range.setHorizontalAlignment("center");
  
  var Ganttsheet = ss.getSheetByName("Gantt Chart");
  var ScaleDownVesselsheet = ss.getSheetByName("Vessel");
  var source = dupsheet.getRange("A1:B100");
  var GVals = Ganttsheet.getRange("B1:B").getValues();
  var lastRowS= GVals.filter(String).length + 5;
  Logger.log(lastRowS);
  
  source.copyTo(Ganttsheet.getRange(lastRowS,1,1,1));
  source.copyTo(ScaleDownVesselsheet.getRange(lastRowS,1,1,1));
  ss.deleteSheet(dupsheet);
  return
}

function setDatesS(){
  Math
  var ss = SpreadsheetApp.getActive();
var datasheet = ss.getSheetByName("Gantt Chart Producer");
  
  var Gvals = datasheet.getRange("A1:A").getValues();
var lastRow = Gvals.filter(String).length;
 
  
 var range = datasheet.getDataRange();
 var firstcell = range.getCell(1,5).getValue();
  var secondcell = range.getCell(2,5).getValue();
 var lastcell = range.getCell(lastRow,5).getValue();
 var difference = Math.floor((lastcell - firstcell));

  var debug = secondcell - firstcell;
  Logger.log(debug);
  Logger.log(difference);
 

  var DateArray = new Array(difference);
  var i =0;
  var j =8.64E7;
  
  var sheet = ss.getSheetByName("Gantt Chart");
  sheet.activate();
  var clearrange = sheet.getRange("A1:BZ1");
  clearrange.clearContent()
  var range = sheet.getRange(1,3,1,difference+1);
  var firstvar = sheet.getRange(1,3).getValues();
  
  
  
  for(i=1;i<=difference+1;i++){
    
    range.getCell(1,i).setValue(firstcell+i);
    range.getCell(1,i).setNumberFormat("m/d");
//    
    DateArray[i] =firstcell + i; 
   //Logger.log(DateArray[i]);
     }
  range.setBorder(true, true, true,true, false, false)
  
  //range.setValues(DateArray);
    
    return
  }
