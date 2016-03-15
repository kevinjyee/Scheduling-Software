function Generategantt() {
  /*
*Ideas: Copy and paste just the first two rows
*Run a remove duplicate function
*All duplicates are moved and then copied and pasted into new sheet
*Print dates
*Index if match else false
*/ 
  
var ss = SpreadsheetApp.getActive();
var datasheet = ss.getSheetByName("Gantt Chart Producer");
var dupsheet = ss.insertSheet("Remove Duplicates");
var lastRow = datasheet.getLastRow();
var source = datasheet.getRange(1,2,lastRow,1);
  
var contentsin = dupsheet.getRange("A1:B100");
contentsin.clearContent();
  
source.copyTo(ss.getRange('Remove Duplicates!A3'));
var source = datasheet.getRange(1,4,lastRow,1);
source.copyTo(ss.getRange('Remove Duplicates!B3'));
  

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
var header = [["Vessel", "Operation Type"]]
var range = dupsheet.getRange(2,1,1,2);
range.setValues(header);
range.setFontWeight("bold");  
  
  var source = dupsheet.getRange("A1:B100");
  source.copyTo(ss.getRange('Gantt Chart!A3'));
  source.copyTo(ss.getRange('Vessel!A3'));
  ss.deleteSheet(dupsheet);
  return
}



function setDates(){
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
  
function testReplaceInSheet(){
  SpreadsheetApp
    var sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
    var replace = SpreadsheetApp.getActive().getSheetByName("Names Replacement");
   
  replaceInSheet(sheet,'ops','OPS');
  replaceInSheet(sheet,'Ops','OPS');
  replaceInSheet(sheet,'Opr','OPS');
  replaceInSheet(sheet,'OPR', 'OPS');
  replaceInSheet(sheet,'opr','OPS');
  replaceInSheet(sheet,'OPS',' ');
   replaceInSheet(sheet,replace.getRange("A2").getValue(), replace.getRange("B2").getValue());
   replaceInSheet(sheet,replace.getRange("A3").getValue(), replace.getRange("B3").getValue());
   replaceInSheet(sheet,replace.getRange("A4").getValue(), replace.getRange("B4").getValue());
  replaceInSheet(sheet,replace.getRange("A5").getValue(), replace.getRange("B5").getValue());
   replaceInSheet(sheet,replace.getRange("A6").getValue(), replace.getRange("B6").getValue());
   replaceInSheet(sheet,replace.getRange("A7").getValue(), replace.getRange("B7").getValue());
   replaceInSheet(sheet,replace.getRange("A8").getValue(), replace.getRange("B8").getValue());
   replaceInSheet(sheet,replace.getRange("A9").getValue(), replace.getRange("B9").getValue());
   replaceInSheet(sheet,replace.getRange("A10").getValue(), replace.getRange("B10").getValue());
   replaceInSheet(sheet,replace.getRange("A11").getValue(), replace.getRange("B11").getValue());
   replaceInSheet(sheet,replace.getRange("A12").getValue(), replace.getRange("B12").getValue());
  replaceInSheet(sheet,replace.getRange("A13").getValue(), replace.getRange("B13").getValue());
  replaceInSheet(sheet,replace.getRange("A14").getValue(), replace.getRange("B14").getValue());
  replaceInSheet(sheet,replace.getRange("A15").getValue(), replace.getRange("B15").getValue());
  replaceInSheet(sheet,replace.getRange("A16").getValue(), replace.getRange("B16").getValue());
   replaceInSheet(sheet,replace.getRange("A17").getValue(), replace.getRange("B17").getValue());
   replaceInSheet(sheet,replace.getRange("A18").getValue(), replace.getRange("B18").getValue());
 replaceInSheet(sheet,replace.getRange("A19").getValue(), replace.getRange("B19").getValue());
   replaceInSheet(sheet,replace.getRange("A20").getValue(), replace.getRange("B20").getValue());
     replaceInSheet(sheet,replace.getRange("A21").getValue(), replace.getRange("B21").getValue());
      replaceInSheet(sheet,replace.getRange("A22").getValue(), replace.getRange("B22").getValue());
     replaceInSheet(sheet,replace.getRange("A23").getValue(), replace.getRange("B23").getValue());
     replaceInSheet(sheet,replace.getRange("A24").getValue(), replace.getRange("B24").getValue());
     replaceInSheet(sheet,replace.getRange("A25").getValue(), replace.getRange("B25").getValue());
     replaceInSheet(sheet,replace.getRange("A26").getValue(), replace.getRange("B26").getValue());
  //  replaceInSheet(sheet,'x 2','');
//  replaceInSheet(sheet,'X2','');
//  replaceInSheet(sheet,'x2','');
//  replaceInSheet(sheet,'XS','');
//  replaceInSheet(sheet,'x3','');
  
  replaceInSheet(sheet,'    ',' ');
  replaceInSheet(sheet,'   ',' ');
  replaceInSheet(sheet,'  ',' ');
 

  return
}

function replaceInSheet(sheet, to_replace, replace_with){
  //get the current data range values as an array
 var values = sheet.getRange("A1:B200").getValues();

  //loop over the rows in the array
  for(var row in values){

    //use Array.map to execute a replace call on each of the cells in the row.
    var replaced_values = values[row].map(function(original_value){
      return original_value.toString().replace(to_replace,replace_with);
    });

    //replace the original row values with the replaced values
    values[row] = replaced_values;
  }

  //write the updated values to the sheet
  sheet.getRange("A1:B200").setValues(values);
}


function format(){
var ss = SpreadsheetApp.getActive();
var gantt = ss.getSheetByName("Gantt Chart");
  var range = gantt.getRange("A5:A").getValues();
  var numberofvessels = range.filter(String).length;
  Logger.log(numberofvessels);
  
  var rangeofA = gantt.getRange(5,1,numberofvessels);
 // rangeofA.activate();
  
  //Logger.log("First Vessel should be x1030");
    
  var firstvessel = rangeofA.getCell(1,1).getValue();
  var arrayofposition = [];
  var j=1;
  
  arrayofposition[0]=1;
  Logger.log(firstvessel);
  
  for(var i =1; i<numberofvessels+1; i++){
    
    if(rangeofA.getCell(i,1).getValue() != firstvessel){
      
      var firstvessel = rangeofA.getCell(i,1).getValue();
      arrayofposition[j] = i;
      j++;
        }
  }
    
   
    
   
 Logger.log(arrayofposition[0]);//10
  Logger.log(arrayofposition[1]);//15
  Logger.log(arrayofposition[2]);//16
  Logger.log(arrayofposition[3]);//23
  Logger.log(arrayofposition[4]);//24
  
  var count =0;
  var differnece=0;
  while(difference == 1){
    var difference= arrayofposition[count+1] - arrayofposition[count];
    count++;
  }
  
  var j =1;
  
  for(var i=0; i<arrayofposition.length-1; i++){
    gantt.getRange(count+4+arrayofposition[i],1,arrayofposition[j] - arrayofposition[i]).merge();
    gantt.getRange(count+4+arrayofposition[i],2,arrayofposition[j] - arrayofposition[i]).sort(2);
    
    j++;
    
    
    
  }
  
  gantt.getRange("A5:A").setHorizontalAlignment("center").setVerticalAlignment("middle");
                             
}








/*
//incrementer for everything 
  var i =0;  
//Length of Column B 
var ss = SpreadsheetApp.getActive();
var datasheet = ss.getSheetByName("Gantt Chart Producer");
var lastRow = datasheet.getLastRow();
Logger.log(lastRow);
//Length of Column G 

var Gvals = datasheet.getRange("G1:G").getValues();
var numberofvessels = Gvals.filter(String).length;
Logger.log(numberofvessels);
  
//Put all the values of Column G into an array
var rangeofG = datasheet.getRange(1,7,numberofvessels,1);
rangeofG.activate();
  var myArray = [];
  for(i =1; i<=numberofvessels; i++){
    myArray[i] = rangeofG.getCell(i,1).getValue();
  }
  
//Put all the values of Column B into an array 

var rangeofB = datasheet.getRange(1,2,lastRow,1);
rangeofB.activate();
var x = [];
for(i=1; i<lastRow; i++){ 
x[i] =  rangeofB.getCell(i,1).getValue();
 
*/
