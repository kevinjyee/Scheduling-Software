
/*****Function: Generategantt**********
* Generates Gantt Chart for Large Scale 
* Remove all Duplicates Vessels and Operation Types to condense Gantt Cart 
*/

function Generategantt() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var datasheet = ss.getSheetByName("Gantt Chart Producer");
  
  try{
    ss.getSheetByName("Remove Duplicates").activate();
  }
  catch(e){
    var dupsheet = ss.insertSheet("Remove Duplicates");
  }
  
var dupsheet = ss.getSheetByName("Remove Duplicates");
var lastRow = datasheet.getLastRow();
var source = datasheet.getRange(1,2,200,1);
var contentsin = dupsheet.getRange("A1:B100");
contentsin.clearContent();

  
source.copyTo(ss.getRange('Remove Duplicates!A3'));
var source = datasheet.getRange(1,4,200,1);
source.copyTo(ss.getRange('Remove Duplicates!B3'));
  

/*Find Duplicates*/  
/* .join() concats entire string to be searched*/
  var duplicatesheet = ss.getSheetByName("Remove Duplicates");
  var data = duplicatesheet.getDataRange().getValues();
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
  
/*Sets new contents*/
  contentsin.clearContent();
  dupsheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);

  
/*Sort the duplicate values just created*/
  
var select= ss.getSheetByName("Remove Duplicates");
var range = select.getRange("A3:B100");
range.sort(1);//sorting messes up formula. go ahead and do this in another sheet (Remove Duplicates Sheet)   
var header = [["Vessel", "Operation Type"]]
var range = dupsheet.getRange(1,1,1,2);
var range1 = dupsheet.getRange(2,1,1,2);
var largescaleheader = "Large Scale" 

range1.setValue(largescaleheader);
range1.merge();
range1.setFontSize(16);
range1.setFontWeight("bold");
range1.setHorizontalAlignment("center");
  
range.setValues(header);
range.setFontWeight("bold");
range.setFontSize(18);  
range.setHorizontalAlignment("center");
range.setHorizontalAlignment("center");
  
var source = dupsheet.getRange("A1:B100");
source.copyTo(ss.getRange('Gantt Chart!A3'));
source.copyTo(ss.getRange('Vessel!A3'));
ss.deleteSheet(dupsheet);
return
}

/*****Function: setDates**********
* Generate Dates as specified by the Large Scale Gantt Chart Producer
*/


function setDates(){
  Math
var ss = SpreadsheetApp.getActive();
var datasheet = ss.getSheetByName("Gantt Chart Producer");
  
//var lastRow = Gvals.filter(String).length; //commented out. Helpful to determine how many values are truly there.
var Gvals = datasheet.getRange("A1:A").getValues();
var lastRow = datasheet.getLastRow();  
var range = datasheet.getDataRange();
var firstcell = range.getCell(1,5).getValue();
Logger.log(firstcell);
var lastcell = range.getCell(200,5).getValue();
  
var difference = Math.floor((lastcell - firstcell)); //difference in dates from last cell and firstcell 


var DateArray = new Array(difference);//size of new array 
var i =0;
var j =8.64E7;
  
  
  var sheet = ss.getSheetByName("Gantt Chart");
  sheet.activate();
  var clearrange = sheet.getRange("A1:BZ1");
  clearrange.clearContent()
  var range = sheet.getRange(1,3,1,difference+1);
  var firstvar = sheet.getRange(1,3).getValues();
  
  
  var j =1;
  for(i=0;i<=difference+1;i++){
    if(j == difference+1){
    break;
    }
    try{
    range.getCell(1,j).setValue(firstcell+i);
    }
    catch(e){
    Logger.log(j);
    Logger.log(firstcell+i);
    }
    range.getCell(1,j).setNumberFormat("m/d");

    DateArray[j] =firstcell + i; 
    j++;
    
     }
  range.setBorder(true, true, true,true, false, false)
    
    return
  }

/*****Function: testReplaceInSheet***********
*Custom Find and Replace
*/
  
function testReplaceInSheet(){
  SpreadsheetApp
  var sheet = SpreadsheetApp.getActive().getSheetByName("Calendar Data"); //sheet to find 
  var replace = SpreadsheetApp.getActive().getSheetByName("Names Replacement"); //sheet to determine what to replace 
   sheet.activate();
  //custom force replacements to account for bad formatting
  replaceInSheet(sheet,'ops','OPS');
  replaceInSheet(sheet,'Ops','OPS');
  replaceInSheet(sheet,'Opr','OPS');
  replaceInSheet(sheet,'OPR', 'OPS');
  replaceInSheet(sheet,'opr','OPS');
  replaceInSheet(sheet,'OPS',' ');
  replaceInSheet(sheet,'"','');
  replaceInSheet(sheet,"Centrifuge",'Centrifuge (Centrifuge)');
  
  /*
  "	
 "	
"	
x 2	x2
x 3	x3
X2	x2
Poros XS 	PorosXS
CA	Capto Adhere
CaptoAdhere 	Capto Adhere
QS 	QSFF
X 6027	X6027
x7400	Homogenizer
Centrifuge	Centrifuge (Centrifuge)
centrifuge	Centrifuge (Centrifuge)
Filter Sizing	Filter Sizing (Filter Sizing)
filter sizing	Filter Sizing (Filter Sizing)
Poros 	PorosXS
"	
*/
  
  var i = 2;
  lastRow = replace.getLastRow()
  Logger.log(lastRow);
  Logger.log(replace.getRange(i,1).getValue());
  Logger.log(replace.getRange(i,2).getValue());
  for(i = 2; i <= lastRow; i++){
  replaceInSheet(sheet,replace.getRange(i,1).getValue(), replace.getRange(i,2).getValue());
  }
  
 //custom force replacements to account for bad formatting 
  replaceInSheet(sheet,'    ',' ');
  replaceInSheet(sheet,'   ',' ');
  replaceInSheet(sheet,'  ',' ');
 

  return
}

function replaceInSheet(sheet, to_replace, replace_with){
  //get the current data range values as an array
 var lastRow = sheet.getLastRow();
 var values = sheet.getRange(1,1,lastRow,2).getValues();
  Logger.log(lastRow);
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
  sheet.getRange(1,1,lastRow,2).setValues(values);
}


function format(){
var ss = SpreadsheetApp.getActive();
var gantt = ss.getSheetByName("Gantt Chart");
  gantt.getRange("A5:A").setHorizontalAlignment("center").setVerticalAlignment("middle");
  gantt.getRange("A5:A").setFontSize(14);
   gantt.getRange("B5:B").setHorizontalAlignment("center").setVerticalAlignment("middle");
  gantt.getRange("B5:B").setFontSize(14);
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
