//Row Info
var tankIndices = new Array(); //Global Array to make inputs more flexible and dynamic

var rowStart; //Top most row to place first block 
var rowEnd;//Bottom most bound

//Column Info
var matrixBegin; //Begining Start Date 
var matrixEnd; // Beginning End Date
var baseDate; // Date that the Schedule Starts on 
var dateOffset; // Column number that the baseDate is located

//Submission Form Info:

var requestIndex = 1; //Column Index of Request Number
var proteinName = 2; //Column Index of Protein Name
var volIndex = 4; // Column Index of Volume
var startDateIndex = 10; // Column Index of Start Date 
var endDateIndex = 11; // Column Index of End Date 
var scheduledIndex = 14; // Column Index of Scheduled Or Not 


//Data Range Length
var dataRangeLength; //length of the queue of array that needs to be scheduled

var data = new Array(); // Static Queue that holds everything that needs to be scheduled


var msPerYear = 24*60*60*1000; //msPerYear to calculate column offsets

//unused misc variables
var index =0;
var offsetPrimary;

/*Function: completeAutomation()
---------------------------------
@ This is the main function for automating the schedule
@ driver to run automative schedule
@ params: none
@ return: none
*/

function completeAutomation()
{ 
  
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var area = ss.getSheetByName("Live Schedule");
 
 var fullRange = setupMatrix();// fullRange holds the entire Column X Row matrix that the schedule will be generated on
 queueSchedule(); // Puts the entire schedule into the data queue
 var completedRange = automateGenerate(fullRange); //takes fullRange as a parameter to generate the schedule on
 
  var maxColOffset = matrixEnd - matrixBegin +2;
  var maxRowOffset = rowEnd - rowStart;
  var scheduleRange = area.getRange(rowStart,matrixBegin,maxRowOffset,maxColOffset);
  
  scheduleRange.setValues(completedRange);
}
 
 /*Function: setupColumns()
 --------------------------------
@setup dates (column indices that the schedule will be scheduled under)
*/
function setupColumns(){
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var controls = ss.getSheetByName("Scheduling Controls");
 var startendInfo = controls.getRange("F2:G4").getValues(); //date info is located in this range. Currently not dynamic

  prebaseDate = new Date(startendInfo[0][0]); //Base Date (Usually January 1st) 
  baseDate = prebaseDate.getTime();
  prematrixBegin = new Date(startendInfo[1][0]); //Matrix Begin Date (usually the min of start date column)
  matrixBegin = Math.floor((prematrixBegin.getTime()-baseDate)/msPerYear) ;
  prematrixEnd = new Date(startendInfo[2][0]) ; //Matrix End Date (usually the max of end date column) 
  matrixEnd = Math.floor((prematrixEnd.getTime()-baseDate)/msPerYear);
  
  dateOffset = startendInfo[1][1];

}
 
 
 /*Function: setupRows()
 --------------------------------
@ setup the row index (Row Start of Each Vessels)
*/

function setupRows(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var controls = ss.getSheetByName("Scheduling Controls");
  
  //get data range of each row
  var columnwithVessels = controls.getRange("A1:A").getValues();
  var numberUnique = columnwithVessels.filter(String).length;
  Logger.log("Number Unique");
  Logger.log(numberUnique);
  var rowInfo = controls.getRange(2,2,numberUnique,2).getValues(); //Data Range possible for rows 
  
  rowStart = rowInfo[0][0]; //index of where the row starts
  
  rowEnd = rowInfo[numberUnique-1][1]; // index of where the row ends
  
  Logger.log("Start and End");
  Logger.log(rowStart);
  Logger.log(rowEnd);
  for(var i =0; i < numberUnique; ++i){
  tankIndices[i] = new Array();
  tankIndices[i][0] = rowInfo[i][0]
  tankIndices[i][1] = rowInfo[i][1]
  }
  
  //setupInfo to be scheduled

}
 
/*Function: setupMatrix()
---------------------------------
@ setup matrix to begin scheduling
@ return: dataMatrix Row X Column matrix range to be scheduled
*/
 
function setupMatrix()
{
 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var controls = ss.getSheetByName("Scheduling Controls");
  
  //setupTankRows for each colums
  setupRows();
  //setupDates to be used in Column
  setupColumns();
  
  
  var area = ss.getSheetByName("Live Schedule");
  
  
  var maxColOffset = matrixEnd - matrixBegin+2;
  var maxRowOffset = rowEnd - rowStart;
  
  var fullRange = area.getRange(rowStart,matrixBegin,maxRowOffset,maxColOffset);
  var fullRangeData = fullRange.getValues();
 
  
  var values = new Array(maxRowOffset); //multidimensional array for batch calls 
  for (var i = 0; i < maxRowOffset; i++) {
 
    values[i] = new Array(maxColOffset);
    for (var j = 0; j < maxColOffset; j++) {
      var cell = fullRange.getCell(i+1,j+1);
        values[i][j] = fullRangeData[i][j];
      }
    }
 
  return values;
}
 




/*Function: queueSchedule()
---------------------------------
@ Params None
@ Places all data needed into the Queue
*/
function queueSchedule()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputs = ss.getSheetByName("Schedule Input");
  var dataRange = inputs.getDataRange().getValues();
   dataRangeLength = dataRange.length;
  
  
  for(var i = 0; i < dataRange.length-1; i++)
  {
    
   
    if(dataRange[i+1][scheduledIndex -1] == "" || dataRange[i+1][scheduledIndex =1] == null){
      
    data.push(getSheetData(inputs,i+2));
    
    }
    
  }
  
  
   
  
  
}


/*Function: getSheetData()
------------------------------------
@Params row and sheet to search
@
*/

function getSheetData(sheet,row)
{
  
 
 
var sheetVals = sheet.getRange(row,1,1,14).getValues();
  
  var dataArray = new Array();
  
  dataArray.req = sheetVals[0][requestIndex-1];
  dataArray.name = sheetVals[0][proteinName-1];
  dataArray.volume = sheetVals[0][volIndex-1];
  dataArray.startDate = sheetVals[0][startDateIndex-1];
  dataArray.endDate = sheetVals[0][endDateIndex-1];
  dataArray.scheduled = sheetVals[0][scheduledIndex-1];
 
  
// data.req = sheet.getRange(row,requestIndex,1,1).getValue();
// data.name =sheet.getRange(row,proteinName,1,1).getValue();
// data.volume = sheet.getRange(row,volIndex,1,1).getValue();
// data.startDate = sheet.getRange(row,startDateIndex,1,1).getValue();
// data.endDate = sheet.getRange(row,endDateIndex,1,1).getValue();
// data.scheduled = sheet.getRange(row,scheduledIndex,1,1).getValue();
// 

 return dataArray; 
}




/*Function: automateGenerate()
-----------------------------------
@ Params Queue to schedule
@ Automatically schedules everything in the Queue

*/


function automateGenerate(map)
{
  var temp; 
  var sIndex;
  var eIndex;
  
  //dequeu what needs to be scheduled
  
  while(index < dataRangeLength-1 )
  {
    
    //find index needed to be mapped in the matrix
    
    sIndex = new Date(data[index].startDate)
    sIndex = Math.floor((sIndex.getTime()-baseDate)/msPerYear);// - baseDate + dateOffset;
    eIndex = new Date(data[index].endDate)
    eIndex = Math.floor((eIndex.getTime()-baseDate)/msPerYear);// - baseDate + dateOffset;
   
    //Schedule 5L 
    if(data[index].volume >= 5 && data[index].volume < 10)
    {
      
    var i = rowStart;
      Logger.log("Debug");
      Logger.log(i);
      Logger.log(eIndex);
      Logger.log(map[i][sIndex - matrixBegin +2]);
      Logger.log(map[i][eIndex - matrixBegin +2]);
      Logger.log(tankIndices[0][1]);
      while(i < tankIndices[0][1])
      {
        if(rowFits(i,map,sIndex,eIndex))
        {
          break; 
        }
        i++;
      }
   
      if(i < tankIndices[0][1]){
        map = placeBlocks(sIndex, eIndex, i, map);
        i++;
      }   
      
    }
    
    else if(data[index].volume >=10 && data[index].volume < 35)
    {
    var i = tankIndices[1][0]
      while(i < tankIndices[1][1])
      {
        if(rowFits(i,map,sIndex,eIndex))
        {
          
          break;
          
        }
        i++;
      }
      
      if(i < tankIndices[1][1]){
        map = placeBlocks(sIndex, eIndex, i, map)
        i++;
      }
      
   
  }
  
      else if(data[index].volume >=35 && data[index].volume < 200)
    {
      
    var i = tankIndices[2][0];
    
      while(i < tankIndices[2][1])
      {
        if(rowFits(i,map))
        {
          
          break;
          
        }
        i++;
      }
      
      if(i < tankIndices[2][1]){
        map = placeBlocks(sIndex, eIndex, i, map);
        i++;
      }
      
   
  }
  
  index++;
  
}
  return map;
}

/*Function: isEmpty
------------------------
@ Params Checks if the queue is empty

*/
function isEmpty(arraytype)
{
  if(arraytype.length <1){
    return true;
  }
  else{
    return false;
}
}



/* Function: placeBlocks
-------------------------------
@ Params StartDateIndex
@ Params EndDateIndex
@ Params Map
@ Return Updated Map Value

*/


function placeBlocks(sIndex, eIndex, row,map)
{
  
 
  var i = sIndex-matrixBegin+2;
  var j = eIndex-matrixBegin+2;
  
  row = row - rowStart;
  if(j-i < 2){return map;}
  
  Logger.log("Begin:");
  Logger.log(matrixBegin);
  Logger.log(matrixEnd);
  Logger.log(i);
  Logger.log(j);
  Logger.log(row);
  Logger.log("End");
  
   map[row][i] = data[index].req;
  map[row][i+1] = data[index].name;
  i +=2;
  
  var k =1;
  while(i < j)
      {
        map[row][i] = "D" + k;
        i++;
        k++;
      }
  map[row][i] = "H";
  return map;
}
  
  /*Function: itFts()
  ---------------------------
  *Determines if the row is availble to place blocks in 
  */
  
  function rowFits(row,map,sIndex,eIndex)
  {
    if((map[row-rowStart][sIndex - matrixBegin +2] == "" || 
        map[row-rowStart][sIndex - matrixBegin +2] == null) && 
       (map[row-rowStart][eIndex - matrixBegin + 2] == "" || 
       map[row-rowStart][eIndex - matrixBegin+2] == null)){
      
     return true; 
    }

    else false;
  }
