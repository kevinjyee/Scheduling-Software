/*Function: Generate Blocks with Defined Range
* ---------------------------
* This function automatically generates blocks based on inputs 
*/
 var blue = "#0066cc";
 var yellow = "#FFFF00"
 var orange = "#FFA500"
 
function GenerateBlocksRangeDefined(Name, ID, Amount, Days){
  var ui = SpreadsheetApp.getUi()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cell = ss.getActiveCell();
  var selectionArea = ss.getActiveRange();
  var area = ss.getActiveSheet();
  
  var firstRowindex = selectionArea.getRowIndex();
  var firstColumnindex = selectionArea.getColumn();
  var bluecolorRange = area.getRange(firstRowindex,firstColumnindex, Amount, 2);
  var yellowcolorRange = area.getRange(firstRowindex, firstColumnindex + 2, Amount, Days);
  var orangecolorRange = area.getRange(firstRowindex, Days+2, Amount, 1);
  
  
  var fullRange = area.getRange(firstRowindex, firstColumnindex, Amount, Days+2);
//  colorRange.setHorizontalAlignment("right");
//  colorRange.setBorder(true, true, true,true, true, true);
  var rowSize = selectionArea.getNumColumns();
  var columnSize = selectionArea.getNumRows();
  Logger.log(rowSize);
  Logger.log(columnSize);
  
  if(rowSize < Amount || columnSize < Days + 2){ui.alert("Selected Area does not have enough space")
  return};
  
  
  var Matrix = new Array(columnSize);
  
  for (var i = 0; i < columnSize; i++) 
  /*Create New Array*/
  {
  Matrix[i] = new Array(rowSize);
     for (var j = 0; j < rowSize; j++) {
     if(j == 1 && i < Amount){
      var num = i+1;
      var n = num.toString()
      Matrix[i][j] = n;
      }
     else if(j == 2 && i < Amount){
        Matrix[i][j] = "N"
        // use set Backgronds(color) later
      } else if (j == Days+2-1 && i < Amount){
        Matrix[i][j] = "H"
      } 
      else if(j > 1 && j < Days+1 && i < Amount){
      var num = j-2;
      
      var n = num.toString();
      Matrix[i][j] = "D" + n ; 
      }
      else{
          Matrix[i][j] = " ";
      }
    }
  }
 
 Matrix[0][0] = Name;
 Matrix[1][0] = ID;

 
 bluecolorRange.setBackground(blue);
 //bluecolorRange.setFontColor('white');
 yellowcolorRange.setBackground(yellow);
 orangecolorRange.setBackground(orange);

 ss.getActiveRange().setValues(Matrix);
 fullRange.setHorizontalAlignment("right");
 fullRange.setBorder(true,true,true,true,true,true);
return;
}




/*Function: Generate Block
* ---------------------------
* This function automatically generates blocks based on inputs 
*/
 var blue = "#0066cc";
 var yellow = "#FFFF00"
 var orange = "#FFA500"
 
function GenerateBlocks(Name, ID, Amount, Days){
  var ui = SpreadsheetApp.getUi()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cell = ss.getActiveCell();
  var selectionArea = ss.getActiveRange();
  var area = ss.getActiveSheet();
  
  var firstRowindex = selectionArea.getRowIndex();
  var firstColumnindex = selectionArea.getColumn();
  var bluecolorRange = area.getRange(firstRowindex,firstColumnindex, Amount, 2);
  var yellowcolorRange = area.getRange(firstRowindex, firstColumnindex + 2, Amount, Days);
 var orangecolorRange = area.getRange(firstRowindex, firstColumnindex + Days+1, Amount,1 );

var fullRange = area.getRange(firstRowindex, firstColumnindex, Amount, Days+2);
//  colorRange.setHorizontalAlignment("right");
//  colorRange.setBorder(true, true, true,true, true, true);
  var rowSize = selectionArea.getNumColumns();
  var columnSize = selectionArea.getNumRows();

  if(!fullRange.isBlank())
  {
    
  ui.alert("There is not enough space to schedule run number" + ID);
  return;
  }
  
  var Matrix = new Array(Amount);
  
  for (var i = 0; i < Amount; i++) 
  /*Create New Array*/
  {
  Matrix[i] = new Array(Days+2);
     for (var j = 0; j < Days+2; j++) {
     if(j == 1){
      var num = i+1;
      var n = num.toString()
      Matrix[i][j] = n;
      }
     else if(j == 2){
        Matrix[i][j] = "N"
        // use set Backgronds(color) later
      } else if (j == Days+1){
        Matrix[i][j] = "H"
      }
   
      else{
          var num = j-2;
      var n = num.toString();
      Matrix[i][j] = "D" + n ;
      }
    }
  }
 
 Matrix[0][0] = Name;
  if(Amount == 1){
    Matrix[0][1] = ID;
  }
  else{
 Matrix[1][0] = ID;
  }
 bluecolorRange.setBackground(blue);
  bluecolorRange.setFontColor("white");
 yellowcolorRange.setBackground(yellow);
 orangecolorRange.setBackground(orange);

   
 area.getRange(firstRowindex, firstColumnindex, Amount,Days+2).setValues(Matrix);
 fullRange.setHorizontalAlignment("right");
 fullRange.setBorder(true,true,true,true,true,true);
return;
}


/*Function: ReterieveRunData
----------------------------------
* Function Retrieves Run Data to be Scheduled
*/

function RetrieveRunData(){
var ui = SpreadsheetApp.getUi();
  try{
var RunNum = ui.prompt('Enter the PDMS run Number');
  }
  catch(e){
    
     Browser.msgBox("Failed"); 
  }
 // Process the user's response.
 if (RunNum.getSelectedButton() == ui.Button.OK) {
  
   var a = RunNum.getResponseText()
 } else {
   Logger.log('The user clicked the close button in the dialog\'s title bar.');
 }
   Logger.log(a);

return a;
}




/*Function: Old Main
------------------------------
* Main Driver for Auto Scheduler
*/

function OldMain(){

var ui = SpreadsheetApp.getUi();


//debugger;
var ss = SpreadsheetApp.getActiveSpreadsheet();
var toSchedule = ss.getSheetByName('Runs to Schedule');
var runNum = RetrieveRunData();


if(runNum != -1){
toSchedule.getRange("A2").setValue(runNum);
}

var Data = toSchedule.getRange("B2:D2").getValues();

Name = Data[0][0];
Quantity = Data[0][1];
Days = Data[0][2];
Logger.log(Name);
Logger.log(Quantity);
Logger.log(Days);
if(Name == "N/A"){OK
ui.alert("PDMS Number Not Found")
}
else{
GenerateBlocks(Name,runNum,Quantity,Days);
}

}


/*Find a <div> tag with the given id.
 * <pre>
 * Example: getDivById( html, 'tagVal' ) will find
 * 
 *          <div id="tagVal">
 * </pre>
 *
 * @param {Element|Document}
 *                     element     XML document or element to start search at.
 * @param {String}     id      HTML <div> id to find.
 *
 * @return {XmlElement}        First matching element (in doc order) or null.
 */
function getDivById( element, id ) {
  // Call utility function to do the work.
  return getElementByVal( element, 'div', 'id', id );
}

/**
 * !Now updated for XmlService!
 *
 * Traverse the given Xml Document or Element looking for a match.
 * Note: 'class' is stripped during parsing and cannot be used for
 * searching, I don't know why.
 * <pre>
 * Example: getElementByVal( body, 'input', 'value', 'Go' ); will find
 * 
 *          <input type="submit" name="btn" value="Go" id="btn" class="submit buttonGradient" />
 * </pre>
 *
 * @param {Element|Document}
 *                     element     XML document or element to start search at.
 * @param {String}     elementType XML element type, e.g. 'div' for <div>
 * @param {String}     attr        Attribute or Property to compare.
 * @param {String}     val         Search value to locate
 *
 * @return {Element}               First matching element (in doc order) or null.
 */
function getElementByVal( element, elementType, attr, val ) {
  // Get all descendants, in document order
  var descendants = element.getDescendants();
  for (var i =0; i < descendants.length; i++) {
    var elem = descendants[i];
    var type = elem.getType();
    // We'll only examine ELEMENTs
    if (type == XmlService.ContentTypes.ELEMENT) {
      var element = elem.asElement();
      var htmlTag = element.getName();
      if (htmlTag === elementType) {
        if (val === element.getAttribute(attr).getValue()) {
          return element;
        }
      }
    }
  }
  // No matches in document
  return null;
}



function myServerFunctionName(argForm) {
  var valueFromFormObj = JSON.stringify(argForm).slice(11,16);
 var ss= SpreadsheetApp.getActive();
  var sh = ss.getSheetByName("Runs to Schedule");

  var rng = sh.getRange("A2");

 rng.setValue(valueFromFormObj);
  
var ss = SpreadsheetApp.getActiveSpreadsheet();
var toSchedule = ss.getSheetByName('Runs to Schedule');
var Data = toSchedule.getRange("B2:D2").getValues();

var runNum = toSchedule.getRange("A2").getValue();
var Name = Data[0][0];
var Quantity = Data[0][1];
var Days = Data[0][2];

if(Name == "N/A"){
ui.alert("PDMS Number Not Found")
}
else{
  debugger;
GenerateBlocks(Name,runNum,Quantity,Days);
}

}
