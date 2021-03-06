

 

/***Function: CopyComments***
----------------------------
*Copy comments from one sheet to the next
*/
function CopyCommentsLS(from,to){
for(var i=0; i<8;i++){
var source = from.getRange(12,4+i*6,22,3);/*E5 to F120*/
source.copyTo(to.getRange(12,4+i*6,22,3));
source.clear({contentsOnly:true});

}
}





function NewWeekLS() {
   
// Functions   
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var LSWeek1 = ss.getSheetByName("LS Week1");
   var LSWeek2 = ss.getSheetByName("LS Week2");
   var LSWeek3 = ss.getSheetByName("LS Week3");
   
//This section moves Available from SS Week2 to SS Week1   
 
 var AvailableTable = ss.getRange('LS Week2!A7:AV9');
 AvailableTable.copyTo(ss.getRange('LS Week1!A7:AV9'),{contentsOnly:true});
 
 CopyCommentsLS(LSWeek2,LSWeek1);
  

//This section moves only the Notes on the Header from SS Week3 to SS Week2   
  var source = ss.getRange('LS Week3!A7:AV9');
  source.copyTo(ss.getRange('LS Week2!A7:AV9'),{contentsOnly:true});

 
CopyCommentsLS(LSWeek3,LSWeek2);

}


