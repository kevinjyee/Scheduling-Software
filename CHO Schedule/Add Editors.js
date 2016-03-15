/*Function: addProtection
--------------------------
* email Addresses to add
*/

function addProtection(){
 
 var emailAddresses = [
  "yeek2@gene.com",
  "kevjyee@gmail.com",
  "tony@gene.com",
  "ongq@gene.com",
  ];
  
  addEditorsST(emailAddresses);
  addEditorsCHO(emailAddresses);
  }


/*Function: addEditorST
--------------------------
* Protect Range in the ST Region
*
*/

function addEditorsST(emailAddresses) {  
 // Protect range B:13:D32
 var ss = SpreadsheetApp.getActive();
 var SSWeek1 = ss.getSheetByName("SS Week1");
 var SSWeek2 = ss.getSheetByName("SS Week2");
 var SSWeek3 = ss.getSheetByName("SS Week3");
 
 for(var i=0; i<7;i++)
 {
 var range1 = SSWeek1.getRange(13,2+i*7,20,3);/*B13 to D32*/
 var protection1 = range1.protect().setDescription('Seed Train Protected Range');
 
 protection1.addEditors(emailAddresses);
  protection1.removeEditors(protection1.getEditors());
 
 
  var range2 = SSWeek2.getRange(13,2+i*7,20,3);/*B13 to D32*/
 var protection2 = range2.protect().setDescription('Seed Train Protected Range');
 
 protection2.addEditors(emailAddresses);
  protection2.removeEditors(protection2.getEditors());
 
  var range3 = SSWeek3.getRange(13,2+i*7,20,3);/*B13 to D32*/
 var protection3 = range3.protect().setDescription('Seed Train Protected Range');
 protection3.addEditors(emailAddresses);
  protection3.removeEditors(protection3.getEditors())
 }

}

/*Function: addEditorsCHO
--------------------------
* Protect Range in CHO Department
*
*/

function addEditorsCHO(emailAddresses)
{

var ss = SpreadsheetApp.getActive();
 var SSWeek1 = ss.getSheetByName("SS Week1");
 var SSWeek2 = ss.getSheetByName("SS Week2");
 var SSWeek3 = ss.getSheetByName("SS Week3");
 
 for(var i=0; i<7;i++)
 {
 var range1 = SSWeek1.getRange(42,2+i*7,74,2);/*B42 to C115*/
 var protection1 = range1.protect().setDescription('CHO Protected Range');
 protection1.addEditors(emailAddresses);
  protection1.removeEditors(protection1.getEditors());
 
var range2 = SSWeek1.getRange(42,2+i*7,74,2);/*B42 to C115*/
 var protection2 = range2.protect().setDescription('CHO Protected Range');

 protection2.addEditors(emailAddresses);
   protection2.removeEditors(protection2.getEditors());
 
var range3 = SSWeek1.getRange(42,2+i*7,74,2);/*B42 to C115*/
 var protection3 = range3.protect().setDescription('CHO Protected Range');

 protection3.addEditors(emailAddresses);
  protection3.removeEditors(protection3.getEditors());
 }


}



