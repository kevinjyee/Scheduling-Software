function main(){
  export_gcal_to_gsheet()
  testReplaceInSheet()
  Generategantt()
  setDates()
  format()

}

function main2(){
  export_gcal_to_gsheetS()
  GenerateganttS()
}


function onOpen() {
 var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Get Data')
  .addItem('Purification Schedule', 'main')
    .addItem('Scaledown Schedule', 'main2')
           .addToUi();
  

           

}


