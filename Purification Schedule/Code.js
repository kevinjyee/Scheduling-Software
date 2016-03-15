function main(){
  export_gcal_to_gsheet()
  export_gcal_to_gsheetS()
  testReplaceInSheet()
  Generategantt()
  GenerateganttS()
  setDates()
  format()

}


function getCal(){
 export_gcal_to_gsheet();
  export_gcal_to_gsheetS();
}

function hide(){
 hideWeekends(); 
}


function show(){
 ShowColumns(); 
}

function onOpen() {
 var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Get Data')
  .addItem('Create Gantt', 'main')
  .addItem('Update PDMS Data','getCal')
   .addItem('Hide Weekends','hide')
  .addItem('Show Weekends','show')
            .addToUi();

           

}


