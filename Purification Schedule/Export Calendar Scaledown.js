function export_gcal_to_gsheetS(){

//
// Export Google Calendar Events to a Google Spreadsheet
//
// This code retrieves events between 2 dates for the specified calendar.
// It logs the results in the current spreadsheet starting at cell A2 listing the events,
// dates/times, etc and even calculates event duration (via creating formulas in the spreadsheet) and formats the values.
//
// I do re-write the spreadsheet header in Row 1 with every run, as I found it faster to delete then entire sheet content,
// change my parameters, and re-run my exports versus trying to save the header row manually...so be sure if you change
// any code, you keep the header in agreement for readability!
//
// 1. Please modify the value for mycal to be YOUR calendar email address or one visible on your MY Calendars section of your Google Calendar
// 2. Please modify the values for events to be the date/time range you want and any search parameters to find or omit calendar entires
// Note: Events can be easily filtered out/deleted once exported from the calendar
// 
// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event
//

 var cal = CalendarApp.getCalendarById('gene.com_qckdld8rb6n421vr2t4o4kosd4@group.calendar.google.com')
 //var cal = CalendarApp.getCalendarByID('gene.com_jh8judd7967q602tgq50aavbrg%40group.calendar.google.com')
 var ui = SpreadsheetApp.getUi();
// var Start = ui.prompt('Enter the Start Date:');
var ss = SpreadsheetApp.getActive();
var dateSheet = ss.getSheetByName("Dates");

  
var a = dateSheet.getRange("A2").getValue();
var b = dateSheet.getRange("B2").getValue();
 // Process the user's response.
// if (Start.getSelectedButton() == ui.Button.OK) {
//  
//    a = Start.getResponseText()
// } else {
//   Logger.log('The user clicked the close button in the dialog\'s title bar.');
// }
//   Logger.log(a);
//  
//
//
//var End = ui.prompt('Enter the End Date:');
//if (End.getSelectedButton() == ui.Button.OK) {
//  
//    b = End.getResponseText()
// } else {
//   Logger.log('The user clicked the close button in the dialog\'s title bar.');
// }
//  Logger.log(a);
//   Logger.log(b);
  


// Optional variations on getEvents
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"));
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
// 
// Explanation of how the search section works (as it is NOT quite like most things Google) as part of the getEvents function:
//    {search: 'word1'}              Search for events with word1
//    {search: '-word1'}             Search for events without word1
//    {search: 'word1 word2'}        Search for events with word2 ONLY
//    {search: 'word1-word2'}        Search for events with ????
//    {search: 'word1 -word2'}       Search for events without word2
//    {search: 'word1+word2'}        Search for events with word1 AND word2
//    {search: 'word1+-word2'}       Search for events with word1 AND without word2
//
//var events = cal.getEvents(new Date('June 1, 2015'), new Date('July 1, 2015'));
//var today = new Date();
 //var events = CalendarApp.getDefaultCalendar().getEventsForDay(today, {search: 'meeting'});
 //Logger.log('Number of events: ' + events.length);

 // Logger.log(startdate);
  //Logger.log(enddate);
 // var events = cal.getEvents(new Date(a), new Date(b),{search: 'x1538'}); // array of values ranging dates a and b
    var events1 = cal.getEvents(new Date(a), new Date(b))
  var events2 = cal.getEvents(new Date(a), new Date(b),{search: 'opr*'})
  var events3 =cal.getEvents(new Date(a), new Date(b),{search: 'filter'})
  var events4= cal.getEvents(new Date(a), new Date(b),{search: 'centrifuge*'})
 // cal.getEvents(new Date(a), new Date(b),{search: 'Assembly'})
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar Data Scaledown");
// Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
sheet.clearContents();  

// Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
var header = [["Event Title", "Event Description", "Event Start"]]
var range = sheet.getRange(1,1,1,3);
range.setValues(header);

var length1=0;
// Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
for (var i=0;i<events1.length;i++) {
var row=i+2; // starts at row 2 
var myformula_placeholder = '';
//// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
//// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
var details=[[events1[i].getTitle(), events1[i].getDescription(),events1[i].getStartTime()]];
var range=sheet.getRange(row,1,1,3);
range.setValues(details);
length1++;
// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
  
//  Uncomment this code if you want to calculate duration as well
//var cell=sheet.getRange(row,7);
//cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
//cell.setNumberFormat('.00');

  
}
  
   return
}
 



