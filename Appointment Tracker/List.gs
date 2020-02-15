//https://www.solodev.com/blog/web-design/adding-a-datetime-picker-to-your-forms.stml

var appointmentTrackerId = '1I9j_4hcWIdEzvp5huy6EiDDwg25J064xpDUxLaBUs5Y'; //Replace with actual sheet ID.
var formId = '1sM5yaAFmx1Q_28ry931RqpTD31aO6ED-tXZ52SHztAA';//Replace with actual form ID.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin')
  .addItem('Add Appointments', 'newEntryDialog')
  .addItem('Update Form','updateFormEvents')
  .addToUi();
}

function updateFormEvents(newSlots) {
  var openSlots = getOpenSlots();
  if(openSlots.length == 0){
    if(newSlots){
      openSlots = openSlots.concat(newSlots);
    } else {
      Logger.log('No available events');
      openSlots = ["No available slots for the moment"];
    }
  }
  var form = FormApp.openById(formId);
  var fields = form.getItems();
  for(var y = 0; y < fields.length; y++){
    var field = fields[y];
    if(field.getTitle() == 'Event'){
      field.asListItem().setChoiceValues(openSlots);
    }
  }
  Logger.log('Form is up to date');
  return;
}
function getOpenSlots(){
  var ss = SpreadsheetApp.openById(appointmentTrackerId);
  var s = ss.getSheetByName('list');
  var data = s.getDataRange().getValues();
  var eventNameInd = data[0].indexOf('Form Title');
  var eventIdInd = data[0].indexOf('EventId');
  var eventList = [];
  for(var x = 1; x < data.length; x++){
    if(!data[x][eventIdInd].length){
      eventList.push(data[x][eventNameInd]);
    }
  }
  return eventList;
}
function newEntryDialog() {
  var html = HtmlService.createHtmlOutputFromFile('form')
      .setWidth(400)
      .setHeight(700)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}
function addNewEntry(f) {
  Logger.log(f);
  var title = f.title;
  var location = f.loc;
  var slots = parseInt(f.slots);
  var start = new Date(f.start);
  var end = new Date(f.end);
  
  var formTitle = title + ' | ' + location + ' | ' + shortDate(start) + ' - ' + shortDate(end);
  
  Logger.log('FORM Title:' + formTitle);
  
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName('list');
  
  var newSlots = [];
  var newFormEntries = [];
  for(var i = 0; i < slots; i++){
    newSlots.push([title,location,start,end,formTitle]);
    newFormEntries.push(formTitle);
  }
  s.getRange(s.getLastRow()+1, 1, newSlots.length, newSlots[0].length).setValues(newSlots);
  updateFormEvents(newFormEntries);
  return;
}
function shortDate(d){
  return Utilities.formatDate(d, "America/New_York", "MM'/'dd'@'ha");
}
function chooseEdit(e){
  //Logger.log(Object.keys(e));
  if(e.authMode){
    Logger.log('AUTH MODE: ' + e.authMode);
    var range = e.range;
    Logger.log('range = ' + range);
    var column = range.getColumn();
    Logger.log('column = ' + column);
    var source = range.getSheet().getSheetName();
    Logger.log('source = ' + source);
    var value = e.range.getValue();
    Logger.log('value = ' + value);
    var oldValue = e.oldValue;
    Logger.log('oldValue = ' + oldValue);
  } else {
    var source = e.source;
    var row = e.row;
    var column = e.column;
  }
  if(source == 'list' && column == 6 && !value.length && oldValue.length){
    updateFormEvents();
  }
  return;
}
/* Developed by dr.denisius@gmail.com */