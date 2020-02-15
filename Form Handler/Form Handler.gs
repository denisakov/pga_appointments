var calendarId = 'upqra3jhelbiegshuod1l0o19c@group.calendar.google.com';//Replace with actual Calendar ID.
var appointmentTrackerId = '1I9j_4hcWIdEzvp5huy6EiDDwg25J064xpDUxLaBUs5Y'; //Replace with actual Tracker sheet ID.
var formId = '1sM5yaAFmx1Q_28ry931RqpTD31aO6ED-tXZ52SHztAA';//Replace with actual form ID.

function handleForm(e){
  setAppointment(e);
  SpreadsheetApp.flush();
  updateFormEvents();
  return;
}
function setAppointment(e) {
  Logger.log(JSON.stringify(e));
  var values = e.namedValues;
  var name = values['Name'][0];
  var email = values['Email'][0];
  var eventName = values['Event'][0];
  
  var ss = SpreadsheetApp.openById(appointmentTrackerId);
  var s = ss.getSheetByName('list');
  var data = s.getDataRange().getValues();
  
  var eventNameInd = data[0].indexOf('Form Title');
  var eventIdInd = data[0].indexOf('EventId');
  var titleInd = data[0].indexOf('Title');
  var locIdInd = data[0].indexOf('Location');
  var startInd = data[0].indexOf('Start Date Time');
  var endInd = data[0].indexOf('End Date Time');
  for (var i = 1; i < data.length; i++){
    if(data[i][eventNameInd] == eventName){
      if(data[i][eventIdInd].length){
        var eventId = data[i][eventIdInd];
        break;
      }
    }
  }
  if(eventId){
    for (var i = 1; i < data.length; i++){
      if(data[i][eventNameInd] == eventName){
        if(!data[i][eventIdInd].length){
          var updated = updateEvent(eventId,email,name);
          if(updated){
            markTaken((i+1),(eventIdInd+1), eventId,s);
          }
          break;
        }
      }
    }
  } else {
    for (var i = 1; i < data.length; i++){
      if(data[i][eventNameInd] == eventName){
        eventId = creteNewEvent(data[i][titleInd], data[i][locIdInd], name, email, data[i][startInd], data[i][endInd]);
        markTaken((i+1),(eventIdInd+1), eventId,s);
        break;
      }
    }
  }
  return;
}
function markTaken(row, col, id, sheet){
  sheet.getRange(row,col).setValue(id);
  return;
}
function updateEvent(id,email,name){
  //id = 'ipst56ksajf5rqgjrvjv8ittkk@google.com';
  //name = 'Denis';
  //email = 'petsinyourhair@gmail.com';
  
  eventId = id.split('@')[0];
  try{
    var event = Calendar.Events.get(calendarId, eventId);
    var attendees = event.attendees;
    attendees.push({email: '"'+name+'" <'+email+'>'});
  
    var resource = { attendees: attendees };
    var args = { sendUpdates: "all" };
  
    Calendar.Events.patch(resource, calendarId, eventId, args);
    return true;
  } catch(e){
    Logger.log(e);
    return false;
  }
}
function creteNewEvent(title, loc, name, email, start, end){
  var event = CalendarApp.getCalendarById(calendarId)
      .createEvent(title,
               start,
               end,
               {location: loc,
                guests: '"'+name+'" <'+email+'>',
                sendInvites: true
               }
  );
  event.addEmailReminder(24*60);
  event.addEmailReminder(60);
  return event.getId();
}
function updateFormEvents() {
  var openSlots = getOpenSlots();
  if(openSlots.length == 0){
    Logger.log('No available events');
    openSlots = ["No available slots for the moment"];
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
/* Developed by dr.denisius@gmail.com */