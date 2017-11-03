// TODO
// If error processing a mail, do retry and upon failure mark it for manual review (add tag to avoid further processing)

var schedulerLabel = "Tennis Scheduler"
var eventLocation = "<Location>"
var fromEmail = '<fromEmail>'
var myname = 'myname'

function createTennisCalEvent(title, startDate, endDate, location)
{
 var now = new Date();
 var existingEventId = null;
 if (startDate < now && endDate < now) {
   Logger.log('Past schedule; no event needed: '+title+" start: "+startDate+" end: "+endDate);
   return null;
 } else {
   var createdEvents = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate);
   for(var i=0; i<createdEvents.length; i++) {
     if (createdEvents[i].getTitle().equals(title) && createdEvents[i].getLocation().equals(location)) {
       existingEventId = createdEvents[i].getId();
       break;
     }
   }
 }
  
 if (existingEventId == null) {
   var event = CalendarApp.getDefaultCalendar().createEvent(title, startDate, endDate, {location: location});
   Logger.log('Created Event: '+title+" start: "+startDate+" end: "+endDate);
   Logger.log('Event ID: ' + event.getId());
   return event.getId();
 } else {
   Logger.log('Already existing event: '+title+" start: "+startDate+" end: "+endDate+" eventId: "+existingEventId);
   return existingEventId;
 }
}

function cancelTennisCalEvent(title, startDate, endDate, location)
{
 var now = new Date();
 if (startDate < now && endDate < now) {
   Logger.log('Past schedule; no need to cancel: '+title+" start: "+startDate+" end: "+endDate);
   return null;
 } else {
   var createdEvents = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate);
   for(var i=0; i<createdEvents.length; i++) {
     if (createdEvents[i].getTitle().equals(title) && createdEvents[i].getLocation().equals(location)) {
       var cancelEvent = createdEvents[i];
       Logger.log('Cancelling event: ' + cancelEvent.getId());
       cancelEvent.deleteEvent();
       return cancelEvent.getId();
     }
   }
 }
  
 Logger.log('Event not found to cancel: '+title+" start: "+startDate+" end: "+endDate);
 return null;
}

function processMailThreads(subject, fromDate, toDate, eventHandle) 
{
  var query = 'from:'+fromEmail+' subject:"'+subject+'" after:'+fromDate+' before:'+toDate+' has:nouserlabels';
  Logger.log('Searching threads for query: '+query);
  var threads = GmailApp.search(query);
  Logger.log("Found email threds : "+threads.length);
  var eventsHandled = 0;
  if (threads.length > 0) 
  {
    for(var i=0; i<threads.length; i++)
    {
      var thread = threads[i];
      
      // search for Tennis Scheduler label just in case
      var labels = thread.getLabels();
      var labelFound = false;
      for(var li=0; li<labels.length; li++) {
        if (labels[li].getName().equals(schedulerLabel)) {
          labelFound = true;
          break;
        }
      }
      
      // Process the email thread only when it was not labeled already
      if (!labelFound) {
        var body = thread.getMessages()[0].getBody();
        var court = body.match('Court .')[0];
        var ctime = body.match('\\D\\D\\D \\d\\d, \\d\\d\\d\\d \\d\\d:\\d\\d [A,P]M')[0];
      //  var cplayer1 = body.match('\\n\\s*The following court has been booked for you.\\s*(\\D*)\\s*vs')[1];
        var cplayer1 = body.match('\\D*The following court (you booked has been cancelled|has been booked for you).\\s*(\\D*)\\s*vs')[2];
        var cplayer2 = body.match('vs (\\D*)\\s(Mon|Tue|Wed|Thu|Fri|Sat|Sun)')[1];
        var sd = cplayer1.match("&");
        var courtTime = 1;
        var teamSep = " vs ";
        if (sd != null) {
          courtTime = 2;
        } else {
          if (cplayer1.equals(myname)) cplayer1 = "";
          if (cplayer2.equals(myname)) cplayer2 = "";
          teamSep = "";
        }
        var cdate1 = adjustTimeForDST(new Date(ctime+' UTC'));
        var cdate2 = new Date(cdate1.getTime());
        cdate2.setHours(cdate2.getHours() + courtTime);
        var title = "Tennis - "+cplayer1+teamSep+cplayer2+" ("+court+")";
        
        Logger.log('Handling event processing for subject: '+subject);
        var eventId = eventHandle(title, cdate1, cdate2, eventLocation);

        if (eventId != null) {
          eventsHandled++;
          
          // Add label
          var tennisSchedulerLabel = GmailApp.getUserLabelByName(schedulerLabel);
          thread.addLabel(tennisSchedulerLabel);          
        }

        // Do not archive it - manual review may be needed to confirm
        //GmailApp.moveThreadToArchive(thread);        
      } else {
        // Not expected to reach here as the search condition now has has:nouserlabels
        Logger.log('Email thread already scheduled: '+thread.getMessages()[0].getSubject());
      }
    } // for
  } // if
  return eventsHandled;
}

// To handle DST. The date objects are created initially with UTC, adjust the time using timezone offset
function adjustTimeForDST(date) {
//  Logger.log('DST before: '+date.getHours()+':'+date.getMinutes()+' dst offset:'+date.getTimezoneOffset());
  date.setTime( date.getTime() + date.getTimezoneOffset()*60*1000 );  
//  Logger.log('DST after: '+date.getHours()+':'+date.getMinutes());
  return date;
}

function myFunction() {
  var todayDate = new Date();
  var lastWeekDate = new Date();
  lastWeekDate.setDate(todayDate.getDate()-7);
  todayDate.setDate(todayDate.getDate()+1);
  
  var fd = lastWeekDate.getFullYear()+"/"+(lastWeekDate.getMonth()+1)+"/"+lastWeekDate.getDate();
  var td = todayDate.getFullYear()+"/"+(todayDate.getMonth()+1)+"/"+todayDate.getDate();
  
  var eventsCreated = processMailThreads("Court Booked", fd, td, createTennisCalEvent);
  Logger.log('Total number of events created: '+eventsCreated);

  var eventsCancelled = processMailThreads("Court Cancelled", fd, td, cancelTennisCalEvent);
  Logger.log('Total number of events cancelled: '+eventsCancelled);
}

