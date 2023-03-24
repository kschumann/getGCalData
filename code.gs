function  getUserData(){
  try{
    const presets = getPresets();
    const calendars = CalendarApp.getAllCalendars();
    const numbCalendars = calendars.length;
    var calendarList = [];
    for(let i=0;i<numbCalendars;i++){
      calendarList.push([calendars[i].getId(),calendars[i].getName()])
    }  
    return [presets,calendarList];
    
  } catch(e){
    console.log("Unhandled Exception for getUserCalendars: " + e);
  }
}

function getCalData(calIds,start,end,keyWordInput,selectedColumns,type){
  try{
    const spreadsheet = SpreadsheetApp.getActive();
    const sheetId = spreadsheet.getId();
    const scriptStart = new Date();

    const calAttr = [
      ["calName","Calendar Name",250],
      ["eventName","Event Name",300],
      ["eventStart","Event Start",100],
      ["eventEnd","Event End",100],
      ["location","Event Location",400],
      ["eventColor","Event Color",100], 
      ["creator","Event Creator",300],
      ["participant","Event Invitees",300],
      ["description","Event Description",500],
      ["hoursDuration","Length (Hours)",100],
      ["dateCreated","Date Created",100],
      ["lastUpdated","Last Updated",100],
      ["visibility","Visibility",100],
      ["numbParticipants","Total Invitees",100],  
      ["isRecurringEvent","Is Recurring Event",100],
      ["calendarTimezone","Calendar Timezone",150]
    ];
    
    /**Create single array of column IDs**/
    var calAttrList = [];
    for(let i=0;i<calAttr.length;i++){
      calAttrList.push(calAttr[i][0]);
    }

  /**Create an array of indices that should be used */
    var selectedIndices = [];
    for(let i=0;i<selectedColumns.length;i++){
      selectedIndices.push(calAttrList.indexOf(selectedColumns[i]));
    }  
  //console.log(selectedIndices);  


    var calData = [];//eventName,staffEmail, eventDate,eventDuration  
    var noEventsList = [];
   // console.log("Calendars running for type: " + type);
  /**Populate calData array **/
    for(let m=0;m<calIds.length;m++){//iterate through calendars
      let calendar = CalendarApp.getCalendarById(calIds[m]);
      var calendarTimezone = calendar.getTimeZone(); //need to get events in Calendar's timezone
      //var calendarLocale = calendar.getTimeZone().toLocaleLowerCase(); //Not sure why calendar locale would be needed
      var newDate = new Date();
      var timezoneOffset = getTimeZoneOffset(newDate, calendarTimezone);  
      //var start = Utilities.formatDate(new Date(start), calendarTimezone, "YYYY-MM-dd hh:mm:ss a");

      /**This seems to be working now using adjusted getTimeZoneOffset, but keeping fallback in just in case.**/
      if(isNaN(timezoneOffset)){
        start = new Date(start); 
        end = new Date(end+24*60*60*1000); 
        console.warn(calendarTimezone + ": Timezone Offset Not Calculated")      
      } else{
        start = new Date(start+timezoneOffset); 
        end = new Date(end+24*60*60*1000+timezoneOffset);
      }
    console.log(sheetId + " -- Extrated started!  Calendars: " + calIds + "; Start Date: " + start + "; End Date: " + end + "; selectedColumns:" + selectedColumns + "; keyword: " + keyWordInput + "; Report Type: " + type);      
      //var end = Utilities.formatDate(new Date(end+24*60*60*1000), calendarTimezone, "YYYY-MM-dd hh:mm:ss a");     //End date time should be midnight of the selected date, so add 24 hours
      let calName = calendar.getName();
      let events = calendar.getEvents(start,end);  

      if(events.length == 0){//if no events are returned for the calendar, skip parsing, but keep track of calendar.
        noEventsList.push(calName);
       // console.log("No events were returned for the calendar [" + calName + "] over the date range specified. Please try again.");
       // return "No events were returned for the calendar [" + calName + "] over the date range specified. Please try again.";
      } else {
    /**Create Array Data**/
      if(type == "participant"){//Iterate at the participant level
        for(let i=0;i<events.length;i++){//for each event in calendar
          let eventStartUnformatted = events[i].getStartTime();
          let eventStart = Utilities.formatDate(eventStartUnformatted, calendarTimezone, "YYYY-MM-dd hh:mm:ss a");
          let eventEndUnformatted = events[i].getEndTime();
          let eventEnd = Utilities.formatDate(eventEndUnformatted, calendarTimezone, "YYYY-MM-dd hh:mm:ss a");        
          let hoursDuration = (eventEndUnformatted-eventStartUnformatted)/(1000*60*60);//time in hours
          let eventName = events[i].getTitle();
          let location = events[i].getLocation();
          let eventColor = convertColorCode(events[i].getColor());
          let description = events[i].getDescription();
          let numbParticipants = 0;
          // var isAllDayEvent = events[i].isAllDayEvent();
          // var isOwnedByMe = events[i].isOwnedByMe();
          let isRecurringEvent = events[i].isRecurringEvent();
          let dateCreated = events[i].getDateCreated();
          let lastUpdated =  events[i].getLastUpdated();        
          let visibility = events[i].getVisibility();
          // var tags = events[i].getAllTagKeys();
          // var status = events[i].getMyStatus();      
          let participants = events[i].getGuestList(true);
          let eventMembers = []; //keep track of team members already counted for this event   
          let owners = events[i].getCreators();
          let owner = owners;
          var guestEmail = "";
          if(eventName.includes(keyWordInput)){//only add participants if event title includes key word
            for(let j=0;j<participants.length;j++){//count the number of team members in the event
              guestEmail = participants[j].getEmail();        
              if(eventMembers.indexOf(guestEmail) == -1){//only count if they have not yet been counted
                let outputArray = [calName,eventName,eventStart,eventEnd,location,eventColor,owner,guestEmail,description,hoursDuration,dateCreated,lastUpdated,visibility,numbParticipants,isRecurringEvent,calendarTimezone];          
                numbParticipants = numbParticipants+1;
                var rowData = [];
                for(let j=0;j<selectedIndices.length;j++){
                  rowData.push(outputArray[selectedIndices[j]]);
                }  
                calData.push(rowData);
                eventMembers.push(guestEmail)
              }
            }
          }
        }
      } else if(type == "event"){//Iterate at the event level 
          for(let i=0;i<events.length;i++){
            let eventStartUnformatted = events[i].getStartTime();
            let eventStart = Utilities.formatDate(eventStartUnformatted, calendarTimezone, "YYYY-MM-dd hh:mm:ss a");
            let eventEndUnformatted = events[i].getEndTime();
            let eventEnd = Utilities.formatDate(eventEndUnformatted, calendarTimezone, "YYYY-MM-dd hh:mm:ss a");        
            let hoursDuration = (eventEndUnformatted-eventStartUnformatted)/(1000*60*60);//time in hours
            let eventName = events[i].getTitle();
            let location = events[i].getLocation();
            let eventColor = convertColorCode(events[i].getColor());       
            let description = events[i].getDescription();
            let numbParticipants = 0;  
          // var isAllDayEvent = events[i].isAllDayEvent();
          // var isOwnedByMe = events[i].isOwnedByMe();
            let isRecurringEvent = events[i].isRecurringEvent();
            let dateCreated = events[i].getDateCreated();
            let lastUpdated =  events[i].getLastUpdated();        
            let visibility = events[i].getVisibility();

          // var status = events[i].getMyStatus();      
            
            let participants = events[i].getGuestList(true);
            let eventMembers = []; //keep track of team members already counted for this event   
            let owners = events[i].getCreators();
            let owner = owners;
            var guestEmails = ""; //using plural emails here since we will be packing all emails into one event row
            if(eventName.includes(keyWordInput)){//only push event if event title includes key word
              for(let j=0;j<participants.length;j++){//count the number of team members in the event
                let guestEmail = participants[j].getEmail();        
                if(eventMembers.indexOf(guestEmail) == -1){//only count if they have not yet been counted
                guestEmails += ", " + guestEmail; //Add guest email to the emails string
                  numbParticipants = numbParticipants+1;
                  eventMembers.push(guestEmail)
                }
              }
              let outputArray = [calName,eventName,eventStart,eventEnd,location,eventColor,owner,guestEmails.substring(2),description,hoursDuration,dateCreated,lastUpdated,visibility,numbParticipants,isRecurringEvent,calendarTimezone];  
                  var rowData = [];
                  for(let j=0;j<selectedIndices.length;j++){
                    rowData.push(outputArray[selectedIndices[j]]);
                  }  
                  calData.push(rowData);           
              // console.log("EventTitle: " + eventName + " ; EventTeamMembers: " + eventMembers)
            }
          }
        }
      }
    }

    if(calData.length == 0){
      if(noEventsList.length == 0){ 
        var feedbackMessage = "Your Event Title Search returned no results for the specified period";
        } else {
        var feedbackMessage = "The following selected calendars had no events for the specified period:  " + noEventsList;
        }
      if(type == "event"){
        console.log("No Calendar Data was returned. " + feedbackMessage + ".  Please try again.");
        return "No Calendar Data was returned. " + feedbackMessage  + ".  Please try again.";
      } else if (type == "participant"){
        console.log("No Calendar Data was returned. There were no meeting participants invited to the events in selected calendars.  Please try again.");
        return "No Calendar Data was returned. There were no meeting participants invited to the events in selected calendars.  Please try again.";        
      }
    }  
    calData.sort(sortFunction);
    
  /**Create or Define Sheet**/
    if(!spreadsheet.getSheetByName("CalData")){
      var calDataSheet = spreadsheet.insertSheet().setName("CalData");
      //console.log("no sheet");
    } else {
      var calDataSheet = spreadsheet.getSheetByName("CalData");
      calDataSheet.getRange(1,1,calDataSheet.getLastRow()+1,20).clear();
      //console.log("sheet");
    }
    
  /**Add Headers to Sheets and Format **/
    var selectedHeaders = [[]];
    for (let i=0;i<selectedIndices.length;i++){
      selectedHeaders[0].push(calAttr[selectedIndices[i]][1]);
      calDataSheet.setColumnWidth(i+1,calAttr[selectedIndices[i]][2]);                             
    }
    calDataSheet.getRange(1,1,1,selectedIndices.length).setValues(selectedHeaders).setFontWeight('900').setBackground('#d9d9d9');

        
    /**Insert Data into Sheet**/
    calDataSheet.getRange(2,1,calData.length,selectedIndices.length).setNumberFormat("General");
    calDataSheet.getRange(2,1,calData.length,selectedIndices.length).setValues(calData);  
    calDataSheet.activate();
    const scriptEnd = new Date();
    const scriptTime = scriptEnd-scriptStart;
    console.log("Extract Completed! Number of Items: " + calData.length + "; Script Runtime: " + scriptTime + "; Time Zone: " + calendarTimezone + " (" + timezoneOffset + ")");
    return null;
  } catch(e){
    console.log("Unhandled Exception for getCalData: " + e);
  }
}

function sortFunction(a, b) {
  if (a[0] < b[0]) {
    return -1;
  }
  if (a[0] > b[0]) {
    return 1;
  }
  return 0;
}


function convertColorCode(code){
  try{
  switch (code) {
    case "1":
      return "Pale Blue"
      break;
    case "2":
      return "Pale Green"
      break;
    case "3":
      return "Mauve"
      break;
    case "4":
      return "Pale Red"
      break;
    case "5":
      return "Yellow"
      break;
    case "6":
      return "Orange"
      break;
    case "7":
      return "Cyan"
      break;
    case "8":
      return "Gray"
      break;
    case "9":
      return "Blue"
      break;
    case "10":
      return "Green"
      break; 
    case "11":
      return "Red"
      break;
    default:
      return "Calendar Default"
      break
  }
  } catch(e){
    console.log("Unhandled Exception for convertColorCode: " + e);
  }
}

function getTimeZoneOffset(date, timeZone) {
  //Cast the date to calendar timezone: e.g. 1/21/2023, 08:09:01
  const iso = date.toLocaleString('en-US', { timeZone: timeZone, hour12: false });
 // console.log(iso);
  //parse time string to get datetime elements
  const time = iso.slice(iso.indexOf(",")+2,iso.length).padStart(8,"0");
  const year = iso.slice(iso.indexOf(",")-4,iso.indexOf(","));
  const month = iso.slice(0,iso.indexOf("/")).padStart(2,"0");
  const secondHalf = iso.slice(iso.indexOf("/")+1,iso.indexOf(","));
  const day = secondHalf.slice(0,secondHalf.indexOf("/")).padStart(2,"0");

  //recreate string in the correct UTC Date format: e.g. 2023-01-21T08:10:52 and turn back into a date
  const utcDateString = year + "-" + month + "-" + day + "T" + time;
  //console.log(utcDateString);
  const utcDate = new Date(utcDateString);
  //console.log(utcDate);
  // Return the difference in miliseconds
  // Positive values are West of GMT
  return -(utcDate - date);
}


/**Save presets**/
function updatePresets(presets){ 
  try{
  PropertiesService.getUserProperties().setProperty("presets", presets);
  console.log("Presets Updated");
  return getPresets();
  } catch(e){
    console.error("Error updating presets.  Error Code: " + e)
    return "Error updating presets";
  }
}

function getPresets(){
  try{
  let presets = PropertiesService.getUserProperties().getProperty("presets");
  console.log("Presets Retrieved!");
  return presets;
  } catch(e){
    console.error("Presets not Retrieved. Error: " + e);
  }
}

function logError(message){
     console.error("Presets not Retrieved. Error: " + message);
}

