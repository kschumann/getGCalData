function  getUserCalendars(){
  try{
    const calendars = CalendarApp.getAllCalendars();
    const numbCalendars = calendars.length;
    var calendarList = [];
    for(var i=0;i<numbCalendars;i++){
      calendarList.push([calendars[i].getId(),calendars[i].getName()])
    }  
    return calendarList;
  } catch(e){
    console.log("Unhandled Exception for getUserCalendars: " + e);
  }
}

function getCalData(calIds,start,end,selectedColumns,type){
  try{
    const spreadsheet = SpreadsheetApp.getActive();
    const scriptStart = new Date();
    console.log("selectedColumns:" + selectedColumns);
    var calAttr = [
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
      ["isRecurringEvent","Is Recurring Event",100]
    ];
    
    /**Create single array of column IDs**/
    var calAttrList = [];
    for(var i=0;i<calAttr.length;i++){
      calAttrList.push(calAttr[i][0]);
    }

  /**Create an array of indices that should be used */
    var selectedIndices = [];
    for(var i=0;i<selectedColumns.length;i++){
      selectedIndices.push(calAttrList.indexOf(selectedColumns[i]));
    }  
  console.log(selectedIndices);  

  /**Create or Define Sheet**/
    if(!spreadsheet.getSheetByName("CalData")){
      var calDataSheet = spreadsheet.insertSheet().setName("CalData");
      console.log("no sheet");
    } else {
      var calDataSheet = spreadsheet.getSheetByName("CalData");
    calDataSheet.getRange(1,1,calDataSheet.getLastRow()+1,20).clear();
      console.log("sheet");
    }
    
  /**Add Headers to Sheets and Format **/
    var selectedHeaders = [[]];
    for (var i=0;i<selectedIndices.length;i++){
    selectedHeaders[0].push(calAttr[selectedIndices[i]][1]);
    calDataSheet.setColumnWidth(i+1,calAttr[selectedIndices[i]][2]);                             
    }
    calDataSheet.getRange(1,1,1,selectedIndices.length).setValues(selectedHeaders).setFontWeight('900').setBackground('#d9d9d9');
    
    var calData = [];//eventName,staffEmail, eventDate,eventDuration  

  /**Populate calData array **/
    for(var m=0;m<calIds.length;m++){//iterate through calendars
      var calendar = CalendarApp.getCalendarById(calIds[m]);
      var start = new Date(start);
      var end = new Date(end+24*60*60*1000);
      var calName = calendar.getName();
      var events = calendar.getEvents(start,end);  
      if(events.length == 0){
        return "No events were returned for the calendar [" + calName + "] over the date range specified. Please try again.";
      }   

  /**Create Array Data**/
    if(type == "participant"){//Iterate at the participant level
      for(var i=0;i<events.length;i++){//for each event in calendar
        var eventStart = events[i].getStartTime();
        var eventEnd = events[i].getEndTime();
        var hoursDuration = (eventEnd-eventStart)/(1000*60*60);//time in hours
        var eventName = events[i].getTitle();
        var location = events[i].getLocation();
        var eventColor = convertColorCode(events[i].getColor());
        var description = events[i].getDescription();
        var numbParticipants = 0;
        // var isAllDayEvent = events[i].isAllDayEvent();
        // var isOwnedByMe = events[i].isOwnedByMe();
        var isRecurringEvent = events[i].isRecurringEvent();
        var dateCreated = events[i].getDateCreated();
        var lastUpdated =  events[i].getLastUpdated();        
        var visibility = events[i].getVisibility();
        // var tags = events[i].getAllTagKeys();
        // var status = events[i].getMyStatus();      
        var participants = events[i].getGuestList(true);
        var eventMembers = []; //keep track of team members already counted for this event   
        var owners = events[i].getCreators();
        var owner = owners;
        var guestEmail = "";
        

        for(var j=0;j<participants.length;j++){//count the number of team members in the event
          var guestEmail = participants[j].getEmail();        
          if(eventMembers.indexOf(guestEmail) == -1){//only count if they have not yet been counted
            var outputArray = [calName,eventName,eventStart,eventEnd,location,eventColor,owner,guestEmail,description,hoursDuration,dateCreated,lastUpdated,visibility,numbParticipants,isRecurringEvent];          
            numbParticipants = numbParticipants+1;
            var rowData = [];
            for(var j=0;j<selectedIndices.length;j++){
              rowData.push(outputArray[selectedIndices[j]]);
            }  
            calData.push(rowData);
            eventMembers.push(guestEmail)
          }
        }
      }
    } else if(type == "event"){//Iterate at the event level 
        for(var i=0;i<events.length;i++){
          var eventStart = events[i].getStartTime();
          var eventEnd = events[i].getEndTime();
          var hoursDuration = (eventEnd-eventStart)/(1000*60*60);//time in hours
          var eventName = events[i].getTitle();
          var location = events[i].getLocation();
          var eventColor = convertColorCode(events[i].getColor());       
          var description = events[i].getDescription();
          var numbParticipants = 0;
          
        // var isAllDayEvent = events[i].isAllDayEvent();
        // var isOwnedByMe = events[i].isOwnedByMe();
          var isRecurringEvent = events[i].isRecurringEvent();
          var dateCreated = events[i].getDateCreated();
          var lastUpdated =  events[i].getLastUpdated();        
          var visibility = events[i].getVisibility();

        // var status = events[i].getMyStatus();      
          
          var participants = events[i].getGuestList(true);
          var eventMembers = []; //keep track of team members already counted for this event   
          var owners = events[i].getCreators();
          var owner = owners;
          var guestEmails = ""; //using plural emails here since we will be packing all emails into one event row
          
          for(var j=0;j<participants.length;j++){//count the number of team members in the event
            var guestEmail = participants[j].getEmail();        
            if(eventMembers.indexOf(guestEmail) == -1){//only count if they have not yet been counted
            guestEmails += ", " + guestEmail; //Add guest email to the emails string
              numbParticipants = numbParticipants+1;
              eventMembers.push(guestEmail)
            }
          }
          var outputArray = [calName,eventName,eventStart,eventEnd,location,eventColor,owner,guestEmails.substring(2),description,hoursDuration,dateCreated,lastUpdated,visibility,numbParticipants,isRecurringEvent];  
              var rowData = [];
              for(var j=0;j<selectedIndices.length;j++){
                rowData.push(outputArray[selectedIndices[j]]);
              }  
              calData.push(rowData);           
          // console.log("EventTitle: " + eventName + " ; EventTeamMembers: " + eventMembers)
        }
      }
    }

    /**Insert Data into Sheet**/
    calDataSheet.getRange(2,1,calData.length,selectedIndices.length).setNumberFormat("General");
    calDataSheet.getRange(2,1,calData.length,selectedIndices.length).setValues(calData);  
    calDataSheet.activate();
    var scriptEnd = new Date();
    var scriptTime = scriptEnd-scriptStart;
    console.log("Number of Items: " + calData.length);
    console.log("Script Runtime: " + scriptTime);
    return null;
  } catch(e){
    console.log("Unhandled Exception for getCalData: " + e);
  }
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