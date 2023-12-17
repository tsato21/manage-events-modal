const update_delete_sheet_name = 'Update-Delete';
const update_delete_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(update_delete_sheet_name);

//function to show modal to input start and end dates
function show_period_input_modal(){
  const template = HtmlService.createTemplateFromFile('period-input').evaluate()
                  .setWidth(400) // set the width
                  .setHeight(300); // set the height
  SpreadsheetApp.getUi().showModalDialog(template,'Search Events during Designated Period');
}

//function to show modal to input keyword and target period
function show_keyword_input_modal(){
  const template = HtmlService.createTemplateFromFile('keyword-input').evaluate()
                    .setWidth(400)
                    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(template,'Search Events with Keyword')
}

function show_search_type_input_modal(){
  const template = HtmlService.createTemplateFromFile('search-type-input').evaluate()
          .setWidth(400)
          .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(template,'Enter a searh type that you want to use');
}

function choice_search_type(search_type){
  if(search_type == 'period'){
    show_period_input_modal();
  } else if (search_type == 'keyword'){
    show_keyword_input_modal();
  }
}

//function to display events that are in the period input by the user in Spreadsheet
function display_events_by_period(start,end) {
  try {
    const start_date = new Date(start);
    const end_date = new Date(end);
    start_date.setHours(0,0);
    end_date.setHours(23,59);
    console.log(start_date,end_date);
    
    SpreadsheetApp.getActiveSpreadsheet().getRange("A2:I").clear();

    //get events during target period and reflect them in the spreadsheet
    const events = CalendarApp.getEvents(start_date,end_date);
    const event_length = events.length;
    if(event_length == 0){
      Browser.msgBox("No event is found during the designated period.");
    } else {
      const range = update_delete_sheet.getRange(2,1,event_length,8);
      const contents = [];
      for (i=0; i < event_length; i++) {
        let start = events[i].getStartTime();
        let start_date = start.getMonth()+1 + '/' + start.getDate();
        let start_time;
        if(start.getMinutes() === 0 ){
          start_time = start.getHours() + ':' + '00';
        } else {
          start_time = start.getHours() + ':' + start.getMinutes();
        }        
        let end = events[i].getEndTime();
        let end_date = end.getMonth()+1 + '/' + end.getDate();
        let end_time;
        if(end.getMinutes() === 0){
          end_time = end.getHours() + ':' + '00';
        } else {
          end_time = end.getHours() + ':' + end.getMinutes();
        }
        let title = events[i].getTitle();
        let location = events[i].getLocation();
        let description = events[i].getDescription();
        let id = events[i].getId();
        contents.push([start_date,start_time,end_date,end_time,title,location,description,id]);
      }
      range.setValues(contents);
      set_border_validation(event_length);
      // console.log(contents);

      Browser.msgBox("Events during the designated period are displayed on the sheet.");
    }
  } catch (e) {
    console.log('Error in display_events: ' + e.message);
  }
}

//function to display events that match a keyword input by the user in Spreadsheet
function display_events_by_keyword(period, keyword) {
  // convert period to milliseconds (1 day = 1*24*60*60*1000 milliseconds)
  const period_ms = period * 24 * 60 * 60 * 1000;
  const now = new Date();
  const start_time = new Date(now.getTime() - period_ms); // Subtract period from current date
  // console.log(start_time,period_ms,now);
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(start_time, now);
  const events_by_keyword = events.filter(function(event) {
    return event.getTitle().toLowerCase().includes(keyword.toLowerCase());
  });
  const event_length = events_by_keyword.length;

  SpreadsheetApp.getActiveSpreadsheet().getRange("A2:I").clear();
  try{
    if(event_length == 0){
      Browser.msgBox("No event is matched to the keyword.");
    } else {
      const range = update_delete_sheet.getRange(2,1,event_length,8);
      const contents = [];
      for(i = 0; i < events_by_keyword.length; i++) {
        let start = events[i].getStartTime();
        let start_date = start.getMonth()+1 + '/' + start.getDate();
        let start_time;
        if(start.getMinutes() === 0 ){
          start_time = start.getHours() + ':' + '00';
        } else {
          start_time = start.getHours() + ':' + start.getMinutes();
        }        
        let end = events[i].getEndTime();
        let end_date = end.getMonth()+1 + '/' + end.getDate();
        let end_time;
        if(end.getMinutes() === 0){
          end_time = end.getHours() + ':' + '00';
        } else {
          end_time = end.getHours() + ':' + end.getMinutes();
        }
        let title = events[i].getTitle();
        let location = events[i].getLocation();
        let description = events[i].getDescription();
        let id = events[i].getId();
        contents.push([start_date,start_time,end_date,end_time,title,location,description,id]);
      }
      // console.log(contents);
      range.setValues(contents);
      set_border_validation(event_length);

      Browser.msgBox('Events that matched the keyword are displayed on the sheet.');
    }
  } catch(e){
    console.log('Error in display_events_by_keyword: ' + e.message);
  }
}


function update_delete_events() {
  const last_row = update_delete_sheet.getLastRow();
  const data = update_delete_sheet.getRange(2, 1, last_row - 1, 9).getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  try {
    for (const each_data of data) {
      if (each_data[8] === 'Update') {
        let event = getEventByIdOrTitle_(calendar, each_data[7], each_data[4], new Date(each_data[0]));
        if (event) {
          if (event.isAllDayEvent()) {
            // For all-day events
            let newStartDate = new Date(each_data[0]);
            let newEndDate = new Date(each_data[2]);
            event.setAllDayDates(newStartDate, newEndDate);
          } else {
            // For events with specific times
            let start_date = new Date(each_data[0]);
            let s_hours = each_data[1].getHours();
            let s_minutes = each_data[1].getMinutes();
            start_date.setHours(s_hours, s_minutes);
            let end_date = new Date(each_data[2]);
            let e_hours = each_data[3].getHours();
            let e_minutes = each_data[3].getMinutes();
            end_date.setHours(e_hours, e_minutes);
            event.setTime(start_date, end_date);
          }
          event.setTitle(each_data[4]);
          event.setLocation(each_data[5]);
          event.setDescription(each_data[6]);
        }
      } else if (each_data[8] === 'Delete') {
        console.log(`${each_data[4]} will be deleted.`);
        deleteEventByIdOrTitle_(calendar, each_data[7], each_data[4], new Date(each_data[0]));
      }
    }
    Browser.msgBox("Successfully updated or deleted designated events from Google Calendar.");
  } catch (e) {
    console.log(`Error updating/deleting: ${e.message}`);
    Browser.msgBox(`Error updating/deleting: ${e.message}`, Browser.Buttons.OK_CANCEL);
  }
}

function getEventByIdOrTitle_(calendar, eventId, eventTitle, eventDate) {
  let event = calendar.getEventById(eventId);
  if (!event) {
    // If event is not found by ID, try finding by title and date
    let events = calendar.getEventsForDay(eventDate);
    for (let i = 0; i < events.length; i++) {
      if (events[i].getTitle() === eventTitle && (events[i].isAllDayEvent() || events[i].getId() === eventId)) {
        return events[i];
      }
    }
  }
  return event;
}

function deleteEventByIdOrTitle_(calendar, eventId, eventTitle, eventDate) {
  try {
    let event = calendar.getEventById(eventId);
    if (!event) {
      // If event is not found by ID, try finding by title and date
      let events = calendar.getEventsForDay(eventDate);
      for (let i = 0; i < events.length; i++) {
        if (events[i].getTitle() === eventTitle && events[i].isAllDayEvent()) {
          event = events[i];
          break;
        }
      }
    }
    if (event) {
      event.deleteEvent();
    }
  } catch (e) {
    console.log(`Error deleting event: ${e.message}`);
  }
}


function test(){
  console.log(CalendarApp.getDefaultCalendar().getEventById("7kukuqrfedlm2f9t5gq1l9s81tu5i6qmcs4587eks8qeslarr9pesedrudmr9ivlhu20").getTitle());
}

