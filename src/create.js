//function to get target data
function get_data_(){
  // range to read (the first row/ the last column in the target table)
  const create_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(create_sheet_name);
  const top_row = 2;
  const last_col = 7;
  const last_row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(create_sheet_name).getLastRow();

  //Get target data
  const data = create_sheet.getRange(top_row, 1, last_row-1, last_col).getValues();
  return data;
}

//function to send an email to guest(s)
function send_email_guest_(schedules){
  for (let i=0;i<guest_emails.length;i++){
    let guest_name = guest_names[i];
    let template = HtmlService.createTemplateFromFile('email-to-guest');
    template.guest_name = guest_name;
    template.schedules = schedules;
    let body = template.evaluate().getContent();
    GmailApp.sendEmail(guest_emails[i],guest_email_subject,body,{htmlBody:body});
  }
}

//function to create schedule in Google Calendar and send an email
function create_schedule() {
  const data = get_data_();
  console.log(data);

  // column number for each item, starting with 0
  const start_date_col = 0;
  const start_time_col = 1;
  const end_date_col = 2;
  const end_time_col = 3;
  const title_col = 4;
  const location_col = 5;
  const description_col = 6;

  //define calendar and email addresses
  const calendar = CalendarApp.getDefaultCalendar();
  const guest_emails_string = guest_emails.join(', ');
  
  //create events in Google Calendar
  const schedules = []; //This variable will be used for the function to send email
    
  try {
      for (i=0; i<data.length; i++) {
      let start_date = data[i][start_date_col];

      schedules.push((start_date.getMonth() + 1) + "/" + start_date.getDate()); //store start date for send an email to guests
      
      let start_time = data[i][start_time_col];
      let end_date = data[i][end_date_col];
      let end_time = data[i][end_time_col];
      let title = data[i][title_col];
      let location = data[i][location_col];
      let description = data[i][description_col];
      
      // set options for creating calendar
      let options = { 
                      location: location,
                      description: description,
                      guests: guest_emails_string
                    };

          // check if the start time and end time are empty
          if (start_time == '' || end_time == '') {
            // if start_date is equal to end_date, create an event for one day
            if (start_date.toString() == end_date.toString()){
              calendar.createAllDayEvent(
                title,
                start_date,
                options
              );
            console.log('This event is one day event with no time specified');
              // otherwise, create an event for multiple days
            } else {
              //since end_date is exclusive for createAllDayEvent, although the event ends on the day of the end_date, it does not include this day, and 1 more day should be added.
              end_date.setDate(end_date.getDate() + 1);
              calendar.createAllDayEvent(
                title,
                start_date,
                end_date,
                options
              );
              console.log('This event is multiple days event with no time specified');
            }
            
          // set other event with more detailed info
          } else {
            start_date.setHours(start_time.getHours(),start_time.getMinutes());
            end_date.setHours(end_time.getHours(),end_time.getMinutes());
            calendar.createEvent(
              title,
              start_date,
              end_date,
              options
            );
            console.log('This event is an event with start and end time specified');          }
      }
  //execute the function to send an email to recipients
    send_email_guest_(schedules);
    Browser.msgBox("Successfully made the schedules in Google Calendar and sent emails to recipients!");
    
  // output log if there is an error
  } catch (e) {
    // Browser.msgBox(`Error making schedule and sending email: ${e.message}`,Browser.Buttons.OK_CANCEL); 
    console.log(e.message);   
  }
  
}
