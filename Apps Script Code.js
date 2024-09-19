//Created by Michael Kang

//your Google Sheets file -> Extensions -> Apps Script
////paste the code below

//Button to be always open
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Create Events')
      .addItem('Create Events', 'event_create')
      .addSeparator()
      .addToUi();
}



function event_create() {

  // your email here!
  var calendar_id = ""; // your email here!
  // your email here!

  var sheet_tasks = SpreadsheetApp.getActive().getSheetByName('Tasks');
  var tasks = sheet_tasks.getRange("A2:K100").getValues(); // adjust range as needed

  for (var i = 0; i < tasks.length; i++) {
    var i_row = i + 2; // row identifier as in Sheets
    
    var task = tasks[i];

    var task_main = task[0]; // Assuming task is in A
    var task_sub = task[1]; // Assuming subtask is in B
    var task_location = task[2]; // Assuming location is in C
    var task_desc = task[3]; // Assuming description is in D
    var added = task[10]; // Assuming indicator whether added is in K
      Logger.log(added);
    var task_time_start = task[4]; // start time
    var task_time_end = task[5];   // end time
    var task_date_start = task[6]; // start date
    var task_date_end = task[7];   // end date
  
    // Check if the event has already been added
    if (added === "Added to Calendar") {
      Logger.log("Skipping: Event already added.");
    } else{

      var event_name = task_main + "-" + task_sub;
      var calendar = CalendarApp.getCalendarById(calendar_id);

      // Convert task_date_start and task_date_end to Date objects
      task_date_start = new Date(task_date_start);
      task_date_end = new Date(task_date_end);

      // Ensure task_name is not null or empty
      if (!task_main || !task_sub) {
        Logger.log("Event name missing.");
      } else{
        if (task_time_start && task_time_end) {
          // Convert task_time_start and task_time_end to HH:mm format
          var formatted_task_time_start = task_time_start.toString().substring(0, 2) + ":" + task_time_start.toString().substring(2, 4);
          var formatted_task_time_end = task_time_end.toString().substring(0, 2) + ":" + task_time_end.toString().substring(2, 4);

          // Combine date and time
          var task_start = new Date(task_date_start.toDateString() + " " + formatted_task_time_start);
          var task_end = new Date(task_date_end.toDateString() + " " + formatted_task_time_end);

          Logger.log("Task with time: " + task_start + " - " + task_end);

          // Create event with time
          calendar.createEvent(event_name, task_start, task_end, {
            description: task_desc,
            location: task_location
          });

        } else {
          Logger.log("No time specified. Creating an all-day event.");

          // If the start and end dates are the same, create a single-day event
          if (task_date_start.getTime() === task_date_end.getTime()) {
            Logger.log("Creating single-day all-day event.");
            calendar.createAllDayEvent(event_name, task_date_start, {
              description: task_desc,
              location: task_location
            });
          } else {
            Logger.log("Creating multi-day all-day event.");
            task_date_end.setDate(task_date_end.getDate() + 1); // Extend end date for multi-day all-day event
            calendar.createAllDayEvent(event_name, task_date_start, task_date_end, {
              description: task_desc,
              location: task_location
            });
          }
        }

        // Mark the task as added in the spreadsheet
        sheet_tasks.getRange(i_row, 11).setValue("Added to Calendar");

      }
    }
  }
}
