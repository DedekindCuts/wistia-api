// Please consult the README for more information about this script

function GetWistiaViews() {
  
  var ui = SpreadsheetApp.getUi();
  
  //url: base Wistia API url to get stats about video views
  var url = "https://api.wistia.com/v1/stats/events.json";
  
  //The variables defined here may be adjusted depending on your preferences
  // password: API access key
  var password = "YOUR_WISTIA_API_KEY_HERE";
  // per_page: how many results to return per page (100 maximum)
  var per_page = "100";
  // page: which page of results to start with (starts at 1, not 0)
  var page = 1;
  // sort_by: sort results by "name", "created", or "updated" (**as far as I can tell, Wistia's Stats API doesn't support sorting of results, these parameters are from the Data API**)
  var sort_by = "created";
  // sort_direction: "1" for ascending, "0" for descending (**as far as I can tell, Wistia's Stats API doesn't support sorting of results, these parameters are from the Data API**)
  var sort_direction = "0";
  
  // this function prompts the user to choose the dates for which they would like to receive data
  var dates = GetDates();
  if(dates == "error") {
    Logger.log("Error getting dates, or else user clicked close button to terminate script.");
    return;
  }
  
  // GetEvents uses the dates returned from GetDates to retrieve the appropriate data from the Wistia API
  // WriteToSheet formats the data, choosing the desired fields, and writes it to the spreadsheet
  // This loop increments the page number, since the Wistia API returns a maximum of 100 results; WriteToSheet returns false when the desired page is empty, exiting the loop
  do {
    var response = GetEvents(url, password, per_page, page, sort_by, sort_direction, dates["start"], dates["end"]);
    if(response == "error") {
      ui.alert("Error, check logs", ui.ButtonSet.OK);
      return;
    }
    try {
      var next_page = WriteToSheet(response);
      page++;
    } catch(error) {
      ui.alert("Error: " + error.message, ui.ButtonSet.OK);
      Logger.log("Error writing events page " + page + ": " + error.message);
      return;
    }
  } while(next_page == true);
  
  ui.alert("Finished! Recorded Wistia video views from " + dates["start"] + " to " + dates["end"] + ".", ui.ButtonSet.OK);
  return;
}

function GetEvents(url, password, per_page, page, sort_by, sort_direction, start_date, end_date) {
  
  // Calls the Wistia API to get video views for a specified date
  
  var ui = SpreadsheetApp.getUi();
  
  var request = url 
              + "?api_password=" + password 
              + "&per_page=" + per_page 
              + "&page=" + page 
              + "&sort_by=" + sort_by 
              + "&sort_direction=" + sort_direction 
              + "&start_date=" + start_date 
              + "&end_date=" + end_date;
  
  try {
    var response = UrlFetchApp.fetch(request);
  } catch(error) {
    ui.alert("Error: " + error.message, ui.ButtonSet.OK_CANCEL);
    Logger.log("Error with API request on page " + page + ": " + error.message);
    return "error";
  }
  
  return response;
  
}

function GetDates() {
  
  // Prompts the user to get the desired dates for which to return data
  
  var ui = SpreadsheetApp.getUi();
  
  var today_prompt = ui.alert('Do you want to only get data for yesterday\'s views?', ui.ButtonSet.YES_NO);
  if (today_prompt == ui.Button.YES) {
    var today = new Date();
    var yesterday = new Date(today.getTime() - 1*24*60*60*1000);
    var start_date = Utilities.formatDate(yesterday, "GMT-7", "yyyy-MM-dd");
    var end_date = start_date;
    return {"start": start_date, "end": end_date}
  } else if (today_prompt == ui.Button.NO) {
    var start_response = ui.prompt('Please enter the FIRST date for which you would like to get data, in the format \"YYYY-MM-DD\"', ui.ButtonSet.OK_CANCEL);
    if (start_response.getSelectedButton() == ui.Button.OK) {
      var start_date = start_response.getResponseText();
      Logger.log('The user inputted the start date: %s.', start_response.getResponseText());
    } else if (start_response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('Start date not provided');
      return "error";
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
      return "error";
    }
    var end_response = ui.prompt('Please enter the LAST date for which you would like to get data, in the format \"YYYY-MM-DD\"', ui.ButtonSet.OK_CANCEL);
    if (end_response.getSelectedButton() == ui.Button.OK) {
      var end_date = end_response.getResponseText();
      Logger.log('The user inputted the end date: %s.', end_response.getResponseText());
    } else if (end_response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('End date not provided');
      return "error";
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
      return "error";
    }
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
    return "error";
  }
  
  return {"start": start_date, "end": end_date}
}

function WriteToSheet(response) {
  
  // Appends the relevant information returned from GetEvents to the bottom of the sheet
  
  var json = response.getContentText();
  var data = JSON.parse(json);
  var sheet = SpreadsheetApp.getActiveSheet();
  
  if(data == '') {
    return false;
  }
  else {
    for(i=0; i<data.length; i++){
      // Remove this IF statement (while keeping the code inside of it) if you want unidentified events to also be written to the sheet
      if(Object.getOwnPropertyNames(data[i]["conversion_data"]).length != 0) {
        var row = [data[i]["received_at"].substring(0,10), 
                   data[i]["media_name"], 
                   data[i]["conversion_data"]["email"], 
                   data[i]["conversion_data"]["first_name"] + " " + data[i]["conversion_data"]["last_name"], 
                   data[i]["percent_viewed"]];
        // The arguments for getRange are (row, column, number of rows, number of columns)
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
      }
    }
    return true;
  }
}

function onOpen() {
  
  //adds a menu button to the spreadsheet to run the function GetWistiaViews; only needs to be run once to add the button to the spreadsheet
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Wistia')
      .addItem('Get Wistia Views','GetWistiaViews')
      .addToUi();
}