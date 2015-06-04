/*  This is google spreadsheet for topcoder attempts aggregator
*/
/***
**** Future Implementation - Dashboard App, Cleaner Score Entry, Form UI for entry, Better Analytics
***/

var today = []
/**
 ** Reset the todays variable
 **/
function resetToday() {
  today = [];
}

/**
 * Get the count of problems attempted
 * Return - Array: [User Data, Questions]
 */
function getCount(logging) {
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var complete_data = {};
  Logger.log(sheets);
  for(var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s]
    if(sheet.getName() === "Result" ||
      sheet.getName() === "garbage_db_sheet" ) continue;
    var rows = sheet.getDataRange();
    var numRows = rows.getNumRows();
    var values = rows.getValues();
    var users = values[0];
    Logger.log(users);
    var attempts = [];
    var sheet_data = {};
    var min = 999;
    var avg = 0;
    Logger.log(sheet.getActiveRange().getColumn());
    logging && Logger.log(sheet.getActiveRange().getColumn());
    
    /* Get the users */
    for(var k = 1; k < users.length; k++) {
      complete_data[users[k]] = {questions: [], attempts: 0};
      sheet_data[users[k]] = {questions: [], attempts: 0};
    }

    if (logging) Logger.log(users);
    var questions = {};
    /* Get their tc scores */
    for (var i = 3; i <= numRows - 1; i++) {
      var row = values[i];
      questions[row[0]] = 0;
      //Logger.log(row);
      for ( var j = 1; j < users.length; j++) {
        if( row[j] != "" && typeof row[j] == 'number') {
          complete_data[users[j]]['questions'].push({'q':row[0],'s':row[j]});
          sheet_data[users[j]]['questions'].push({'q':row[0],'s':row[j]});
          complete_data[users[j]]['attempts'] += 1;
          sheet_data[users[j]]['attempts'] += 1;
          questions[row[0]] += 1;
        }
      }
    }
   
    // Write data into the spreadsheet on attempts
    for(var user in sheet_data) {
      attempts.push(sheet_data[user]['attempts']);
      if(min > sheet_data[user]['attempts'])
        min = sheet_data[user]['attempts'];
      avg += sheet_data[user]['attempts'];
    }
    avg = Math.round(avg/(attempts.length));
    if (logging) Logger.log("Average: "+avg);
    
    /* Color the cells - yellow for below average, red for bottom */
    var pos = 1;
    for(user in sheet_data) {
      if (logging) Logger.log(user);
      pos += 1;
      if(sheet_data[user]['attempts'] == min) {
        //color red
        if (logging) Logger.log("red here");
        sheet.getRange(1, pos, 1, 1).setBackground("red");
      }
      else if(sheet_data[user]['attempts'] < avg) {
        //color yellow
        if (logging) Logger.log("yellow here");
        sheet.getRange(1, pos, 1, 1).setBackground("yellow");
      }
      else {
        //color white
        if (logging) Logger.log("white here");
        sheet.getRange(1, pos, 1, 1).setBackground("white");
      }
      
    }
    
    /* Write attempts aggregate to the sheet */
    if (logging) Logger.log(sheet_data);
    var dataRange = sheet.getRange(2, 2, 1, users.length-1);
    var attempts_1 = [];
    attempts_1.push(attempts);
    dataRange.setValues(attempts_1);
    
    /* Complete the chart data, in specified time only*/
    var date = new Date();
    if (logging) Logger.log(typeof date.getHours());
    if(date.getHours() > 0 && date.getHours() < 3) {
      dataRange = sheet.getRange(3,2,1, users.length - 1);
      var prev_data = values[2];
      var to_write_data = [];
      if (logging) Logger.log(prev_data);
      var i = 1;
      var todayDate = date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getYear();
      for(user in sheet_data) {
        //      Logger.log(prev_data[i]);
        str = todayDate + ":" + sheet_data[user]['attempts'] + ";"
        if(prev_data[i] != str)
          str = prev_data[i]+str;
        else
          str = prev_data[i];
        to_write_data.push(str);
        i++;
      }
      if (logging) Logger.log(":"+prev_data);
      if (logging) Logger.log(prev_data.length);
      //prev_data.pop();
      dataRange.setValues([to_write_data]);
    }
  }
  return [complete_data,questions];
}

function testGetCount() { try { Logger.log("hello!?"); getCount(true); } catch(e) { Logger.log(e.message); } }

// ** Get 
function myOnOpen() {
  
  var complete_data = getCount()[0];
  var questions = getCount()[1];
  var app = UiApp.createApplication();
  app.setTitle("Summary");
  var vapp = app.createVerticalPanel();
  var message = "<b>Questions not attempted: </b>";
  var q_attempted = "<br /> <b> Questions attempted: </b>";;
  for (q in questions) {
    if(questions[q] == 0)
      message += q + ";";
    else
      q_attempted += q + ":" + questions[q] + ";  ";
  }
  message += q_attempted;
  message += "<br /><ul>";
  for (data in complete_data) {
    message += "<li> <b>" + data + "</b> <ul><li> Attempts: " + complete_data[data]['attempts'] +"</li><li> Questions: ";
    for(var i = 0; i < (complete_data[data]['questions']).length; i++) {
      message += complete_data[data]['questions'][i]['q'] + "(" + complete_data[data]['questions'][i]['s'] + ")";
      if ( i != ((complete_data[data]['questions']).length - 1)) {
        message += ", ";
      }
    }
    
    message += "</li></ul></li>";
  }
  message += "</ul>";
  vapp.add(app.createHTML(message));
  Logger.log(message);
  
  var scroll = app.createScrollPanel().setPixelSize(500, 400);
  scroll.add(vapp);
  app.add(scroll);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
  
};

function mySendEmail() {
  try {
    var complete_data = getCount()[0];
    
    var message = "";
    for (data in complete_data) {
      message += data + ": Attempts: " + complete_data[data]['attempts'] +"; Questions: ";
      for(var i = 0; i < (complete_data[data]['questions']).length; i++) {
        message += complete_data[data]['questions'][i]['q'] + "(" + complete_data[data]['questions'][i]['s'] + ")";
        if ( i != ((complete_data[data]['questions']).length - 1)) {
          message += ", ";
        }
      }
      message += "\n";
    }
    
    Logger.log(message);
    
    MailApp.sendEmail("adicoolrao@gmail.com", "Top Coder Attempts", message);
    
    
    
    
  } catch(e) {
    MailApp.sendEmail("adicoolrao@gmail.com", "Top Coder Attempts - Error!", e.message);
  }
}

/* Draw charts based on data */
function chartsApp() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var data = [];
  var dataTable = Charts.newDataTable();
  dataTable.addColumn(Charts.ColumnType.STRING, "Dates");
  
  /* Get data */
  for(var s = 0; s < sheets.length-1; s++) {
    var sheet = sheets[s]
    var rows = sheet.getDataRange();
    var numRows = rows.getNumRows();
    var values = rows.getValues();
    var users = values[0];
    var sheet_data = values[2];
    Logger.log(sheet_data);
    var flag = 0;
    for(var i = 1; i < users.length; i++) {
      dataTable.addColumn(Charts.ColumnType.NUMBER, users[i]);
      data.push(String(sheet_data[i]).split(";"));
      //Logger.log(sheet_data[i].split(";")[2]);
    }
  }
  //Logger.log(data);
  // Get dates
  var to_insert = {};
  for(var i = 0; i < data.length; i++) {
    for(var j = 0; j < data[i].length; j++) {
      var d = String(data[i][j]).split(":");
      if(d.length > 1) {
        if(!(d[0] in to_insert)) {
          to_insert[d[0]] = [];
        }
        to_insert[d[0]].push(d[1]);       
      }
    }
  }
  Logger.log(to_insert);
  for(var date in to_insert) {
    //Logger.log(to_insert[date]);
    dataTable.addRow([date].concat(to_insert[date]));
  }
  /* Draw Chart */
  var chart = Charts.newLineChart()
      .setDataTable(dataTable)
      .setTitle("Number of attempts in topcoder")
      .setDimensions(900, 650)
      .build();
  
  var app = UiApp.createApplication().setTitle("Topcoder Chart");
  app.setHeight(650);
  app.setWidth(900);
  app.add(chart);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

/** 
 * Rewrite chart data and remove duplicates.
 */
function rewriteData() {

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  for(var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var rows = sheet.getDataRange();
    var numRows = rows.getNumRows();
    var values = rows.getValues()[2];
    var toWrite = [];
    
    
    for(var i = 1; i < values.length; i++) {
      var data = String(values[i]).split(";");
      var a = {};
      var b = "";
      for(var j = 0; j < data.length-1; j++) {
        
        if(!(data[j] in a)) {
          a[data[j]] = 1;
          b = b + String(data[j]) + ";";
        }
      }
      toWrite.push(b);
    }
    dataRange = sheet.getRange(3,2,1, toWrite.length);
    Logger.log(toWrite);
    dataRange.setValues([toWrite]);
  }

}

function mergeMine(){
  
  var mySheet = SpreadsheetApp.
  openByUrl("https://docs.google.com/spreadsheet/ccc?key=0AhvKbSbi6uNmdDNqYTNfcTlpNkRqaWRXRmw0SHBXa0E#gid=0").getSheets()[0];
  var sheetData = mySheet.getDataRange();
  var numRows = sheetData.getNumRows();
  var data = {};
  var values = sheetData.getValues();
  //Logger.log(values);
  for(var i = 1; i < numRows; i++) {
    //Logger.log(values[i][1] + "::" + values[i][2]);
    data[values[i][1]] = values[i][2];
  }
  var tcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2];
  sheetData = tcSheet.getDataRange();
  numRows = sheetData.getNumRows();
  values = sheetData.getValues();
  var toInsert = "";
  var notAttempted = "";
  for(var i = 1; i < numRows; i++) {
    if(values[i][0] in data) {
      Logger.log(data[values[i][0]]);
      toInsert = tcSheet.getRange(i+1, 9, 1, 1);
      toInsert.setValues([[String(data[values[i][0]])]]);
    }
    else {
      notAttempted += values[i][0] + ",";
    }
  }
  Logger.log(notAttempted);
}

/** 
 * Create summary in the 'Results' sheet. 
 * Includes -
 *           Updating the participants list, in order of #questions attempted.
 *           Rise and Fall in Rankings
 *           Biggest rise
 * Created: July 7, 2013
**/
function calculateRanks(triggered) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet = sheets[4];
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  Logger.log(triggered);

  // Get old_data
  var old_data = [];
  for(var i = 2; i < values.length; ++i) {
    var row = [];
    row.push(values[i][1]);
    row.push(values[i][2]);
    old_data.push(row);
  }
  Logger.log(old_data);
  
  // Get new_data from the score entry sheets 1-3
  var gc_data = getCount(false)[0];
  var new_data = [];
  for( var user in gc_data ) {
    var row = [];
    row.push(user);
    row.push(gc_data[user]['attempts']);
    new_data.push(row);
  }
  
  // Sort new_data
  new_data.sort(function(a,b){
    return b[1] - a[1];
  });

  Logger.log(new_data);
  
  // For data object to be written to the sheet.
  // Also if somebody has higher attempts than previous day, that changes will be shown.
  // Coloring applies to those who rises or drop in rankings.
  var   row_from = 3        // Row from which the data should be entered to the sheet
      , change_column = 5   // Column to display change on that day
      , db_column = 4       // To store changes
      , charting = [];
  for(var i = 0; i < new_data.length; ++i) {
    for(var j = 0; j < old_data.length; ++j) {
      if( new_data[i][0] === old_data[j][0] ) {
        if(triggered) {
          if ( i < j ) {
            sheet.getRange(i+row_from, 1, 1, 1).setBackground("green");
            sheet.getRange(i+row_from, 6, 1, 1).setValue("\u2191");
          } else if ( i > j ) {
            sheet.getRange(i+row_from, 1, 1, 1).setBackground("red");
            sheet.getRange(i+row_from, 6, 1, 1).setValue("\u2193");
          } else {
            sheet.getRange(i+row_from, 1, 1, 1).setBackground("white");
            sheet.getRange(i+row_from, 6, 1, 1).setValue("\u2194");
          }
        }
        
        var db_cell = sheet.getRange(i+row_from, db_column, 1, 1);
        
        if( new_data[i][1] > old_data[j][1] ) {
          var change = new_data[i][1] - old_data[j][1];
          if(triggered) { 
            sheet.getRange(i+row_from, change_column, 1, 1).setValue("+" + change);
            db_cell.setValue(db_cell.getValue() + "," + change);
          }
          charting.push([new_data[i][0], change]);
        }
        else {
          if(triggered) { 
            sheet.getRange(i+row_from, change_column, 1, 1).clear();
            db_cell.setValue(db_cell.getValue() + "," + 0);
          }
        }
      }
    }
  }
  Logger.log(charting.length);
  if(typeof triggered === "undefined") {
    if(charting.length < 1 ) {
      var sheet_data = sheet.getSheetValues(3, 2, new_data.length, 4);
      for(var i = 0; i < sheet_data.length; ++i) {
        //Logger.log(typeof parseInt(sheet_data[i][3]) + "::" + parseInt(sheet_data[i][3]));
        if(sheet_data[i][3]) 
          charting.push([sheet_data[i][0], sheet_data[i][3]]);
      }
      Logger.log(charting);
    }
    Logger.log(charting.length);
    var chart_data = Charts.newDataTable()
                     .addColumn(Charts.ColumnType.STRING, "Name")
                     .addColumn(Charts.ColumnType.NUMBER, "Attempts");
    for(var i = 0; i < charting.length; ++i) {
      Logger.log(charting[i]);
      chart_data.addRow(charting[i]);
    }
    chart_data.build();

    var chart = Charts.newBarChart()
       .setTitle('TopCoder Attempts for ' + (new Date()).toDateString())
       .setXAxisTitle('Attempts')
       .setYAxisTitle('Name')
       .setDimensions(1000, 600)
       .setDataTable(chart_data)
       .build();
    
    var app = UiApp.createApplication().setTitle("Topcoder Chart");
    app.setHeight(650);
    app.setWidth(900);
    app.add(chart);
    SpreadsheetApp.getActiveSpreadsheet().show(app);    
  }
  
  return new_data;
}

/**
 * To trigger the calculateRanks function and write to the 'Results' sheet
 *
 */
function triggerCalculateRanks() {
  var data = calculateRanks(true);
  var data_range = SpreadsheetApp.getActiveSpreadsheet()
                   .getSheetByName("Result").getRange(3, 2, data.length, 2);
  Logger.log("YAY");
  Logger.log(data);
  data_range.setValues(data);
}

/**
 * 
 *
 */
function chartScores(logging) {
  logging = logging || false;
  var date_start   = new Date(2013, 7, 7);
  var result_sheet = SpreadsheetApp.getActiveSpreadsheet()
                     .getSheetByName("Result");
  var data = getCount(false)[0];
}

/**
 * Extract data from 3rd row of the data sheets(1-3) to sheet "Garbage Db Sheet"
 * For future purpose, charting
 */
function extractAndCreateDB() {
  var sheet_db = SpreadsheetApp.getActiveSpreadsheet()
              .getSheetByName("garbage_db_sheet");
  var data_sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var user_data = calculateRanks(false);
  Logger.log(data_sheets);
  
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Summary",
    functionName : "myOnOpen"
  }
  /*,{
    name: "Charts",
    functionName : "chartsApp"
  }*/
  ,{ 
    name: "Ranks"
    ,functionName : "calculateRanks"
  }
  ];
  sheet.addMenu("Summary and Ranks", entries);
};
