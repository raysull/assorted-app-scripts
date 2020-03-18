/**
  *Create a cool new menu
**/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.
createMenu('Ray Scripts')
  .addItem('Automatic Headers', 'createHeaders')
  .addItem('Quickly Match Assignment', 'autoPaste')
  .addItem('xF Finder', 'xfFinder')
  .addItem('Assigned Today', 'assignToday')
  .addItem('Remove Duplicates', 'removeDuplicates')
  .addItem('Previously Assigned Date', 'previousAssign')
  .addItem('Identify Heme Cases', 'hemeFinder')
  .addToUi();
  
  function autoPaste() {
  // Quickly Match from assignment sheet.
}
  function removeDuplicates() {
  //Remove duplicates from Orders  
  }
  function createHeaders() {
  //Autocreate Headers
  }
  function xfFinder() {
  }
  function assignToday(){
  }
  function previousAssign(){
  }
  function hemeFinder(){
  }
}

/**
 * Removes duplicate rows from the current sheet.
 */
function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
   if(row[3] == newData[j][3]){
     duplicate = true;
   }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

/** 
  * Create the headers that I have to add into each sheet every time
**/
function createHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Freezes the first row
  sheet.setFrozenRows(1);

  // Set the values we want for headers
  var values = [
    ["DESIRED","HEADER","Values","here"]
  ];

  // Set the range of cells
  var range = sheet.getRange("A1:AC1");

  // Call the setValues method on range and pass in our values
  range.setValues(values).setFontWeight("bold");
}

/** 
  * Automatically paste the formula to determine if a case has been signed out already
**/

function autoPaste() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var data = sheet.getLastRow()-1;
  var cell = sheet.getRange(2,5,data);
      cell.setFormulaR1C1("=Index(Sheet1!C[2]:C[2],MATCH(R[0]C[-1],Sheet1!C[0]:C[0],0))");

}

/** 
  * Automatically paste the formula to determine if a patient is an xF
**/
function xfFinder() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var data = sheet.getLastRow()-1;
 var cell = sheet.getRange(2,8,data);
  cell.setFormulaR1C1('=If(R[0]C[1]="xF.v2","xF"," " )');
}
/** 
* Determine when a previously assigned case was assigned:
Side note, this will feed through and show current assigned date.
**/
function previousAssign() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var data = sheet.getLastRow()-1;
  var cell = sheet.getRange(2,10,data);
      cell.setFormulaR1C1("=IFERROR(Index(Sheet1!C[-2]:C[-2],MATCH(R[0]C[-6],Sheet1!C[-5]:C[-5],0)))");
 var prev = cell.getValues();
      cell.setValues(prev).setNumberFormat('m/d/yy');
}
/** 
* Assign the case to today
**/
function assignToday() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var data = sheet.getLastRow()-1;
  var cell = sheet.getRange(2,11,data);
  var today = new Date();
  var date = (today.getMonth()+1)+'/'+today.getDate()+'/'+today.getFullYear();
      cell.setValue(date).setNumberFormat('m/d/yy');
      
}
/** 
* Identify Heme Cases
**/
function hemeFinder() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var data = sheet.getLastRow()-1;
 var cell = sheet.getRange(2,17,data);
  cell.setFormulaR1C1('=IFERROR(IF(MATCH(R[0]C[-1],Heme_diagnoses!C[-16]:C[-16],0),"Heme",""),"")');
}
