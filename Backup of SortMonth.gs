/*
//@author Daniel Yan
//Needs dynamic selection rather than static
//Currently working for static version (col S-V)
//@date 10/30/2018


//Global variables, alterative is properties service for better privacy
//Function scoping doesn't apply to GoogleScript
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheetName = ss.getActiveSheet().getSheetName();
 var sheet = ss.getActiveSheet(); // gets current selected sheet
 var lastCol = sheet.getLastColumn();
 var lastRow = sheet.getLastRow();
 var data = sheet.getDataRange().getValues();

//Creates a custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SortMonths')
    .addItem('updateMonths','getMonthTotal')
    //Need to add one for Resolved and Unresolved
    .addToUi();
}

//Reference used https://stackoverflow.com/questions/45562955/how-can-i-clear-a-column-in-google-sheets-using-google-apps-script
function clearColumnElement(cellElement){
 //Assumed user input is first column, row, entry
  var firstCell = sheet.getRange(cellElement);
    Logger.log("firstCell value: %s",firstCell.getValue());
  var numRows = (sheet.getLastRow() - firstCell.getRow()) + 1;
  var range = sheet.getRange(firstCell.getRow() + 1, firstCell.getColumn(), numRows);
  range.clear();
}

//Reference used https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

//Reference used https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

//Assumes resolved column is a single String input ********DEFUNCT******
function resolvedCheck(resolvedColumn) {
var i;
//var values = ss.getRange(((resolvedColumn.toUpperCase() + 1).toString + ":" + resolvedColumn.toUpperCase()).toString()).getValues();
//var mLength = values.filter(String).length;
for(i = 2; i < lastRow; i++){
  Logger.log(data[i][13]);
  }
}


function modifyValues(totals, resolved, unresolved, resolveValue) {
  //Adds a value to Total cases per month
  sheet.getRange(totals).setValue(sheet.getRange(totals).getValue()+1);
  //Adds a value to Resolved or Unresolved depending on range value
  (resolveValue === "Yes") ? sheet.getRange(resolved).setValue(sheet.getRange(resolved).getValue()+1)
            : sheet.getRange(unresolved).setValue(sheet.getRange(unresolved).getValue()+1);
            
}

//Assumes A1 noation for all inputs
function setZero(totals, resolved, unresolved) {
if(sheet.getRange(totals).getValue() === "" ) {
  sheet.getRange(totals).setValue(0);
  sheet.getRange(resolved).setValue(0);
  sheet.getRange(unresolved).setValue(0);
  }
}

//Assumes A1 notation for all inputs
function setTable(monthCol, totalCol, resolvedCol, unresolvedCol) {
  sheet.getRange(monthCol).setValue("Months");
  var i;
  for(i = 3; i < 15; i++) {
    //Logger.log("letter to col index: %s", letterToColumn(monthCol.substring(0,1).toUpperCase()));
    //Gets month column dependent on passed on month character in A1 notation
    var range = sheet.getRange(i,letterToColumn(monthCol.substring(0,1).toUpperCase()));
    switch (i) {
      case 3:
      range.setValue("January");
        break;
      case 4:
      range.setValue("February");
        break;
      case 5:
      range.setValue("March");
        break;  
      case 6:
      range.setValue("April");
        break;
      case 7:
      range.setValue("May");
        break;
      case 8:
      range.setValue("June");
        break;
      case 9:
      range.setValue("July");
        break;
      case 10:
      range.setValue("August");
        break;
      case 11:
      range.setValue("September");
        break;
      case 12:
      range.setValue("October");
        break;
      case 13:
      range.setValue("November");
        break;
       case 14:
       range.setValue("December");
        break; 
    }
  }
  sheet.getRange(totalCol).setValue("Totals");
  sheet.getRange(resolvedCol).setValue("Resolved");
  sheet.getRange(unresolvedCol).setValue("Unresolved");
}

//Searches through all column cells of 'Last Status' and totals number based on month from 'Last Contact'
//Clears any prior total value
function getMonthTotal() {
  //User prompt for where total cells go
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter in the character of the column you wish to clear');
  var cellElement  = response.getResponseText();
  while(cellElement === "" || cellElement.toString().length > 1) {
      var attempt = ui.prompt('Enter in the character of the column you wish to clear');
      cellElement = attempt.getResponseText();
  }
  var i = 0;
  Logger.log("cell element: %s",cellElement);
  
 // resolvedCheck("M");
  //Clears next 2 adjacent columns to input character
  //Dynamically made from user prompt input
  while(i < 4) {
    clearColumnElement((cellElement + 1).toString().toUpperCase());
    var index = letterToColumn(cellElement.substring(0,1).toUpperCase());
    index++;
    //Move to next column
    cellElement = columnToLetter(index);
    //Iterate loop
    i++;
   }
  
  //Temporary static setting, needs to be made dynamic
  setTable("S1","T1","U1","V1");

  //sheet.getRange("S1").setValue("Months");
  //sheet.getRange("T1").setValue("Totals");
  //sheet.getRange("U1").setValue("Resolved");
  //sheet.getRange("V1").setValue("Unresolved");

  //Temporary static setting, needs to be made dynamic
  for(i = 3; i < 15; i++) {
  setZero(("T"+i).toString(),("U"+i).toString(),("V"+i).toString());
  }
  
  //Starting row
  for(i = 2; i <= lastRow; i++) {
   //Takes the ith row in column 2
   var range = sheet.getRange(i,2);
   //This is a raw unformatted Date value
   var date  = range.getValue();
   //Skip empty cells
    if (date === "") {}
    else { 
   //Reformat date object to US Time format
    var formattedDate = Utilities.formatDate(new Date(date), "GMT", "MM-dd-yyyy'T'HH:mm:ss'Z'");
    var month = formattedDate.substring(0,2);
    var resolveRange = sheet.getRange(i,13);
    var resolveValue = resolveRange.getValue();
    
    //Parse with base 10
    //Temporary static setting, needs to be made dynamic
      switch(parseInt(month,10)) {
      //January
        case 01:
          modifyValues("T3", "U3", "V3", resolveValue);
          break;
      //February
        case 02:
           modifyValues("T4", "U4", "V4", resolveValue);
          break;
      //March
        case 03:
          modifyValues("T5", "U5", "V5", resolveValue);
          break;
      //April    
        case 04: 
          modifyValues("T6", "U6", "V6", resolveValue);
          break;
      //May    
        case 05:
          modifyValues("T7", "U7", "V7", resolveValue);
          break;
      //June    
        case 06:
          modifyValues("T8", "U8", "V8", resolveValue);
          break;          
      //July    
        case 07:
          modifyValues("T9", "U9", "V9", resolveValue);
          break;
      //August    
        case 08:
          modifyValues("T10", "U10", "V10", resolveValue);
          break;
      //September    
        case 09:
          modifyValues("T11", "U11", "V11", resolveValue);
          break;
      //October    
        case 10:
          modifyValues("T12", "U12", "V12", resolveValue);
          break;
      //November    
        case 11:
          modifyValues("T13", "U13", "V13", resolveValue);
          break;
      //December    
        case 12:
          modifyValues("T14", "U14", "V14", resolveValue);
          break;
        
        }
       }
      }
     }
   

*/

