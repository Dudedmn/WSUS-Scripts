//@author Daniel Yan
//Needs dynamic selection rather than static
//Currently working for static version (col S-W)
//@date 10/30/2018
//References used:
//https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
//https://stackoverflow.com/questions/45562955/how-can-i-clear-a-column-in-google-sheets-using-google-apps-script
//https://developers.google.com/apps-script/reference/spreadsheet/range#getvalues



//Global variables, alterative is properties service for better privacy
//Function scoping doesn't apply to Google App Script
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheetName = ss.getActiveSheet().getSheetName();
 var sheet = ss.getActiveSheet(); // gets current selected sheet
 var lastCol = sheet.getLastColumn();
 var lastRow = sheet.getLastRow();
 var data = sheet.getDataRange().getValues();
 var backgroundColor = "#ff9900";
 
//Dynamic thoughts?
//Get position of row & column based on user input
// cellElement = user inp
//data[cellElement.row][cellElement.col]
//Modify all other cells based on that cellElement input

//Creates a custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SortMonths')
    .addItem('updateMonths','getMonthTotal')
    //Need to add one for Resolved and Unresolved
    .addToUi();
}

//Originator user Pierre-Marie Richard
//Assume A1 input for column
function clearColumnElement(cellElement){
 //Assumed user input is first column, row, entry
  var firstCell = sheet.getRange(cellElement);
    Logger.log("firstCell value: %s",firstCell.getValue());
  var numRows = (sheet.getLastRow() - firstCell.getRow()) + 1;
  var range = sheet.getRange(firstCell.getRow() + 1, firstCell.getColumn(), numRows);
  range.clear();
}

//Originator user AdamL
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

//Originator user AdamL
function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}


function modifyValues(totals, resolved, unresolved, resolveValue) {
  //Adds a value to Total cases per month
  sheet.getRange(totals).setValue(sheet.getRange(totals).getValue()+1);
  //Adds a value to Resolved or Unresolved depending on range value
  (resolveValue === "Yes") ? sheet.getRange(resolved).setValue(sheet.getRange(resolved).getValue()+1)
            : sheet.getRange(unresolved).setValue(sheet.getRange(unresolved).getValue()+1);
            
}

//Assumes A1 notation for all inputs
function setZero(totals, resolved, unresolved) {
if(sheet.getRange(totals).getValue() === "" ) {
  sheet.getRange(totals).setValue(0);
  sheet.getRange(resolved).setValue(0);
  sheet.getRange(unresolved).setValue(0);
  }
}

//Assumes A1 notation for all inputs
function setBackgroundColor(cellElement) {
var start = parseInt(cellElement.substring(1,2),10);
var i; var col;
for( col = 0; col < 5; col++) {
    for (i = start; i < (start + 13); i++) {
    sheet.getRange((cellElement.substring(0,1) + i).toString()).setBackgroundColor(backgroundColor); 
      }
    //Iterate to next columns
    var index = letterToColumn(cellElement.substring(0,1));
    index++;
    cellElement = ((columnToLetter(index) + start).toString());
  }
}

//Assumes A1 notation for all inputs
function setTable(yearCol, monthCol, totalCol, resolvedCol, unresolvedCol) {
  var date = sheet.getRange("B2").getValue();
  var formattedDate = Utilities.formatDate(new Date(date), "GMT", "MM-dd-yyyy'T'HH:mm:ss'Z'");
  var year = formattedDate.substring(6,10);
   Logger.log("setYear: %s", year);
  sheet.getRange(yearCol).setValue("Year").setFontWeight("bold");
  sheet.getRange((yearCol.substring(0,1) + 3).toString()).setValue(year);
  sheet.getRange(monthCol).setValue("Months").setFontWeight("bold");
  sheet.getRange(totalCol).setValue("Totals").setFontWeight("bold");
  sheet.getRange(resolvedCol).setValue("Resolved").setFontWeight("bold");
  sheet.getRange(unresolvedCol).setValue("Unresolved").setFontWeight("bold");
  var i;
  //Responds to row numbers  ******STATIC********
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
}

//Returns type of reference to input
function checkType(reference) {
return typeof reference;
}

function logTest() {
  Logger.log("return test string: %s", checkType("asd"));
  Logger.log("return test int: %s", checkType(2));
  Logger.log("return test float: %s", checkType(2.0));
  Logger.log("return test char : %s", checkType('a'));
  Logger.log("parse value for d: %s", parseInt("d",10));
  Logger.log("parse value for D: %s", parseInt("D",10));

}

//Searches through all column cells of 'Last Status' and totals number based on month from 'Last Contact'
//Clears any prior total value
function getMonthTotal() {
  //User prompt for where total cells go
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter in the cell you wish to construct the data using A1 Notation (e.g. \'s2\')');
  //All column char names are upper case
  var cellElement  = response.getResponseText().toUpperCase();
  var original = cellElement;
  
  //User prompt verification which only accepts A1 Notation, checks if second input is a number or not
  while(cellElement === "" || cellElement.toString().length != 2  || isNaN(cellElement.substring(1,2))) {
      var attempt = ui.prompt('Enter in the cell you wish to construct the data using A1 Notation (e.g. \'s2\')');
      cellElement = attempt.getResponseText().toUpperCase();
  }
  
  
  
  //Variables for looping. 
  var i = 0, colErase = 5, num = parseInt(cellElement.substring(1,2),10);
  Logger.log("cell element: %s",cellElement);
  
  //Clears next specified adjacent columns to input character
  //Dynamically made from user prompt input
  //colErase is number of columns to wipe on the right
  while(i < colErase) {
  clearColumnElement(cellElement);
  //Find the column char
  var index = letterToColumn(cellElement.substring(0,1));
  //Move to next column
  index++;
  //Add num back to get into A1 notation
  cellElement = (columnToLetter(index) + num).toString();
  //Iterate loop
  i++;
  }

  
  //Temporary static setting, needs to be made dynamic ******STATIC********
  setTable("S2","T2","U2","V2","W2");
  
  //Temporary static setting, needs to be made dynamic ******STATIC********
  for(i = 3; i < 15; i++) {
  setZero(("U"+i).toString(),("V"+i).toString(),("W"+i).toString());
  }
  
  //Dynamic background color setting 
  setBackgroundColor(original);
  
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
    //corresponds to column 'M'
    var resolveRange = sheet.getRange(i,letterToColumn("M"));
    var resolveValue = resolveRange.getValue();
    
    //Parse with base 10
    //Temporary static setting, needs to be made dynamic ******STATIC********
      switch(parseInt(month,10)) {
      //January
        case 01:
          modifyValues("U3", "V3", "W3", resolveValue);
          break;
      //February
        case 02:
           modifyValues("U4", "V4", "W4", resolveValue);
          break;
      //March
        case 03:
          modifyValues("U5", "V5", "W5", resolveValue);
          break;
      //April    
        case 04: 
          modifyValues("U6", "V6", "W6", resolveValue);
          break;
      //May    
        case 05:
          modifyValues("U7", "V7", "W7", resolveValue);
          break;
      //June    
        case 06:
          modifyValues("U8", "V8", "W8", resolveValue);
          break;          
      //July    
        case 07:
          modifyValues("U9", "V9", "W9", resolveValue);
          break;
      //August    
        case 08:
          modifyValues("U10", "V10", "W10", resolveValue);
          break;
      //September    
        case 09:
          modifyValues("U11", "V11", "W11", resolveValue);
          break;
      //October    
        case 10:
          modifyValues("U12", "V12", "W12", resolveValue);
          break;
      //November    
        case 11:
          modifyValues("U13", "V13", "W13", resolveValue);
          break;
      //December    
        case 12:
          modifyValues("U14", "V14", "W14", resolveValue);
          break;
        
        }
       }
      }
     }
   



