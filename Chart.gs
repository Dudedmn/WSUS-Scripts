/*
function onOpen() {
  //set up custom menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Chart')
    .addItem('Insert chart...','newChart')
    .addToUi();
};

function newChart() {
  // Generate a chart representing the data in the range of A1:B15.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.BAR)
     .addRange(sheet.getRange('A1:B15'))
     .setPosition(5, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
}
*/