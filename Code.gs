/*
google.charts.load('current', {packages: ['corechart', 'bar']});
google.charts.setOnLoadCallback(drawAxisTickColors);

//add menu to google sheet
function onOpen() {
  //set up custom menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Draw Chart')
    .addItem('Insert chart...','drawAxisTickColors')
    .addToUi();
};

function drawAxisTickColors() {
      var data = new google.visualization.DataTable();
      data.addColumn('timeofday', 'Time of Day');
      data.addColumn('number', 'Motivation Level');
      data.addColumn('number', 'Energy Level');

      data.addRows([
        [{v: [8, 0, 0], f: '8 am'}, 1, .25],
        [{v: [9, 0, 0], f: '9 am'}, 2, .5],
        [{v: [10, 0, 0], f:'10 am'}, 3, 1],
        [{v: [11, 0, 0], f: '11 am'}, 4, 2.25],
        [{v: [12, 0, 0], f: '12 pm'}, 5, 2.25],
        [{v: [13, 0, 0], f: '1 pm'}, 6, 3],
        [{v: [14, 0, 0], f: '2 pm'}, 7, 4],
        [{v: [15, 0, 0], f: '3 pm'}, 8, 5.25],
        [{v: [16, 0, 0], f: '4 pm'}, 9, 7.5],
        [{v: [17, 0, 0], f: '5 pm'}, 10, 10],
      ]);

      var options = {
        title: 'Motivation and Energy Level Throughout the Day',
        focusTarget: 'category',
        hAxis: {
          title: 'Time of Day',
          format: 'h:mm a',
          viewWindow: {
            min: [7, 30, 0],
            max: [17, 30, 0]
          },
          textStyle: {
            fontSize: 14,
            color: '#053061',
            bold: true,
            italic: false
          },
          titleTextStyle: {
            fontSize: 18,
            color: '#053061',
            bold: true,
            italic: false
          }
        },
        vAxis: {
          title: 'Rating (scale of 1-10)',
          textStyle: {
            fontSize: 18,
            color: '#67001f',
            bold: false,
            italic: false
          },
          titleTextStyle: {
            fontSize: 18,
            color: '#67001f',
            bold: true,
            italic: false
          }
        }
      };

      var chart = new google.visualization.ColumnChart(document.getElementById('chart_div'));
      chart.draw(data, options);
    }
 
// function to create waterfall chart
function waterfallChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  var data = range.getValues();
  
  var newData = [['Label','Endpoints','Base','Postive Cols above','Positive Cols below',
                  'Negative Cols above','Negative Cols below']];
  
  var tempTotal = 0;
  var tempTotalPrior = 0;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  for (i = 1; i < data.length; i++) {
    
    // running totals
    tempTotalPrior = tempTotal;  // assign previous total to new variable to keep track of it
    tempTotal += data[i][1];  // add up total of values so far
    
    // Endpoints
    if (i == 1 || i == data.length - 1) {
      newData.push([data[i][0],data[i][1],0,'','','','']);
    }
    
    // Non-endpoints
    else {
      
      // Base values
      var baseVal = Math.max(0,Math.min(tempTotal,tempTotalPrior)) + Math.min(0,Math.max(tempTotal,tempTotalPrior));
      
      // calculate minimum of running total and current value
      var val1 = Math.min(tempTotal,data[i][1]);
      
      // calculate maximum of running total and current value
      var val2 = Math.max(tempTotal,data[i][1]);
      
      // Postive Cols above
      // if val1 is negative, set to 0, otherwise take val1, which is min of running total and current value
      var posValAbove = Math.max(0,val1);
      
      // Postive Cols below
      // subtract current value from Positive Col Value to catch any part of column below 0. If a positive value, set to 0 by using minimum
      var posValBelow = Math.min(posValAbove - data[i][1],0);
      
      // Negative Cols below
      // if val2 is positive, set to 0, otherwise take val2, which is the max of running total and current value
      var negValBelow = Math.min(0,val2);
      
      // Negative Cols above
      // subtract current value from Negative Col Value to catch any part of column above 0. If a negative value, set to 0 by using maximum
      var negValAbove = Math.max(negValBelow - data[i][1],0);
       
      // push all new datapoints into newData array
      newData.push([data[i][0],0,baseVal,posValAbove,posValBelow,negValAbove,negValBelow]);
    }
    
  }
  
  // paste the new data into sheet
  sheet.getRange(lastRow - data.length + 1, lastCol + 2, data.length, newData[0].length).setValues(newData);
 
  // get the new data for the chart
  var chartData = sheet.getRange(lastRow - data.length + 1, lastCol + 2, data.length, newData[0].length);
  
  // make the new waterfall chart
  sheet.insertChart(
    sheet.newChart()
    .addRange(chartData)
    .setChartType(Charts.ChartType.COLUMN)
    .asColumnChart()
    .setStacked()
    .setColors(['grey','none','green','green','red','red'])
    .setOption('title','Waterfall Chart')
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(lastRow - data.length + 4,lastCol + 4,0,0)
    .build()
  );
}
*/